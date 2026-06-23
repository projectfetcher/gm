import os
import re
import csv
import sys
import time
import json
import base64
import hashlib
import logging
from datetime import datetime, timedelta
from urllib.parse import urljoin, urlparse, urlsplit, urlunsplit, parse_qsl, urlencode

import requests
from bs4 import BeautifulSoup

# Optional: load secrets from a local .env file if python-dotenv is installed.
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Optional heavy deps used for Excel export only.
try:
    import pandas as pd
    import openpyxl
    _XLSX_AVAILABLE = True
except ImportError:
    _XLSX_AVAILABLE = False

# Optional heavy deps used for paraphrase quality gating.
try:
    import language_tool_python
    from sentence_transformers import SentenceTransformer, util as st_util
    _NLP_AVAILABLE = True
except ImportError:
    _NLP_AVAILABLE = False

# =============================================================================
#  CONFIG
# =============================================================================
#
#  SOURCES
#  -------
#  Two URLs were requested:
#
#    * https://gamjobs.com/                 -> REAL source (scraped here). A standard
#                                              JobMonster WordPress board: a paginated
#                                              /jobs/ archive of listing cards, each
#                                              linking to a /jobs/<slug>/ detail page.
#    * https://gambia.gov.gm/vacancies/     -> NOT a listing source. It is the official
#                                              government "gateway" page; its only job
#                                              content is two buttons linking out to the
#                                              PSC site and to GamJobs. The scraper still
#                                              *reads* it (resolve_gov_gateway) to confirm
#                                              GamJobs as the live source and to surface
#                                              PSC as the other official source — but the
#                                              jobs themselves come from GamJobs.
#
#  APPLY RULE (hard, network-wide)
#  -------------------------------
#  A job only posts if it exposes a PUBLIC apply path: an email or an external apply
#  URL found in its "How to Apply" text. GamJobs' on-page "Apply for this job" button
#  just opens a login/registration modal, so it is NEVER treated as a valid apply
#  destination. Jobs without a public email/URL are written to the flagged CSV.
#  REQUIRE_PUBLIC_APPLY (default "1"/on) enforces this; set to "0" to post everything.
# =============================================================================

BASE_URL  = "https://gamjobs.com"
JOBS_URL  = os.environ.get("GAMJOBS_JOBS_URL", "https://gamjobs.com/jobs/")

# Official government gateway page (no listings of its own — links to GamJobs + PSC).
GOV_GATEWAY_URL = os.environ.get("GAMBIA_GOV_VACANCIES_URL",
                                 "https://gambia.gov.gm/vacancies/")
# If the gateway page is reachable, follow the GamJobs link it advertises.
FOLLOW_GOV_GATEWAY = os.environ.get("FOLLOW_GOV_GATEWAY", "1") != "0"

# Enforce the public-apply-only rule (email or external URL required to post).
REQUIRE_PUBLIC_APPLY = os.environ.get("REQUIRE_PUBLIC_APPLY", "1") != "0"

REQUEST_DELAY   = float(os.environ.get("REQUEST_DELAY", "1.0"))
MAX_JOBS        = int(os.environ.get("MAX_JOBS", "0"))     # 0 = unlimited
MAX_PAGES       = int(os.environ.get("MAX_PAGES", "15"))   # archive pagination cap
REQUEST_TIMEOUT = int(os.environ.get("REQUEST_TIMEOUT", "25"))

OUTPUT_FILE        = "gamjobs_gambia_jobs.xlsx"
PROCESSED_IDS_FILE = "gamjobs_gambia_processed.csv"
FLAGGED_FILE       = "gamjobs_gambia_flagged.csv"

# CSV column names — defined once so _init_tracker, load, and upsert all agree.
_TRACKER_FIELDS = ["Job ID", "Job URL", "Job Title", "Company Name",
                   "Status", "Timestamp", "WP ID"]

_FLAGGED_FIELDS = ["Source", "Title", "Company", "Location", "Salary",
                   "Deadline", "Reason", "Apply Note", "Job URL", "Timestamp"]

# ── WordPress ────────────────────────────────────────────────────────────────
WP_URL      = os.environ.get("WP_BASE_URL", "")
WP_USER     = os.environ.get("WP_USERNAME", "")
WP_PASSWORD = os.environ.get("WP_APP_PASSWORD", "")
WP_BASE      = WP_URL.rstrip("/")
WP_JOBS_URL  = f"{WP_BASE}/job-listings"
WP_MEDIA_URL = f"{WP_BASE}/media"

# ── Mistral ──────────────────────────────────────────────────────────────────
MISTRAL_API_KEY = os.environ.get("MISTRAL_API_KEY", "")
MISTRAL_MODEL   = "mistral-small-latest"
MISTRAL_URL     = "https://api.mistral.ai/v1/chat/completions"

ENABLE_PARAPHRASE = True

# ── Startup warnings ─────────────────────────────────────────────────────────
for _var, _val, _feature in [
    ("MISTRAL_API_KEY", MISTRAL_API_KEY, "paraphrasing"),
    ("WP_USERNAME",     WP_USER,         "WordPress posting"),
    ("WP_APP_PASSWORD", WP_PASSWORD,     "WordPress posting"),
]:
    if not _val:
        logging.getLogger(__name__).warning(
            f"Environment variable {_var} is not set — {_feature} will be disabled/skipped."
        )

JOB_TYPE_MAPPING = {
    "full-time": "full-time", "full time": "full-time",
    "part-time": "part-time", "part time": "part-time",
    "contract":  "contract",  "temporary": "temporary",
    "internship":"internship","freelance": "freelance",
    "volunteer": "volunteer", "permanent": "full-time",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml",
    "Accept-Charset": "utf-8",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)

# Known Gambian towns/cities, used as a fallback to pull a location from free text.
GAMBIA_LOCATIONS = [
    "Banjul", "Abuko", "Brikama", "Bakau", "Serrekunda", "Serekunda",
    "Kanifing", "Old Yundum", "New Yundum", "Yundum", "Basse", "Farafenni",
    "Soma", "Kerewan", "Mansakonko", "Janjangbureh", "Kuntaur", "Sukuta",
    "Gunjur", "Lamin", "Bwiam", "Kanilai", "Essau", "Barra", "Kotu", "Fajara",
    "Bakoteh", "Bijilo", "Brusubi", "Brufut", "Kololi", "Tanji", "Tallinding",
    "Latrikunda", "Manjai", "Jabang", "Kairaba Avenue",
]
# GamJobs uses "The Gambia" as the country-level catch-all location.
DEFAULT_LOCATION = os.environ.get("GAMJOBS_DEFAULT_LOCATION", "The Gambia")

# Hosts/paths that are never a real external apply destination.
_NON_APPLY_HOST_SUBSTR = (
    "gamjobs.com", "facebook.", "twitter.", "x.com", "linkedin.",
    "instagram.", "wa.me", "whatsapp", "t.me", "telegram",
    "plus.google", "pinterest.", "youtube.",
)
_NON_APPLY_PATH_SUBSTR = (
    "/member-2", "action=login", "mode=register", "#share", "/share",
    "/wp-login", "/cart", "/checkout",
)
# Emails belonging to the board itself are never a real apply address — these
# appear in the topbar/footer ("info@gamjobs.com") and must not be posted as the
# place to apply.
_NON_APPLY_EMAIL_DOMAINS = ("gamjobs.com",)

def _is_real_apply_email(email: str) -> bool:
    if not email or "@" not in email:
        return False
    dom = email.rsplit("@", 1)[-1].lower()
    return not any(dom == d or dom.endswith("." + d) for d in _NON_APPLY_EMAIL_DOMAINS)

# =============================================================================
#  LOGGING / COLOUR
# =============================================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log_ = logging.getLogger(__name__)

_USE_COLOUR = sys.stdout.isatty()

def _c(code, text):
    return f"\033[{code}m{text}\033[0m" if _USE_COLOUR else text

C_HEADER  = lambda t: _c("1;36",  t)
C_LABEL   = lambda t: _c("1;33",  t)
C_VALUE   = lambda t: _c("97",    t)
C_DIM     = lambda t: _c("2",     t)
C_GREEN   = lambda t: _c("1;32",  t)
C_RED     = lambda t: _c("1;31",  t)
C_BLUE    = lambda t: _c("1;34",  t)
C_DIVIDER = lambda: _c("2", "─" * 80)

def log(msg):
    print(msg, flush=True)

EMAIL_PATTERN = re.compile(r"[A-Za-z0-9.+_-]+@[A-Za-z0-9-]+\.[A-Za-z0-9.-]+")
URL_PATTERN   = re.compile(r"https?://[^\s)>\"']+", re.I)

TRACKING_PARAM_PREFIXES = ("utm_",)
TRACKING_PARAM_EXACT = {
    "fbclid", "gclid", "msclkid", "mc_cid", "mc_eid", "ref", "referrer",
}

MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11,
    "december": 12, "jan": 1, "feb": 2, "mar": 3, "apr": 4, "jun": 6, "jul": 7,
    "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
}

# Ordinal text date e.g. "30th June 2026" / "7th July, 2026".
TEXT_DATE_RE = re.compile(
    r"(\d{1,2})\s*(?:st|nd|rd|th)?\s+([A-Za-z]+)\s*[.,]?\s*(\d{4})", re.I
)
# Numeric DD/MM/YYYY or DD-MM-YYYY (GamJobs meta line + Job Overview).
DMY_DATE_RE = re.compile(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b")

# Labels inside the JobMonster "Job Overview" box.
DEADLINE_LABELS = ("application deadline", "closing date", "deadline",
                   "expiry date", "expires")

# Body headings that introduce the application instructions. Matched against a
# *stripped, short* line (see _is_apply_heading_line) so it never trips on
# 'Application Deadline:' / 'Application Format' or a mid-sentence 'to apply'.
_APPLY_HEAD_PHRASES = re.compile(
    r"^(?:how\s*(?:and|&)\s*deadline\s*to\s*apply|how\s*to\s*apply(?:\s*(?:and|&)\s*deadline)?|"
    r"how\s*to\s*submit|to\s*apply|application\s*(?:and|&)\s*deadline|"
    r"mode\s*of\s*application|method\s*of\s*application|"
    r"application\s*(?:procedure|process|instructions?|method|guidelines?)|"
    r"submission\s*of\s*applications?|deadline\s*(?:and|&)?\s*(?:how\s*)?to\s*apply)\b",
    re.I,
)

# Boilerplate that marks the end of usable post content on a detail page.
_BODY_CUT_MARKERS = [
    "related jobs", "leave your thoughts", "you must be logged in",
    "email me jobs like these", "send to a friend", "company information",
    "leave a reply", "post a comment",
]
# Standalone UI lines to drop from the description.
_BODY_DROP_LINES = {
    "apply for this job", "save", "share", "share:", "bookmark job",
    "quick view", "send to friend", "send to a friend", "clear all",
    "filter", "view more",
}

# =============================================================================
#  TEXT CLEANUP / SANITIZATION
# =============================================================================

_MOJIBAKE = [
    ("Â", ""), ("â€™", "'"), ("â€œ", '"'), ("â€\x9d", '"'), ("â€", '"'),
    ("â€¢", "•"), ("â„¢", "™"), ("\u00a0", " "), ("\u200b", ""), ("\ufeff", ""),
]

def _fix_mojibake(text: str) -> str:
    for pattern, replacement in _MOJIBAKE:
        text = text.replace(pattern, replacement)
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)
    return text

def sanitize_text(text, is_url=False) -> str:
    if not isinstance(text, str):
        text = str(text) if (text is not None and str(text) not in ("nan", "None", "NaN")) else ""
    text = text.strip()
    if text in ("nan", "None", "NaN", "", "N/A", "n/a", "NA", "na"):
        return ""
    text = _fix_mojibake(text)
    if is_url:
        return re.sub(r"[ \t\r\n\f\v]+", " ", text).strip()
    text = re.sub(r"#+\s*", "", text)
    text = re.sub(r"\*\*", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

def clean_text(el):
    if el is None:
        return ""
    return re.sub(r"\s+", " ", el.get_text(" ", strip=True)).strip()

def extract_email(text):
    if not text:
        return ""
    m = EMAIL_PATTERN.search(text)
    return m.group(0) if m else ""

def strip_tracking_params(url):
    if not url:
        return url
    parts = urlsplit(url)
    if not parts.query:
        return url
    kept = []
    for key, value in parse_qsl(parts.query, keep_blank_values=True):
        key_lower = key.lower()
        if key_lower.startswith(TRACKING_PARAM_PREFIXES) or key_lower in TRACKING_PARAM_EXACT:
            continue
        kept.append((key, value))
    new_query = urlencode(kept)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, parts.fragment))

# =============================================================================
#  BASIC HTTP / PARSING HELPERS
# =============================================================================

def get_soup(url):
    resp = SESSION.get(url, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    resp.encoding = resp.encoding or "utf-8"
    try:
        return BeautifulSoup(resp.text, "lxml")
    except Exception:
        return BeautifulSoup(resp.text, "html.parser")

def slugify(text, maxlen=80):
    s = re.sub(r"[^a-z0-9]+", "-", (text or "").lower()).strip("-")
    return s[:maxlen] or "job"

def html_block_to_text(el) -> str:
    """
    Convert a BeautifulSoup element to readable plain text, preserving line
    breaks for block-level tags and turning <li> into bullet lines. The block is
    mutated in place — only ever call this on a throwaway/per-job element.
    """
    if el is None:
        return ""
    for br in el.find_all("br"):
        br.replace_with("\n")
    for li in el.find_all("li"):
        txt = li.get_text(" ", strip=True)
        li.replace_with("\n• " + txt + "\n")
    for tag in el.find_all(["p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "tr"]):
        tag.insert_before("\n")
        tag.insert_after("\n")
    text = el.get_text("\n")
    text = _fix_mojibake(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

# =============================================================================
#  DATE / FIELD EXTRACTORS
# =============================================================================

def dmy_dates(text: str) -> list:
    """Return ISO dates parsed from DD/MM/YYYY (or DD-MM-YYYY), in order."""
    out = []
    for d, m, y in DMY_DATE_RE.findall(text or ""):
        try:
            out.append(datetime(int(y), int(m), int(d)).strftime("%Y-%m-%d"))
        except ValueError:
            pass
    return out

def text_dates(text: str) -> list:
    """Return ISO dates parsed from ordinal text form ('30th June 2026'), in order."""
    out = []
    for d, mon, y in TEXT_DATE_RE.findall(text or ""):
        month = MONTHS.get(mon.lower())
        if not month:
            continue
        try:
            out.append(datetime(int(y), month, int(d)).strftime("%Y-%m-%d"))
        except ValueError:
            pass
    return out

def parse_any_date(text: str) -> str:
    """Best single date from a label value (prefers numeric, then text)."""
    ds = dmy_dates(text)
    if ds:
        return ds[-1]
    ts = text_dates(text)
    return ts[-1] if ts else ""

def clean_title(raw: str) -> str:
    """Strip the trailing ' 1442 views' / 'views' counter from a detail H1."""
    t = sanitize_text(raw)
    t = re.sub(r"\s*\d[\d,]*\s*views?\s*$", "", t, flags=re.I)
    t = re.sub(r"\s*views?\s*$", "", t, flags=re.I)
    return t.strip()

def map_job_type(raw: str) -> str:
    key = (raw or "").lower().strip()
    return JOB_TYPE_MAPPING.get(key, "full-time")

def pick_location(locations: list) -> str:
    """Prefer a specific town over the country-level 'The Gambia' catch-all."""
    specific = [l for l in locations if l and l.strip().lower() not in ("the gambia", "gambia")]
    if specific:
        return specific[0].strip()
    if locations:
        return locations[0].strip()
    return DEFAULT_LOCATION

def location_from_text(text: str) -> str:
    if text:
        for town in GAMBIA_LOCATIONS:
            if re.search(rf"\b{re.escape(town)}\b", text, re.I):
                return "Serrekunda" if town.lower() == "serekunda" else town
    return DEFAULT_LOCATION

def extract_experience(qual_text: str) -> str:
    if not qual_text:
        return ""
    m = re.search(r"(?:at least|minimum(?: of)?)\s+\d+\s+years?[^.\n;]*", qual_text, re.I)
    if m:
        return m.group(0).strip().rstrip(".")
    m = re.search(r"\b\d+\s+years?[^.\n;]*experience", qual_text, re.I)
    if m:
        return m.group(0).strip().rstrip(".")
    return ""

def extract_salary(text: str) -> str:
    """Best-effort salary. GamJobs rarely lists a figure; returns '' if none."""
    if not text:
        return ""
    m = re.search(r"(?:GMD|D|GMD\s|D\s|₵)\s*([0-9]{1,3}(?:,\s?[0-9]{3})+(?:\.[0-9]+)?)", text)
    if m:
        amt = re.sub(r"\s+", "", m.group(1))
        return f"GMD {amt}"
    m = re.search(r"\b(?:salary|remuneration)\b[^.\n]{0,80}", text, re.I)
    if m and re.search(r"\d", m.group(0)):
        return m.group(0).strip().rstrip(".")
    return ""

# =============================================================================
#  NLP TOOLS (lazy init, optional)
# =============================================================================

_grammar_tool = None
_sim_model    = None

def _get_grammar_tool():
    global _grammar_tool
    if _grammar_tool is None and _NLP_AVAILABLE:
        try:
            _grammar_tool = language_tool_python.LanguageTool(
                "en-US", remote_server="https://api.languagetool.org")
        except Exception as e:
            log_.warning(f"LanguageTool init failed: {e}")
    return _grammar_tool

def _get_sim_model():
    global _sim_model
    if _sim_model is None and _NLP_AVAILABLE:
        try:
            _sim_model = SentenceTransformer("all-MiniLM-L6-v2", device="cpu")
        except Exception as e:
            log_.warning(f"SentenceTransformer init failed: {e}")
    return _sim_model

def grammar_correct(text: str) -> str:
    tool = _get_grammar_tool()
    if tool:
        try:
            return language_tool_python.utils.correct(text, tool.check(text))
        except Exception:
            pass
    return text

def similarity_score(a: str, b: str) -> float:
    model = _get_sim_model()
    if model:
        try:
            emb = model.encode([a, b], convert_to_tensor=True)
            return float(st_util.pytorch_cos_sim(emb[0], emb[1]))
        except Exception:
            pass
    def tokens(s):
        return set(re.sub(r"[^a-z0-9 ]", " ", s.lower()).split())
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / max(len(ta), len(tb))

def clean_output(text: str) -> str:
    text = _fix_mojibake(text)
    for pat in [r"\[/?INST\]", r"</?s>",
                r"(?i)(rewritten?|rephrased?|output|paraphrase[d]?)[:\s]+",
                r"\*\*", r"###", r"---"]:
        text = re.sub(pat, "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return grammar_correct(text.strip())

# =============================================================================
#  MISTRAL API
# =============================================================================

def mistral_generate(prompt: str, max_tokens: int = 400, temperature: float = 0.7) -> str:
    if not MISTRAL_API_KEY:
        log_.warning("MISTRAL_API_KEY not set — skipping paraphrase")
        return ""
    try:
        response = requests.post(
            MISTRAL_URL,
            headers={
                "Authorization": f"Bearer {MISTRAL_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": MISTRAL_MODEL,
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": max_tokens,
                "temperature": temperature,
            },
            timeout=30,
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        log_.error(f"Mistral API error: {e}")
        return ""

# =============================================================================
#  PARAPHRASE FUNCTIONS
# =============================================================================

def _print_wrapped(text: str, prefix: str = "   ", width: int = 100):
    words = text.split()
    line  = []
    for w in words:
        line.append(w)
        if len(" ".join(line)) >= width:
            print(f"{prefix}{' '.join(line)}")
            line = []
    if line:
        print(f"{prefix}{' '.join(line)}")

def paraphrase_title(title: str) -> str:
    if not ENABLE_PARAPHRASE:
        return title
    clean = sanitize_text(title)
    if not clean:
        return title

    print(f"\n ┌─ TITLE PARAPHRASE {'─'*45}")
    print(f" │ Original : \"{clean}\"")
    print(f" │ {'─'*60}")

    best_result = None
    best_sim    = 0.0

    for attempt in range(4):
        temp = round(0.68 + attempt * 0.06, 2)
        print(f" │ Attempt {attempt+1} (temp={temp}):")

        prompt = (
            f"Rewrite this job title professionally using different words. "
            f"Output ONLY the rewritten title, nothing else. "
            f"Keep it between 4 and 12 words.\n\nJob title: {clean}"
        )

        raw    = mistral_generate(prompt, max_tokens=50, temperature=temp)
        result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")

        wc     = len(result.split()) if result else 0
        sim    = similarity_score(clean, result) if result else 0.0
        is_dup = result.lower().strip() == clean.lower().strip()

        print(f" │    Output  : \"{result}\"")
        print(f" │    Words   : {wc} | Similarity: {sim:.3f} | Duplicate: {'Yes' if is_dup else 'No'}")

        valid = bool(result) and 4 <= wc <= 14 and sim >= 0.55 and not is_dup

        if not valid:
            reasons = []
            if not result:  reasons.append("empty output")
            if wc < 4:      reasons.append(f"too short ({wc} words, min=4)")
            if wc > 14:     reasons.append(f"too long ({wc} words, max=14)")
            if sim < 0.55:  reasons.append(f"sim={sim:.3f} < 0.55")
            if is_dup:      reasons.append("identical to original")
            print(f" │    -> REJECTED — {', '.join(reasons)}")
        else:
            if sim > best_sim:
                best_sim    = sim
                best_result = result
                print(f" │    -> ACCEPTED — new best candidate (sim={sim:.3f})")
            else:
                print(f" │    -> VALID but not better than current best (best sim={best_sim:.3f})")

        print(f" │ {'─'*60}")
        time.sleep(1)

    if best_result:
        print(f" │ FINAL SELECTED : \"{best_result}\"")
        print(f" │    Similarity  : {best_sim:.3f}")
        print(f" └{'─'*65}")
        return best_result
    else:
        print(f" │ No valid paraphrase found -> Keeping original: \"{clean}\"")
        print(f" └{'─'*65}")
        return clean

def paraphrase_description(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text

    paragraphs  = [p.strip() for p in re.split(r"\n+", clean) if p.strip()]
    if not paragraphs:
        paragraphs = [clean]
    rewritten   = []
    success_count = 0

    print(f"\n ┌─ DESCRIPTION PARAPHRASE ({len(paragraphs)} paragraph(s)) {'─'*15}")

    for i, para in enumerate(paragraphs):
        orig_wc = len(para.split())

        print(f"\n │ ┌─ Paragraph {i+1}/{len(paragraphs)} {'─'*50}")
        print(f" │ │ ORIGINAL ({orig_wc} words):")
        _print_wrapped(para, prefix=" │ │    ")
        print(f" │ │ {'─'*60}")

        # Very short fragments (section labels, single bullets) — keep as-is.
        if orig_wc < 8:
            print(f" │ │ (kept — too short to paraphrase safely)")
            rewritten.append(para)
            print(f" │ └{'─'*62}")
            continue

        prompt = (
            f"Rewrite this job description paragraph professionally. "
            f"Keep ALL facts, requirements, and responsibilities. "
            f"Use different sentence structure and vocabulary. "
            f"Output ONLY the rewritten paragraph — no labels, no explanation.\n\n"
            f"Original:\n{para}"
        )

        best_result = None
        best_sim    = 0.0
        accepted_text = None

        for attempt in range(3):
            temp = round(0.65 + attempt * 0.08, 2)
            print(f" │ │ Attempt {attempt+1}/3 (temp={temp}):")

            raw    = mistral_generate(prompt, max_tokens=500, temperature=temp)
            result = clean_output(raw).strip()

            rw  = len(result.split()) if result else 0
            sim = similarity_score(para, result) if result and rw >= 5 else 0.0

            if result:
                print(f" │ │    Paraphrased ({rw} words, sim={sim:.3f}):")
                _print_wrapped(result, prefix=" │ │       ")
            else:
                print(f" │ │    Paraphrased : (no output from model)")

            valid = bool(result) and rw >= 8 and sim >= 0.48

            if not valid:
                reasons = []
                if not result: reasons.append("empty output")
                if rw < 8:     reasons.append(f"too short ({rw} words, min=8)")
                if sim < 0.48: reasons.append(f"sim={sim:.3f} < 0.48")
                print(f" │ │    -> REJECTED — {', '.join(reasons)}")
                if result and sim > best_sim:
                    best_sim    = sim
                    best_result = result
                    print(f" │ │       (stored as best fallback, sim={sim:.3f})")
            else:
                print(f" │ │    -> ACCEPTED on attempt {attempt+1}")
                rewritten.append(result)
                success_count += 1
                accepted_text = result
                break

            print(f" │ │ {'─'*60}")
            time.sleep(1)

        if accepted_text is None:
            print(f" │ │ {'─'*60}")
            if best_result and best_sim >= 0.40:
                print(f" │ │ FALLBACK — Using best attempt (sim={best_sim:.3f}):")
                _print_wrapped(best_result, prefix=" │ │    ")
                rewritten.append(best_result)
                success_count += 1
            else:
                print(f" │ │ KEPT ORIGINAL — no acceptable paraphrase (best sim={best_sim:.3f})")
                rewritten.append(para)

        print(f" │ └{'─'*62}")

    print(f"\n │ SUMMARY: {success_count}/{len(paragraphs)} paragraphs successfully paraphrased")
    print(f" └{'─'*80}\n")

    return "\n\n".join(rewritten)

# =============================================================================
#  DUPLICATE TRACKER — pure stdlib csv, NO pandas dependency
# =============================================================================

def _init_tracker():
    if not os.path.exists(PROCESSED_IDS_FILE):
        try:
            with open(PROCESSED_IDS_FILE, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow(_TRACKER_FIELDS)
            log_.info(f"Tracker file created: {PROCESSED_IDS_FILE}")
        except Exception as e:
            log_.error(f"Could not create tracker file {PROCESSED_IDS_FILE}: {e}")

def load_processed_ids() -> tuple:
    _init_tracker()
    ids, urls = set(), set()
    try:
        with open(PROCESSED_IDS_FILE, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("Job ID"):
                    ids.add(row["Job ID"].strip())
                if row.get("Job URL"):
                    urls.add(row["Job URL"].strip())
    except Exception as e:
        log_.error(f"Could not read tracker file: {e}")
    return ids, urls

def _upsert_row(job_id: str, updates: dict):
    _init_tracker()
    rows = []
    try:
        with open(PROCESSED_IDS_FILE, newline="", encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
    except Exception as e:
        log_.error(f"Tracker read error: {e}")
        rows = []

    found = False
    for row in rows:
        if row.get("Job ID", "").strip() == str(job_id):
            row.update(updates)
            row["Timestamp"] = datetime.now().isoformat()
            found = True
            break

    if not found:
        new_row = {k: "" for k in _TRACKER_FIELDS}
        new_row["Job ID"]    = str(job_id)
        new_row["Timestamp"] = datetime.now().isoformat()
        new_row.update(updates)
        rows.append(new_row)

    try:
        with open(PROCESSED_IDS_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=_TRACKER_FIELDS, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
    except Exception as e:
        log_.error(f"Tracker write error: {e}")

def make_job_id(job_url: str, title: str = "", company: str = "") -> str:
    if job_url:
        return hashlib.md5(job_url.encode()).hexdigest()[:16]
    seed = f"{title}{company}"
    return hashlib.md5(seed.encode()).hexdigest()[:16]

def mark_scraped(job_id, job_url, title, company):
    log_.info(f"Tracker -> scraped: {job_id} | {title}")
    _upsert_row(job_id, {
        "Job URL":      job_url,
        "Job Title":    title,
        "Company Name": company,
        "Status":       "scraped",
        "WP ID":        "",
    })

def mark_paraphrased(job_id):
    _upsert_row(job_id, {"Status": "paraphrased"})

def mark_posted(job_id, wp_id, wp_url):
    _upsert_row(job_id, {"Status": "posted", "WP ID": str(wp_id)})

def mark_failed(job_id, reason):
    _upsert_row(job_id, {"Status": f"failed|{reason}"})

# =============================================================================
#  FLAGGED CSV (non-qualifying / login-only apply)
# =============================================================================

def _init_flagged():
    if not os.path.exists(FLAGGED_FILE):
        try:
            with open(FLAGGED_FILE, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow(_FLAGGED_FIELDS)
        except Exception as e:
            log_.error(f"Could not create flagged file {FLAGGED_FILE}: {e}")

def write_flagged(raw_job: dict, reason: str, apply_note: str):
    _init_flagged()
    try:
        with open(FLAGGED_FILE, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                "GamJobs",
                raw_job.get("title", ""),
                raw_job.get("company_name", ""),
                raw_job.get("location", ""),
                raw_job.get("salary", ""),
                raw_job.get("deadline", ""),
                reason,
                apply_note,
                raw_job.get("job_url", ""),
                datetime.now().isoformat(),
            ])
    except Exception as e:
        log_.error(f"Flagged write error: {e}")

# =============================================================================
#  WORDPRESS POSTING
# =============================================================================

def _wp_auth_headers() -> dict:
    token = base64.b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()
    return {"Authorization": f"Basic {token}", "Content-Type": "application/json"}

def get_or_create_term(taxonomy_url: str, name: str):
    if not name or not name.strip():
        return None
    slug = re.sub(r"[^a-z0-9-]", "-", name.lower().strip())
    h = _wp_auth_headers()
    try:
        r = requests.get(f"{taxonomy_url}?slug={slug}", headers=h, timeout=10, verify=False)
        terms = r.json()
        if isinstance(terms, list) and terms:
            return terms[0]["id"]
    except Exception:
        pass
    try:
        r = requests.post(taxonomy_url, json={"name": name, "slug": slug},
                          headers=h, auth=(WP_USER, WP_PASSWORD), timeout=10, verify=False)
        return r.json().get("id")
    except Exception as e:
        log_.error(f"Term create error '{name}': {e}")
        return None

def post_job_to_wordpress(job: dict) -> tuple:
    if not WP_USER or not WP_PASSWORD:
        log_.warning("WP_USERNAME / WP_APP_PASSWORD not set — skipping WordPress post")
        return None, None

    h = _wp_auth_headers()

    title       = sanitize_text(job.get("jobTitle", ""))
    description = sanitize_text(job.get("jobDescription", ""))
    if not title or not description:
        return None, None

    slug = re.sub(r"[^a-z0-9-]", "-", title.lower())[:80]
    try:
        r = requests.get(f"{WP_JOBS_URL}?slug={slug}", headers=h, timeout=10, verify=False)
        posts = r.json()
        if isinstance(posts, list) and posts:
            log_.info(f"Job already on WP: {title}")
            return posts[0]["id"], posts[0].get("link")
    except Exception:
        pass

    logo_url    = sanitize_text(job.get("companyLogo", ""), is_url=True)
    location    = sanitize_text(job.get("jobLocation", ""))
    raw_type    = sanitize_text(job.get("jobType", "")) or "Full-time"
    job_type_s  = JOB_TYPE_MAPPING.get(raw_type.lower().strip(), "full-time")
    company     = sanitize_text(job.get("companyName", ""))
    application = sanitize_text(job.get("application", ""), is_url=True)
    company_url = sanitize_text(job.get("companyUrl", ""), is_url=True)
    deadline    = sanitize_text(job.get("deadline", ""))
    co_website  = sanitize_text(job.get("companyWebsite", ""), is_url=True)
    qualif      = sanitize_text(job.get("jobQualifications", ""))
    experience  = sanitize_text(job.get("jobExperience", ""))
    co_address  = sanitize_text(job.get("companyAddress", ""))
    job_field   = sanitize_text(job.get("jobField", ""))
    salary      = sanitize_text(job.get("salaryRange", ""))
    about       = sanitize_text(job.get("companyDetails", ""))

    is_email = bool(re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", application))
    is_url_v = bool(re.match(r"^https?://[^\s]+$", application))
    if not (is_email or is_url_v):
        application = ""

    # Upload logo
    attachment_id = None
    if logo_url:
        try:
            img_r = requests.get(logo_url, headers={"User-Agent": HEADERS["User-Agent"]}, timeout=15)
            if img_r.status_code == 200:
                ct  = img_r.headers.get("Content-Type", "image/jpeg")
                ext = "png" if "png" in ct else "jpg"
                fn  = re.sub(r"[^a-z0-9]", "-", company.lower()) + "-logo." + ext
                up_h = dict(_wp_auth_headers())
                up_h["Content-Disposition"] = f"attachment; filename={fn}"
                up_h["Content-Type"] = ct
                up_r = requests.post(WP_MEDIA_URL, headers=up_h, data=img_r.content,
                                     auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
                if up_r.status_code in (200, 201):
                    attachment_id = up_r.json().get("id")
        except Exception as e:
            log_.warning(f"Logo upload failed: {e}")

    region_term_id   = get_or_create_term(f"{WP_BASE}/job_listing_region", location)
    job_type_term_id = get_or_create_term(f"{WP_BASE}/job_listing_type",
                                           job_type_s.replace("-", " ").title())

    payload = {
        "title":          title,
        "content":        description,
        "status":         "publish",
        "featured_media": attachment_id or 0,
        "meta": {
            "_job_title":          title,
            "_job_location":       location,
            "_job_type":           job_type_s,
            "_job_description":    description,
            "_application":        application,
            "_company_url":        company_url,
            "_job_expires":        deadline,
            "_company_name":       company,
            "_company_website":    co_website,
            "_company_logo":       str(attachment_id) if attachment_id else "",
            "_company_address":    co_address,
            "_company_details":    about,
            "_job_qualifications": qualif,
            "_job_experiences":    experience,
            "_job_field":          job_field,
            "_job_salary":         salary,
        },
    }
    if region_term_id:   payload["job_listing_region"] = [region_term_id]
    if job_type_term_id: payload["job_listing_type"]   = [job_type_term_id]

    for attempt in range(3):
        try:
            r = requests.post(WP_JOBS_URL, json=payload, headers=h,
                              auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
            r.raise_for_status()
            post = r.json()
            log_.info(f"Job posted: '{title}' -> WP ID {post.get('id')}")
            return post.get("id"), post.get("link")
        except Exception as e:
            log_.error(f"Job post attempt {attempt+1} failed: {e}")
            if attempt < 2:
                time.sleep(2 ** attempt)
    return None, None

# =============================================================================
#  STEP 0 — OFFICIAL GOV GATEWAY (gambia.gov.gm/vacancies/)
# =============================================================================
#
#  The government vacancies page carries no listings of its own; it only links
#  out to GamJobs and to the PSC site. We read it to (a) confirm GamJobs as the
#  live source and follow the exact URL it advertises, and (b) surface the PSC
#  link as the other official source. If the page is unreachable we fall back to
#  the configured GamJobs JOBS_URL.
# =============================================================================

def resolve_gov_gateway() -> str:
    """Return the GamJobs jobs URL to scrape, after consulting the gov gateway."""
    if not FOLLOW_GOV_GATEWAY:
        return JOBS_URL
    print(C_BLUE(f"\n  Reading official gov gateway: {GOV_GATEWAY_URL}"))
    try:
        soup = get_soup(GOV_GATEWAY_URL)
    except Exception as e:
        log(C_DIM(f"  Gateway unreachable ({e}); using configured JOBS_URL."))
        return JOBS_URL

    gamjobs_links, psc_links = [], []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        host = (urlparse(href).netloc or "").lower()
        if "gamjobs.com" in host:
            gamjobs_links.append(href)
        elif "pscgov.gm" in host:
            psc_links.append(href)

    if psc_links:
        log(C_DIM(f"  Gateway also lists PSC (other official source): {psc_links[0]}"))
    if not gamjobs_links:
        log(C_DIM("  Gateway did not expose a GamJobs link; using configured JOBS_URL."))
        return JOBS_URL

    # Prefer a direct /jobs/ link; otherwise normalise the homepage link to /jobs/.
    jobs_link = next((h for h in gamjobs_links if "/jobs" in urlparse(h).path), "")
    if not jobs_link:
        p = urlparse(gamjobs_links[0])
        jobs_link = f"{p.scheme}://{p.netloc}/jobs/"
    log(C_GREEN(f"  Gateway confirms GamJobs as live source -> {jobs_link}"))
    return jobs_link

# =============================================================================
#  STEP 1 — COLLECT JOB DETAIL URLS FROM THE GAMJOBS ARCHIVE (paginated)
# =============================================================================

def _norm_job_url(href: str) -> str:
    """Canonicalise a /jobs/<slug>/ URL: https host, no query/fragment, trailing /."""
    if not href:
        return ""
    absu = urljoin(BASE_URL + "/", href)
    p = urlsplit(absu)
    path = p.path
    if not path.endswith("/"):
        path += "/"
    return urlunsplit(("https", p.netloc.lower(), path, "", ""))

def _is_job_detail_path(path: str) -> bool:
    """True for /jobs/<slug>/ exactly (one segment after 'jobs', not 'page')."""
    parts = [s for s in path.split("/") if s]
    return len(parts) == 2 and parts[0] == "jobs" and parts[1].lower() != "page"

def _page_url(jobs_url: str, page: int) -> str:
    if page <= 1:
        return jobs_url
    base = jobs_url if jobs_url.endswith("/") else jobs_url + "/"
    return f"{base}page/{page}/"

def collect_job_links(jobs_url: str, max_pages: int = MAX_PAGES) -> list:
    """Walk the archive pages and return ordered, de-duplicated detail URLs."""
    print(C_BLUE(f"\n  Collecting job links from: {jobs_url}"))
    seen, ordered = set(), []
    empty_streak = 0

    for page in range(1, max_pages + 1):
        url = _page_url(jobs_url, page)
        try:
            soup = get_soup(url)
        except requests.HTTPError as e:
            log(C_DIM(f"  Page {page}: HTTP {getattr(e.response,'status_code','?')} — stopping."))
            break
        except Exception as e:
            log(C_DIM(f"  Page {page}: fetch error ({e}) — stopping."))
            break

        page_new = 0
        for a in soup.find_all("a", href=True):
            p = urlparse(a["href"])
            # Anchor may be relative; resolve path for the test.
            path = p.path or urlparse(urljoin(BASE_URL + "/", a["href"])).path
            if not _is_job_detail_path(path):
                continue
            norm = _norm_job_url(a["href"])
            if norm and norm not in seen:
                seen.add(norm)
                ordered.append(norm)
                page_new += 1

        log(f"    Page {page}: {page_new} new job link(s) (total {len(ordered)})")

        if page_new == 0:
            empty_streak += 1
            # Two consecutive empty pages -> we've run off the end.
            if empty_streak >= 2:
                break
        else:
            empty_streak = 0

        time.sleep(REQUEST_DELAY)

    return ordered

# =============================================================================
#  STEP 2 — PARSE ONE GAMJOBS DETAIL PAGE
# =============================================================================

# Order matters: the FIRST highly-specific selectors target JobMonster's real
# job body (`<div class="map-style-2" itemprop="description">`). Only if those
# miss do we fall back to generic WP content containers — and the <body> fallback
# is a last resort that would drag in site chrome (topbar/footer), so it is gated
# behind a "looks like a real post body" check below.
_CONTENT_SELECTORS = [
    'div.map-style-2[itemprop="description"]',   # JobMonster job body (exact)
    'div[itemprop="description"]',               # JobMonster job body (generic)
    "div.map-style-2",
    "div.single-job-content", "div.job-description", "div.single-job-description",
    "div.noo-job-content", "div.job-content", "div.job_description",
    "article .entry-content", "div.entry-content", "main .entry-content",
    "div.page-content",
]

def _find_content(soup):
    """Return the element most likely to hold the job post body."""
    best, best_len = None, 0
    for sel in _CONTENT_SELECTORS:
        el = soup.select_one(sel)
        if el:
            txt = el.get_text(" ", strip=True)
            if len(txt) > best_len:
                best, best_len = el, len(txt)
        if best and best_len > 300:
            return best
    if best:
        return best
    # Last resort: a single related-job <article> or <main>. Deliberately NOT the
    # whole <body> — returning <body> is what dragged the topbar/footer (Hotline,
    # Login, Navigation, footer email) into the description and made the site's
    # own info@gamjobs.com look like an apply address.
    main = soup.select_one("div.noo-main") or soup.find("main")
    if main:
        return main
    return soup.find("article") or soup.body or soup

def _anchors_in(scope, needle):
    out = []
    for a in scope.find_all("a", href=True):
        if needle in (urlparse(a["href"]).path or a["href"]):
            t = a.get_text(" ", strip=True)
            if t:
                out.append(t)
    return out

def _is_real_apply_url(href: str) -> bool:
    if not href:
        return False
    low = href.lower()
    if low.startswith("mailto:") or low.startswith("#") or low.startswith("javascript:"):
        return False
    if not low.startswith("http"):
        return False
    if any(s in low for s in _NON_APPLY_HOST_SUBSTR):
        return False
    if any(s in low for s in _NON_APPLY_PATH_SUBSTR):
        return False
    return True

def _is_apply_heading_line(line: str) -> bool:
    """
    True only for a SHORT standalone line that introduces the application
    instructions, e.g. 'How to Apply', 'HOW AND DEADLINE TO APPLY',
    'Mode of Application', 'Application Procedure'. The length guard stops it
    matching a mid-sentence 'to apply' inside a normal paragraph.
    """
    s = line.strip().lstrip("•*-–—#:. ").strip()
    if not s or len(s.split()) > 9:
        return False
    return bool(_APPLY_HEAD_PHRASES.match(s))

def _split_description_and_apply(content_text: str):
    """
    Given the cleaned body text, drop trailing boilerplate / UI lines and split
    off the 'How to Apply' tail. Returns (description, apply_text). The apply
    tail (which holds the apply email/URL and the deadline sentence) is removed
    from the description so the posted body is just the job content.
    """
    if not content_text:
        return "", ""

    lines = content_text.split("\n")
    kept = []
    for ln in lines:
        low = ln.strip().lower()
        if low in _BODY_DROP_LINES:
            continue
        if any(low.startswith(m) for m in _BODY_CUT_MARKERS):
            break
        kept.append(ln)

    # First line that is an apply heading marks the description/apply boundary.
    apply_idx = None
    for i, ln in enumerate(kept):
        if _is_apply_heading_line(ln):
            apply_idx = i
            break

    if apply_idx is None:
        return "\n".join(kept).strip(), ""

    description = "\n".join(kept[:apply_idx]).strip()
    apply_text  = "\n".join(kept[apply_idx:]).strip()
    if not description:                 # heading was the very first line — keep all
        return "\n".join(kept).strip(), ""
    return description, apply_text

def scrape_job_detail(url: str) -> dict:
    """Parse a single GamJobs /jobs/<slug>/ page into a raw_job dict."""
    soup = get_soup(url)

    # --- Title --------------------------------------------------------------
    h1 = (soup.select_one("h1.page-title") or soup.select_one("h1.job-title")
          or soup.select_one("h1.entry-title") or soup.find("h1"))
    title = clean_title(h1.get_text(" ", strip=True) if h1 else "")

    # --- Logo (og:image is the employer logo on this theme) -----------------
    logo = ""
    og = soup.find("meta", attrs={"property": "og:image"}) or \
         soup.find("meta", attrs={"name": "og:image"})
    if og and og.get("content"):
        logo = og["content"].strip()
    if not logo:
        emp_img = soup.select_one('a[href*="/employers/"] img')
        if emp_img and emp_img.get("src"):
            logo = emp_img["src"].strip()

    # --- Employer / company -------------------------------------------------
    company_name = ""
    company_url  = ""
    emp_a = soup.select_one('a[href*="/employers/"]')
    if emp_a:
        company_name = emp_a.get_text(" ", strip=True)
        company_url  = urljoin(BASE_URL + "/", emp_a["href"])
    if not company_name:
        company_name = "GamJobs Employer"

    # Company's own website (sidebar, title="Website") + address.
    company_website = ""
    site_a = soup.find("a", attrs={"title": "Website"})
    if site_a and site_a.get("href") and site_a["href"].startswith("http") \
            and "gamjobs.com" not in site_a["href"].lower():
        company_website = site_a["href"].strip()

    company_address = ""
    full_addr_label = soup.find(string=re.compile(r"Full Address", re.I))
    if full_addr_label:
        parent = getattr(full_addr_label, "parent", None)
        if parent:
            txt = parent.get_text(" ", strip=True)
            company_address = re.sub(r".*Full Address[:\s]*", "", txt, flags=re.I).strip()

    # --- Meta taxonomies ----------------------------------------------------
    job_type_opts  = _anchors_in(soup, "/job-type/")
    location_opts  = _anchors_in(soup, "/job-location/")
    category_opts  = _anchors_in(soup, "/job-category/")

    job_type = map_job_type(job_type_opts[0]) if job_type_opts else "full-time"
    location = pick_location(location_opts)
    job_field = ", ".join(dict.fromkeys(category_opts)) if category_opts else ""

    # --- Dates --------------------------------------------------------------
    # The meta row (near the type anchor) usually carries DD/MM/YYYY or a
    # 'posted - closing' range. The Job Overview box may carry an explicit
    # "Application Deadline".
    date_posted = ""
    deadline    = ""

    meta_text = ""
    type_a = soup.select_one('a[href*="/job-type/"]') or \
             soup.select_one('a[href*="/job-location/"]')
    if type_a:
        node = type_a
        for _ in range(4):
            node = node.parent
            if node is None:
                break
            txt = node.get_text(" ", strip=True)
            if DMY_DATE_RE.search(txt):
                meta_text = txt
                break
    meta_ds = dmy_dates(meta_text)
    if len(meta_ds) >= 2:
        date_posted, deadline = meta_ds[0], meta_ds[-1]
    elif len(meta_ds) == 1:
        deadline = meta_ds[0]

    # Explicit deadline label anywhere on the page wins if present.
    page_text_full = soup.get_text("\n")
    for lab in DEADLINE_LABELS:
        m = re.search(rf"{lab}\s*[:\-]?\s*([^\n<]+)", page_text_full, re.I)
        if m:
            d = parse_any_date(m.group(1))
            if d:
                deadline = d
                break
    if not date_posted:
        date_posted = datetime.now().strftime("%Y-%m-%d")
    if not deadline:
        # Sensible fallback so WP expiry is populated.
        deadline = (datetime.now() + timedelta(days=30)).strftime("%Y-%m-%d")

    # --- Body: description + how-to-apply -----------------------------------
    content_el = _find_content(soup)
    # Work on a copy so we don't disturb the live tree for later lookups.
    content_copy = BeautifulSoup(str(content_el), "lxml")
    content_text = html_block_to_text(content_copy)
    description, apply_text = _split_description_and_apply(content_text)
    if not description:
        description = content_text

    # --- Qualifications block (best-effort, for the WP field) ---------------
    # Heading forms seen on GamJobs: "Qualifications", "Qualifications & Experience",
    # "QUALIFICATIONS REQUIRED", "Qualification(s) and Experience".
    qualifications = ""
    qm = re.search(
        r"(?:^|\n)[ \t]*qualifications?(?:\s*(?:&|and)\s*experience)?(?:\s+\w+){0,3}\s*:?[ \t]*\n"
        r"(.*?)"
        r"(?:\n[ \t]*(?:how\s*(?:and|&)?\s*(?:deadline\s*)?to\s*apply|what\s+we\s+offer|"
        r"key\s+competenc|duration\s+of|method\s+of\s+application|mode\s+of\s+application)\b"
        r"|\n[ \t]*[A-Z][^\n]{0,60}:[ \t]*\n|\Z)",
        description, re.I | re.S)
    if qm:
        qualifications = qm.group(1).strip()[:1500]
    experience = extract_experience(qualifications or description)

    # --- Apply target (email or external URL) -------------------------------
    # Prefer anchors inside the apply tail; never use the on-page login button.
    apply_email = ""
    apply_url   = ""

    # 1) anchors within the content body
    for a in content_el.find_all("a", href=True):
        href = a["href"].strip()
        if href.lower().startswith("mailto:"):
            cand = extract_email(href[7:])
            if cand and _is_real_apply_email(cand):
                apply_email = apply_email or cand
        elif _is_real_apply_url(href):
            apply_url = apply_url or strip_tracking_params(href)

    # 2) plain-text fallbacks from the apply tail (or whole body)
    scan = apply_text or description
    if not apply_email:
        cand = extract_email(scan)
        if cand and _is_real_apply_email(cand):
            apply_email = cand
    if not apply_url:
        for u in URL_PATTERN.findall(scan):
            if _is_real_apply_url(u):
                apply_url = strip_tracking_params(u.rstrip(".,);"))
                break

    salary = extract_salary(description)

    return {
        "title":          title,
        "company_name":   company_name,
        "company_url":    company_url,
        "company_website":company_website,
        "company_address":company_address,
        "company_logo":   logo,
        "job_type":       job_type,
        "location":       location,
        "job_field":      job_field,
        "date_posted":    date_posted,
        "deadline":       deadline,
        "description":    description,
        "qualification":  qualifications,
        "experience":     experience,
        "salary":         salary,
        "apply_email":    apply_email,
        "apply_url":      apply_url,
        "apply_text":     apply_text,
        "job_url":        _norm_job_url(url),
    }

# =============================================================================
#  STEP 3 — DEDUPLICATE + PARAPHRASE + APPLY-RULE GATING
# =============================================================================

def process_job(raw_job: dict, processed_ids: set, processed_urls: set, seen_content: set):
    """
    Returns (status, job_dict_or_None):
        ("duplicate", None) — already processed / seen this run
        ("flagged",   None) — failed public-apply rule, written to flagged CSV
        ("ok",        dict) — ready to post to WordPress
    """
    job_url  = raw_job.get("job_url", "")
    title    = raw_job.get("title", "")
    company  = raw_job.get("company_name", "")
    location = raw_job.get("location", "")

    if not title:
        return "duplicate", None  # nothing usable

    job_id = make_job_id(job_url, title, company)

    if job_id in processed_ids or job_url in processed_urls:
        log(C_DIM(f"  Already processed (tracker) — skipped: {title}"))
        return "duplicate", None

    fingerprint = (title.lower().strip(), company.lower().strip(), location.lower().strip())
    if fingerprint in seen_content:
        log(C_DIM(f"  Duplicate content this run — skipped: {title}"))
        return "duplicate", None
    seen_content.add(fingerprint)

    # ---- Public-apply rule -------------------------------------------------
    apply_email = raw_job.get("apply_email", "")
    apply_url   = raw_job.get("apply_url", "")
    qualifies   = bool(apply_email) or bool(apply_url)

    if REQUIRE_PUBLIC_APPLY and not qualifies:
        write_flagged(raw_job,
                      "no public apply email or external URL (login-only on GamJobs)",
                      raw_job.get("apply_text", "")[:300])
        log(C_RED(f"  FLAGGED (no public apply) — {title}"))
        return "flagged", None

    # Record on scrape — before paraphrasing or posting.
    mark_scraped(job_id, job_url, title, company)
    processed_ids.add(job_id)
    processed_urls.add(job_url)

    description = raw_job.get("description", "")
    paraphrased_title = title
    paraphrased_desc  = description

    if ENABLE_PARAPHRASE and MISTRAL_API_KEY:
        print(C_BLUE(f"\n  Paraphrasing '{title}' ..."))
        paraphrased_title = paraphrase_title(title)
        paraphrased_desc  = paraphrase_description(description)
        mark_paraphrased(job_id)
    else:
        print(C_DIM("  Paraphrasing skipped (ENABLE_PARAPHRASE=False or MISTRAL_API_KEY not set)"))

    application = apply_email or apply_url
    apply_method = ("description_email" if apply_email
                    else "external_url" if apply_url else "not_found")

    company_link = raw_job.get("company_website") or raw_job.get("company_url", "")

    return "ok", {
        "jobTitle":          paraphrased_title,
        "jobDescription":    paraphrased_desc,
        "companyDetails":    "",
        "originalTitle":     title,
        "originalDesc":      description,
        "jobType":           raw_job.get("job_type", "full-time"),
        "jobQualifications": raw_job.get("qualification", ""),
        "jobExperience":     raw_job.get("experience", ""),
        "jobLocation":       location,
        "jobField":          raw_job.get("job_field", ""),
        "datePosted":        raw_job.get("date_posted", datetime.now().strftime("%Y-%m-%d")),
        "deadline":          raw_job.get("deadline", ""),
        "application":       application,
        "companyUrl":        company_link,
        "companyName":       company,
        "companyLogo":       raw_job.get("company_logo", ""),
        "companyWebsite":    raw_job.get("company_website", ""),
        "companyAddress":    raw_job.get("company_address", "") or location,
        "jobUrl":            job_url,
        "salaryRange":       raw_job.get("salary", ""),
        "_jobId":            job_id,
        "_apply_method":     apply_method,
        "_apply_raw":        raw_job.get("apply_text", "")[:160],
    }

# =============================================================================
#  VERBOSE PRINTER
# =============================================================================

def print_job_verbose(index, job):
    desc = job.get("jobDescription", "")
    desc_preview = (desc[:400] + " [...]") if len(desc) > 400 else desc

    print()
    print(C_DIVIDER())
    print(C_HEADER(f"  JOB #{index}"))
    print(C_DIVIDER())
    print(f"  {C_LABEL('Title (original)')}    : {C_VALUE(job.get('originalTitle',''))}")
    print(f"  {C_LABEL('Title (paraphrased)')} : {C_GREEN(job.get('jobTitle',''))}")
    print(f"  {C_LABEL('Job Type')}             : {job.get('jobType','') or C_DIM('—')}")
    print(f"  {C_LABEL('Qualification')}        : {(job.get('jobQualifications','')[:120] or C_DIM('—'))}")
    print(f"  {C_LABEL('Experience')}           : {job.get('jobExperience','') or C_DIM('—')}")
    print(f"  {C_LABEL('Location')}             : {job.get('jobLocation','') or C_DIM('—')}")
    print(f"  {C_LABEL('Category/Field')}       : {job.get('jobField','') or C_DIM('—')}")
    print(f"  {C_LABEL('Salary')}               : {job.get('salaryRange','') or C_DIM('—')}")
    print(f"  {C_LABEL('Posted')}               : {job.get('datePosted','') or C_DIM('—')}")
    print(f"  {C_LABEL('Deadline')}             : {job.get('deadline','') or C_DIM('—')}")

    application = job.get("application", "")
    print(f"  {C_LABEL('Apply')}                : {C_GREEN(application) if application else C_DIM('— not found —')}")
    print(f"  {C_LABEL('Apply Method')}         : {C_DIM(job.get('_apply_method',''))}")

    print()
    print(f"  {C_BLUE('── EMPLOYER ─────────────────────────────────────────')}")
    print(f"  {C_LABEL('Name')}      : {C_VALUE(job.get('companyName','') or C_DIM('—'))}")
    print(f"  {C_LABEL('Website')}   : {job.get('companyWebsite','') or C_DIM('—')}")
    print(f"  {C_LABEL('Source')}    : {job.get('companyUrl','') or C_DIM('—')}")
    print(f"  {C_LABEL('Logo')}      : {job.get('companyLogo','') or C_DIM('— none —')}")

    print()
    print(f"  {C_BLUE('── DESCRIPTION PREVIEW ─────────────────────────────')}")
    print(desc_preview if desc_preview else C_DIM("   — no description —"))
    print(f"  {C_LABEL('Job URL')}   : {job.get('jobUrl','')}")
    print(C_DIVIDER())

# =============================================================================
#  EXCEL SAVE (standardized column order)
# =============================================================================

EXCEL_HEADERS = [
    "Job Title", "Job Type", "Job Qualifications", "Job Experience",
    "Job Location", "Job Field", "Date Posted", "Deadline",
    "Job Description", "Application", "Company URL", "Company Name",
    "Company Logo", "Company Website", "Company Address",
    "Company Details", "Job URL", "Salary Range",
]

def _save_excel(jobs: list):
    if not _XLSX_AVAILABLE:
        log_.warning("pandas/openpyxl not installed — skipping Excel export")
        return
    if not jobs:
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(EXCEL_HEADERS)
    for job in jobs:
        ws.append([
            job["jobTitle"], job["jobType"], job["jobQualifications"], job["jobExperience"],
            job["jobLocation"], job["jobField"], job["datePosted"], job["deadline"],
            job["jobDescription"], job["application"], job["companyUrl"], job["companyName"],
            job["companyLogo"], job["companyWebsite"], job["companyAddress"],
            job["companyDetails"], job["jobUrl"], job["salaryRange"],
        ])
    wb.save(OUTPUT_FILE)
    log_.info(f"Saved {len(jobs)} rows -> {OUTPUT_FILE}")

# =============================================================================
#  MAIN
# =============================================================================

def main():
    start_time = datetime.now()

    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  GAMJOBS (GAMBIA) SCRAPER + MISTRAL PARAPHRASE + WORDPRESS POSTING"))
    print(C_HEADER("=" * 80))
    print(f"  Gov gateway     : {GOV_GATEWAY_URL}")
    print(f"  Jobs archive    : {JOBS_URL}")
    print(f"  Public-apply    : {'✅ enforced (flag others)' if REQUIRE_PUBLIC_APPLY else '❌ off (post all)'}")
    print(f"  Max new jobs    : {'unlimited' if not MAX_JOBS else MAX_JOBS}")
    print(f"  Max pages       : {MAX_PAGES}")
    print(f"  Paraphrase      : {'✅ enabled' if (ENABLE_PARAPHRASE and MISTRAL_API_KEY) else '❌ disabled'}")
    print(f"  WordPress post  : {'✅ enabled' if (WP_USER and WP_PASSWORD) else '❌ disabled'}")
    print(f"  Excel export    : {'✅ enabled' if _XLSX_AVAILABLE else '❌ disabled (pip install pandas openpyxl)'}")
    print(f"  NLP gating      : {'✅' if _NLP_AVAILABLE else '⚠️  no sentence-transformers / language-tool'}")
    print(f"  Started         : {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(C_HEADER("=" * 80))

    _init_tracker()
    _init_flagged()
    processed_ids, processed_urls = load_processed_ids()
    print(f"  Tracker loaded: {len(processed_ids)} previously processed job IDs")

    # Step 0: consult the official gov gateway, then collect links from GamJobs.
    jobs_url = resolve_gov_gateway()

    try:
        job_links = collect_job_links(jobs_url)
    except Exception as e:
        log(C_RED(f"  FATAL: could not collect job links: {e}"))
        return

    if not job_links:
        log(C_RED("  No job links found — nothing to do."))
        return
    print(C_GREEN(f"\n  Found {len(job_links)} job detail page(s) to process.\n"))

    jobs_out = []
    seen_content = set()
    posted_count = 0
    flagged_count = 0
    dup_count = 0
    errors = 0
    scraped = 0

    for link in job_links:
        # Skip detail fetch entirely if URL already processed.
        if link in processed_urls:
            dup_count += 1
            log(C_DIM(f"  Already processed (tracker) — skipped: {link}"))
            continue

        try:
            raw_job = scrape_job_detail(link)
            scraped += 1
        except Exception as e:
            errors += 1
            log(C_RED(f"  ERROR scraping {link} : {e}"))
            time.sleep(REQUEST_DELAY)
            continue

        try:
            status, job = process_job(raw_job, processed_ids, processed_urls, seen_content)
        except Exception as e:
            errors += 1
            log(C_RED(f"  ERROR processing '{raw_job.get('title','')}' : {e}"))
            continue

        if status == "duplicate":
            dup_count += 1
            time.sleep(REQUEST_DELAY)
            continue
        if status == "flagged":
            flagged_count += 1
            time.sleep(REQUEST_DELAY)
            continue

        jobs_out.append(job)
        print_job_verbose(len(jobs_out), job)

        print(C_BLUE("\n  Posting to WordPress …"))
        wp_id, wp_url = post_job_to_wordpress(job)
        if wp_id:
            mark_posted(job["_jobId"], wp_id, wp_url or "")
            posted_count += 1
            print(C_GREEN(f"  WP ID={wp_id}  {wp_url}"))
        else:
            mark_failed(job["_jobId"], "wp_post_failed_or_skipped")
            print(C_RED("  WordPress post failed / skipped"))

        if len(jobs_out) % 25 == 0:
            _save_excel(jobs_out)

        if MAX_JOBS and len(jobs_out) >= MAX_JOBS:
            log(f"\nMAX_JOBS limit ({MAX_JOBS}) reached, stopping.")
            break

        time.sleep(REQUEST_DELAY)

    _save_excel(jobs_out)

    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds() / 60.0
    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  SCRAPE COMPLETE"))
    print(C_HEADER("=" * 80))
    print(f"  {C_LABEL('Job links found')}           : {len(job_links)}")
    print(f"  {C_LABEL('Detail pages scraped')}      : {scraped}")
    print(f"  {C_LABEL('New jobs processed')}        : {C_GREEN(str(len(jobs_out)))}")
    print(f"  {C_LABEL('Posted to WordPress')}       : {C_GREEN(str(posted_count))}")
    print(f"  {C_LABEL('Flagged (no public apply)')} : {flagged_count}")
    print(f"  {C_LABEL('Duplicates skipped')}        : {dup_count}")
    print(f"  {C_LABEL('Errors')}                    : {C_RED(str(errors)) if errors else '0'}")
    print(f"  {C_LABEL('Duration')}                  : ~{duration:.1f} min")
    print(f"  {C_LABEL('Output file')}               : {OUTPUT_FILE}")
    print(f"  {C_LABEL('Tracker file')}              : {PROCESSED_IDS_FILE}")
    print(f"  {C_LABEL('Flagged file')}              : {FLAGGED_FILE}")

    if jobs_out:
        with_apply = sum(1 for j in jobs_out if j.get("application"))
        with_email = sum(1 for j in jobs_out if "@" in (j.get("application") or ""))
        with_url   = with_apply - with_email
        print(f"\n  {C_LABEL('Application links:')}")
        print(f"    External URL : {with_url}")
        print(f"    Email found  : {with_email}")

        para_count = sum(1 for j in jobs_out if j.get("jobTitle") != j.get("originalTitle"))
        print(f"\n  {C_LABEL('Paraphrased titles')} : {para_count}/{len(jobs_out)}")

        with_deadline = sum(1 for j in jobs_out if j.get("deadline"))
        print(f"  {C_LABEL('Deadline captured')}  : {with_deadline}/{len(jobs_out)}")

    print(C_HEADER("=" * 80))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)
