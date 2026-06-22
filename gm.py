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
#  This script scrapes TWO independent Gambian job sources and feeds both
#  into the same dedup tracker / paraphrase / WordPress-posting / Excel-export
#  pipeline:
#
#    1) PSC (Personnel Services Commission) - https://pscgov.gm/vacancies-1/
#       A single Elementor accordion page holding every current civil-service
#       vacancy. No per-job emails or external apply URLs; every posting
#       points at the same e-recruitment portal (https://portal.pscgov.gm).
#
#    2) Gamjobs - https://gamjobs.com/jobs/
#       A JobMonster/WP-Job-Manager style board: paginated listing pages,
#       per-job detail pages with structured "Job Overview" meta, real
#       per-employer "How to Apply" sections (public emails or URLs, no
#       login wall), categories, locations and employer logos. Per the
#       site's own taxonomy this also includes tenders/RFPs/RFQs/EOIs
#       alongside ordinary vacancies -- by design, all of it is scraped.
#
#  APPLY-METHOD NOTE (PSC only)
#  -----------------------------
#  Every PSC posting shares ONE application method: complete form 16(A/B/C)
#  on the e-recruitment portal https://portal.pscgov.gm (requires candidate
#  registration). There are no per-job emails or external apply URLs. This
#  collides with the hard "public-apply-only" rule used elsewhere.
#
#  PSC_PORTAL_AS_APPLY (default "1"/on):
#     on  -> the public portal URL counts as a valid apply destination; jobs post.
#     off -> jobs are treated as non-qualifying and written to the flagged CSV
#            instead of being posted (strict public-apply behaviour).
#
#  Gamjobs postings carry their own real apply email/URL per listing, so they
#  are not subject to the PSC_PORTAL_AS_APPLY toggle.
# =============================================================================

BASE_URL      = "https://pscgov.gm"
VACANCIES_URL = os.environ.get("PSC_VACANCIES_URL", "https://pscgov.gm/vacancies-1/")

# The single, constant application destination for every PSC vacancy.
PSC_PORTAL_URL = os.environ.get("PSC_PORTAL_URL", "https://portal.pscgov.gm/")
# PSC is the recruiting body; use its logo for every posting unless overridden.
PSC_LOGO_URL = os.environ.get(
    "PSC_LOGO_URL",
    "https://pscgov.gm/wp-content/uploads/2022/06/cropped-PSCLOGO-270x270.png",
)
PSC_WEBSITE = "https://pscgov.gm/"

# Treat the (public, external) e-recruitment portal as a valid apply URL.
PSC_PORTAL_AS_APPLY = os.environ.get("PSC_PORTAL_AS_APPLY", "1") != "0"

# ── Gamjobs ──────────────────────────────────────────────────────────────────
GAMJOBS_ENABLED   = os.environ.get("GAMJOBS_ENABLED", "1") != "0"
GAMJOBS_BASE_URL  = "https://gamjobs.com"
GAMJOBS_JOBS_URL  = os.environ.get("GAMJOBS_JOBS_URL", "https://gamjobs.com/jobs/")
GAMJOBS_WEBSITE   = "https://gamjobs.com/"
GAMJOBS_MAX_PAGES = int(os.environ.get("GAMJOBS_MAX_PAGES", "0"))  # 0 = no cap

REQUEST_DELAY   = float(os.environ.get("REQUEST_DELAY", "1.0"))
MAX_JOBS        = int(os.environ.get("MAX_JOBS", "0"))
REQUEST_TIMEOUT = int(os.environ.get("REQUEST_TIMEOUT", "25"))

OUTPUT_FILE        = "gambia_jobs.xlsx"
PROCESSED_IDS_FILE = "gambia_jobs_processed.csv"
FLAGGED_FILE       = "gambia_jobs_flagged.csv"

# CSV column names — defined once so _init_tracker, load, and upsert all agree.
_TRACKER_FIELDS = ["Job ID", "Job URL", "Job Title", "Company Name",
                   "Status", "Timestamp", "WP ID"]

_FLAGGED_FIELDS = ["Source", "Title", "Ministry", "Location", "Salary",
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
    "volunteer": "volunteer",
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

# Known Gambian towns/cities, used to pull a location out of free text.
GAMBIA_LOCATIONS = [
    "Banjul", "Abuko", "Brikama", "Bakau", "Serrekunda", "Serekunda",
    "Kanifing", "Old Yundum", "Yundum", "Basse", "Farafenni", "Soma",
    "Kerewan", "Mansakonko", "Janjanbureh", "Kuntaur", "Sukuta", "Gunjur",
    "Lamin", "Bwiam", "Kanilai", "Essau", "Barra", "Kotu", "Fajara",
]
DEFAULT_LOCATION = os.environ.get("PSC_DEFAULT_LOCATION", "Banjul")

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

TRACKING_PARAM_PREFIXES = ("utm_",)
TRACKING_PARAM_EXACT = {
    "fbclid", "gclid", "msclkid", "mc_cid", "mc_eid", "ref", "referrer",
}

# Headers that delimit the sections inside a single vacancy block.
# Sorted longest-first at match time so "QUALIFICATION AND EXPERIENCE" wins over
# "QUALIFICATION", etc.
SECTION_HEADERS = [
    "VACANCY ANNOUNCEMENT",
    "JOB PURPOSE",
    "JOB TITLE",
    "NUMBER OF VACANCIES",
    "QUALIFICATION AND EXPERIENCE",
    "QUALIFICATIONS",
    "QUALIFICATION",
    "DUTIES AND RESPONSIBILITIES",
    "DUTIES AND RESPONSIBILTIES",   # site frequently misspells this
    "KEY RESPONSIBILITIES",
    "COMPETENCIES/ SKILLS",
    "COMPETENCIES/SKILLS",
    "COMPETENCIES",
    "SKILLS AND ABILITIES",
    "SKILLS",
    "OUTPUTS/DELIVERABLES",
    "OUTPUTS",
    "SALARY",
    "APPLICATION FORMAT",
    "CLOSING DATE",
    "APPLICATION DEADLINE",
]
_HEADERS_BY_LEN = sorted(SECTION_HEADERS, key=len, reverse=True)

MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11,
    "december": 12, "jan": 1, "feb": 2, "mar": 3, "apr": 4, "jun": 6, "jul": 7,
    "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
}

DATE_RE = re.compile(
    r"(\d{1,2})\s*(?:st|nd|rd|th)?\s+([A-Za-z]+)\s*[.,]?\s*(\d{4})", re.I
)

# Gamjobs detail pages render the deadline as DD/MM/YYYY (sometimes as a
# "publish - expiry" range); this pattern handles that format separately
# from the PSC "7th June 2026" style handled by DATE_RE above.
SLASH_DATE_RE = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{4})")

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

def _normalize_header_line(line: str):
    """
    If `line` begins with one of the known section headers (ignoring markdown
    decoration and case), return (CANONICAL_HEADER, inline_remainder). Else None.
    """
    stripped = line.strip()
    if not stripped:
        return None
    # Drop leading markdown decoration: #, *, •, -, whitespace.
    bare = re.sub(r"^[#*•\-\s]+", "", stripped)
    bare = bare.replace("*", "")
    upper = bare.upper()
    for header in _HEADERS_BY_LEN:
        if upper.startswith(header):
            remainder = bare[len(header):]
            remainder = remainder.lstrip(" :–-").strip()
            # Guard: a real header is short. Avoid matching a sentence that merely
            # starts with the word "SKILLS"/"SALARY" mid-paragraph by requiring the
            # remainder to be empty or to look like a value (not a long sentence)
            # for the very short headers.
            if header in ("SKILLS", "SALARY", "QUALIFICATION", "COMPETENCIES",
                          "OUTPUTS") and len(remainder.split()) > 18:
                continue
            return header, remainder
    return None

def segment_sections(text: str) -> dict:
    """Split a vacancy's text into {CANONICAL_HEADER: content} sections."""
    sections = {}
    current = "_PREAMBLE"
    buf = []

    def flush():
        if buf:
            existing = sections.get(current, "")
            joined = "\n".join(buf).strip()
            sections[current] = (existing + "\n" + joined).strip() if existing else joined

    for line in text.split("\n"):
        hdr = _normalize_header_line(line)
        if hdr:
            flush()
            buf = []
            current = hdr[0]
            if hdr[1]:
                buf.append(hdr[1])
        else:
            if line.strip():
                buf.append(line.strip())
    flush()
    return sections

def section(sections: dict, *names) -> str:
    for n in names:
        if sections.get(n):
            return sections[n].strip()
    return ""

# =============================================================================
#  GAMBIA-SPECIFIC EXTRACTORS
# =============================================================================

def parse_gambia_date(text: str) -> str:
    """
    Parse a closing-date string into 'YYYY-MM-DD'. Handles ordinals
    (7th, 2nd, 19th), optional commas, and ranges like
    '19th June – 18th July, 2026' (returns the later/end date).
    """
    if not text:
        return ""
    matches = DATE_RE.findall(text)
    if not matches:
        return ""
    # Walk matches newest-last; return the last one whose month name is real
    # (covers "start – end, YEAR" ranges -> end date, and ignores stray matches).
    for day_s, mon_s, year_s in reversed(matches):
        month = MONTHS.get(mon_s.lower())
        if not month:
            continue
        try:
            return datetime(int(year_s), month, int(day_s)).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""

def parse_slash_date(text: str, take_last=True) -> str:
    """
    Parse Gamjobs-style DD/MM/YYYY dates into 'YYYY-MM-DD'. When a listing
    shows a "publish - expiry" range (e.g. '16/06/2026 - 30/06/2026'),
    take_last=True returns the expiry (later) date.
    """
    if not text:
        return ""
    matches = SLASH_DATE_RE.findall(text)
    if not matches:
        return ""
    ordered = list(reversed(matches)) if take_last else matches
    for day_s, mon_s, year_s in ordered:
        try:
            return datetime(int(year_s), int(mon_s), int(day_s)).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""

def extract_salary_gmd(text: str):
    """
    Return (salary_string, grade_string) from a SALARY section.
    e.g. 'GMD 88,512 per annum (Grade 7)', 'Grade 7'.
    """
    if not text:
        return "", ""
    grade = ""
    gm = re.search(r"Grade\s*([0-9]+)", text, re.I)
    if gm:
        grade = f"Grade {gm.group(1)}"
    amt = ""
    # Amounts on the PSC page appear as D88,512 / D 133,860.00 / GMD101, 892.00
    # (note the stray space after the comma in some rows).
    am = re.search(r"(?:GMD|D)\s*([0-9]{1,3}(?:,\s?[0-9]{3})*(?:\.[0-9]+)?)", text)
    if am:
        amt = re.sub(r"\s+", "", am.group(1))
    if amt:
        salary = f"GMD {amt} per annum"
        if grade:
            salary += f" ({grade})"
        return salary, grade
    if grade:
        return grade, grade
    return "", ""

def extract_experience(qual_text: str) -> str:
    """Pull a short experience requirement out of the qualifications text."""
    if not qual_text:
        return ""
    m = re.search(r"at least\s+\d+\s+years?[^.\n;]*", qual_text, re.I)
    if m:
        return m.group(0).strip().rstrip(".")
    m = re.search(r"\b\d+\s+years?[^.\n;]*experience", qual_text, re.I)
    if m:
        return m.group(0).strip().rstrip(".")
    return ""

def extract_ministry(title_hint: str, body_text: str) -> str:
    """Resolve the recruiting ministry/employer for a vacancy."""
    # 1) Parenthetical on the accordion title: "... (Ministry of Agriculture)-..."
    if title_hint:
        m = re.search(r"\(([^)]*Ministry[^)]*)\)", title_hint, re.I)
        if m:
            return re.sub(r"\s+", " ", m.group(1)).strip()
    # 2) Body phrasing: "... at/under the Ministry of X, ..."
    if body_text:
        m = re.search(
            r"(?:at|under|within)\s+the\s+(Ministry of [^.,\n;]+)", body_text, re.I
        )
        if m:
            ministry = re.sub(r"\s+", " ", m.group(1)).strip()
            # Trim trailing location fragments that sometimes ride along.
            ministry = re.split(r"\s+(?:The Quadrangle|MECCNAR|Complex)\b",
                                 ministry, flags=re.I)[0].strip()
            return ministry
    return "Government of The Gambia"

def extract_location(body_text: str) -> str:
    if body_text:
        for town in GAMBIA_LOCATIONS:
            if re.search(rf"\b{re.escape(town)}\b", body_text, re.I):
                # Normalise a couple of spellings.
                return "Serrekunda" if town.lower() == "serekunda" else town
    return DEFAULT_LOCATION

def build_description(sections: dict) -> str:
    """Assemble the human-readable description (purpose + duties + skills)."""
    parts = []
    purpose = section(sections, "JOB PURPOSE")
    if purpose:
        parts.append(purpose)

    vacancies = section(sections, "NUMBER OF VACANCIES")
    if vacancies:
        parts.append(f"Number of vacancies: {vacancies}")

    duties = section(sections, "DUTIES AND RESPONSIBILITIES",
                     "DUTIES AND RESPONSIBILTIES", "KEY RESPONSIBILITIES")
    if duties:
        parts.append("Duties and Responsibilities:\n" + duties)

    skills = section(sections, "SKILLS AND ABILITIES", "COMPETENCIES/ SKILLS",
                     "COMPETENCIES/SKILLS", "COMPETENCIES", "SKILLS")
    if skills:
        parts.append("Skills and Abilities:\n" + skills)

    outputs = section(sections, "OUTPUTS/DELIVERABLES", "OUTPUTS")
    if outputs:
        parts.append("Outputs/Deliverables:\n" + outputs)

    return "\n\n".join(p.strip() for p in parts if p.strip()).strip()

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
                raw_job.get("source", "Unknown"),
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
#  STEP 1a — SCRAPE PSC (single Elementor accordion page)
# =============================================================================

def _find_job_blocks(soup: BeautifulSoup):
    """
    Return a list of (title_hint, body_text) tuples — one per vacancy.

    Strategy A (DOM): pair Elementor accordion/toggle titles with their content
    panels across the various widget class names Elementor has shipped.
    Strategy B (text fallback): if DOM detection yields < 2 blocks, segment the
    whole page text on 'JOB TITLE' anchors, which every posting contains.
    """
    blocks = []

    title_selectors = [
        ".elementor-accordion .elementor-tab-title",
        ".elementor-accordion .elementor-accordion-title",
        ".elementor-toggle .elementor-tab-title",
        ".elementor-toggle .elementor-toggle-title",
        "details.e-n-accordion-item > summary",
    ]
    content_selectors = [
        ".elementor-tab-content",
        ".elementor-accordion-content",
        ".elementor-toggle-content",
        ".e-n-accordion-item-content",
    ]

    for tsel in title_selectors:
        titles = soup.select(tsel)
        if not titles:
            continue
        for t in titles:
            title_hint = clean_text(t)
            body_el = None
            # Look for the matching content panel: aria-controls -> id, else the
            # next sibling content element, else nearest content descendant.
            ctrl = t.get("aria-controls")
            if ctrl:
                body_el = soup.find(id=ctrl)
            if body_el is None:
                sib = t.find_next_sibling()
                if sib is not None and any(
                    cls in (sib.get("class") or [])
                    for cls in ("elementor-tab-content", "elementor-accordion-content",
                                "elementor-toggle-content", "e-n-accordion-item-content")
                ):
                    body_el = sib
            if body_el is None:
                # For <details><summary>, the content is the remaining children.
                parent = t.parent
                if parent is not None:
                    body_el = parent
            if body_el is not None and title_hint:
                blocks.append((title_hint, html_block_to_text(body_el)))
        if blocks:
            break  # first selector family that works wins

    # Filter DOM blocks to those that actually look like vacancies.
    blocks = [(h, b) for (h, b) in blocks
              if b and re.search(r"job\s*title|vacancy announcement|application format",
                                 b, re.I)]

    if len(blocks) >= 2:
        return blocks

    # ---- Strategy B: text segmentation fallback ----------------------------
    log("  DOM accordion detection weak — falling back to text segmentation.")
    main = (soup.select_one("main") or soup.select_one("#content")
            or soup.select_one("article") or soup.body or soup)
    full = html_block_to_text(main)

    # Split on each 'JOB TITLE' marker. keep the delimiter with the following chunk.
    pieces = re.split(r"(?im)^\s*(?:#+\s*|\*+\s*)?JOB\s*TITLE\s*:?", full)
    # pieces[0] is the page preamble before the first job; skip it.
    text_blocks = []
    for chunk in pieces[1:]:
        chunk = chunk.strip()
        if not chunk:
            continue
        # The first line after the JOB TITLE marker is the title itself.
        first_line = chunk.split("\n", 1)[0].strip(" :*").strip()
        # Re-prepend a JOB TITLE header so segment_sections sees it.
        rebuilt = "JOB TITLE: " + chunk
        text_blocks.append((first_line, rebuilt))

    return text_blocks if text_blocks else blocks

def scrape_psc_vacancies():
    """Fetch the PSC vacancies page and return a list of raw job dicts."""
    log(f"\n{'='*80}\nFETCHING PSC VACANCIES PAGE: {VACANCIES_URL}\n{'='*80}")
    soup = get_soup(VACANCIES_URL)

    raw_blocks = _find_job_blocks(soup)
    log(f"  Detected {len(raw_blocks)} vacancy block(s) on the page.")

    today = datetime.now()
    fallback_deadline = (today + timedelta(days=90)).strftime("%Y-%m-%d")

    jobs = []
    seen_anchor = set()

    for title_hint, body_text in raw_blocks:
        if not body_text:
            continue
        sections = segment_sections(body_text)

        # Clean title: prefer the explicit JOB TITLE line.
        title = section(sections, "JOB TITLE").split("\n")[0].strip()
        if not title:
            # Fall back to the accordion header, minus ministry parenthetical / month.
            title = re.sub(r"\([^)]*\)", "", title_hint)
            title = re.sub(r"[-–]\s*[A-Za-z]+\s*\d{4}\s*$", "", title).strip()
        if not title:
            continue

        ministry  = extract_ministry(title_hint, body_text)
        location  = extract_location(body_text)
        qual      = section(sections, "QUALIFICATION AND EXPERIENCE",
                            "QUALIFICATIONS", "QUALIFICATION")
        experience = extract_experience(qual)
        salary_section = section(sections, "SALARY")
        salary, grade  = extract_salary_gmd(salary_section)
        closing  = section(sections, "CLOSING DATE", "APPLICATION DEADLINE")
        app_text = section(sections, "APPLICATION FORMAT")

        deadline = parse_gambia_date(closing) or parse_gambia_date(app_text) or fallback_deadline

        description = build_description(sections)
        if not description:
            # Last resort: use the preamble (the "Applications are invited..." text).
            description = section(sections, "_PREAMBLE", "VACANCY ANNOUNCEMENT")

        # Application destination: a real email if the posting ever lists one,
        # otherwise the public e-recruitment portal.
        apply_email = extract_email(app_text)
        apply_url   = ""
        portal_match = re.search(r"https?://portal\.pscgov\.gm/?", app_text, re.I)
        if portal_match:
            apply_url = portal_match.group(0)
        elif re.search(r"portal", app_text, re.I):
            apply_url = PSC_PORTAL_URL

        # Stable, unique job_url for dedup (page + title/ministry slug).
        anchor = slugify(f"{title}-{ministry}-{grade}")
        if anchor in seen_anchor:
            continue
        seen_anchor.add(anchor)
        job_url = f"{VACANCIES_URL}#{anchor}"

        jobs.append({
            "source":        "PSC Gambia",
            "title":         title,
            "job_url":       job_url,
            "company_name":  ministry,
            "location":      location,
            "qualification": qual,
            "experience":    experience,
            "salary":        salary,
            "grade":         grade,
            "deadline":      deadline,
            "description":   description,
            "apply_url":     apply_url,
            "apply_email":   apply_email,
            "apply_text":    app_text,
            "company_logo":  PSC_LOGO_URL,
            "company_website": PSC_WEBSITE,
            "source_page":   VACANCIES_URL,
            "job_type":      "",
        })

    log(f"  Parsed {len(jobs)} unique vacancy record(s).")
    return jobs

# =============================================================================
#  STEP 1b — SCRAPE GAMJOBS (paginated WP Job Manager / JobMonster board)
# =============================================================================

def _gamjobs_listing_page_url(page_num: int) -> str:
    if page_num <= 1:
        return GAMJOBS_JOBS_URL
    base = GAMJOBS_JOBS_URL.rstrip("/")
    return f"{base}/?page={page_num}"

def _gamjobs_extract_listing_links(soup: BeautifulSoup) -> list:
    """
    Pull unique job detail-page URLs from a /jobs/ listing page. Job links are
    <h3><a href=".../jobs/<slug>/"></a></h3> style headings inside job cards;
    selecting on the href pattern is more robust than chasing card classes
    across theme tweaks.
    """
    links = []
    seen = set()
    for a in soup.select("a[href]"):
        href = a.get("href", "")
        if not href or "/jobs/" not in href:
            continue
        # Skip taxonomy / pagination / filter links, only keep individual job
        # detail permalinks: https://gamjobs.com/jobs/<slug>/
        parsed = urlparse(href)
        path = parsed.path.rstrip("/")
        if not path.startswith("/jobs/"):
            continue
        slug = path[len("/jobs/"):]
        if not slug or "?" in href.split("/jobs/")[-1]:
            continue
        # The bare listing root and its paginated variants aren't job detail pages.
        if slug in ("", "page"):
            continue
        clean_url = urljoin(GAMJOBS_BASE_URL, path + "/")
        if clean_url in seen:
            continue
        seen.add(clean_url)
        links.append(clean_url)
    return links

def _gamjobs_total_count(soup: BeautifulSoup):
    """Parse 'Showing 1–20 of 50 jobs' if present, to know when paging is done."""
    text = clean_text(soup.body) if soup.body else ""
    m = re.search(r"Showing\s+[\d,]+\s*[-–]\s*([\d,]+)\s+of\s+([\d,]+)\s+jobs", text, re.I)
    if not m:
        return None, None
    shown_so_far = int(m.group(1).replace(",", ""))
    total        = int(m.group(2).replace(",", ""))
    return shown_so_far, total

def gamjobs_collect_links(already_processed_urls: set) -> list:
    """
    Walk /jobs/ pagination, collecting detail-page URLs, until either:
      - a page returns no new (unprocessed) links, or
      - we've paged past the site's own reported total, or
      - GAMJOBS_MAX_PAGES is hit (if configured as a safety cap).
    """
    log(f"\n{'='*80}\nFETCHING GAMJOBS LISTING: {GAMJOBS_JOBS_URL}\n{'='*80}")
    all_links = []
    seen_links = set()
    page_num = 1

    while True:
        if GAMJOBS_MAX_PAGES and page_num > GAMJOBS_MAX_PAGES:
            log(f"  Reached GAMJOBS_MAX_PAGES cap ({GAMJOBS_MAX_PAGES}), stopping pagination.")
            break

        page_url = _gamjobs_listing_page_url(page_num)
        log(f"  Page {page_num}: {page_url}")
        try:
            soup = get_soup(page_url)
        except Exception as e:
            log(C_RED(f"    FAILED to fetch page {page_num}: {e}"))
            break

        page_links = _gamjobs_extract_listing_links(soup)
        new_links  = [u for u in page_links if u not in seen_links]

        if not page_links:
            log("    No job links found on this page — stopping pagination.")
            break

        for u in new_links:
            seen_links.add(u)
            all_links.append(u)

        new_unprocessed = [u for u in new_links if u not in already_processed_urls]
        log(f"    Found {len(page_links)} link(s) on page "
            f"({len(new_links)} new, {len(new_unprocessed)} unprocessed).")

        shown, total = _gamjobs_total_count(soup)
        if total is not None and shown is not None and shown >= total:
            log(f"    Listing reports {shown}/{total} shown — reached end of results.")
            break

        # Stop once an entire page contributes nothing we haven't already
        # tracked from a previous run — going further would just re-fetch
        # jobs we already posted.
        if page_num > 1 and not new_unprocessed and all(u in already_processed_urls for u in page_links):
            log("    Entire page already in tracker — stopping pagination.")
            break

        time.sleep(REQUEST_DELAY)
        page_num += 1

        # Hard safety stop regardless of config, to avoid a runaway loop if the
        # "Showing X of Y" marker ever goes missing on every page.
        if page_num > 200:
            log(C_RED("    Safety stop: exceeded 200 pages."))
            break

    log(f"  Collected {len(all_links)} unique job detail link(s) across pagination.")
    return all_links

_OVERVIEW_LABELS = ["Vacancy Ref", "Department", "Reports To", "Location",
                    "Employment Type", "Application Deadline", "Salary"]

def _gamjobs_parse_overview(soup: BeautifulSoup) -> dict:
    """
    Pull the "Job Overview" key/value list. The Gamjobs job-description
    Markdown is rendered so that each VALUE line is immediately followed by
    its own bold LABEL line (value first, "**Label:**" second) rather than
    the more conventional label-then-value order, e.g.:

        Information Technology Department
        **Department:**
        Chief Technology Officer (CTO)
        **Reports To:**

    So for each known label we find the <strong>/<b> tag containing it and
    take the text immediately *preceding* that tag (its previous sibling, or
    the tail of the previous block) as the value, rather than text after it.
    Falls back to a flat-text regex scan for layouts that do put value after
    label, so both orderings are covered.
    """
    overview = {}
    container = (soup.select_one(".job_overview, .job-overview, .job_listing-meta")
                 or soup.select_one("article") or soup.body)
    if container is None:
        return overview

    seen_labels = {l.lower() for l in _OVERVIEW_LABELS}

    for label in _OVERVIEW_LABELS:
        label_re = re.compile(rf"^\s*{re.escape(label)}\s*:?\s*$", re.I)
        tag = None
        for candidate in container.find_all(["strong", "b"]):
            if label_re.match(candidate.get_text(strip=True)):
                tag = candidate
                break
        if tag is None:
            continue

        value = ""
        prev = tag.find_previous(string=True)
        steps = 0
        while prev is not None and steps < 6:
            candidate_text = prev.strip()
            if candidate_text and candidate_text.lower().rstrip(":") not in seen_labels:
                value = candidate_text
                break
            prev = prev.find_previous(string=True)
            steps += 1

        if value:
            overview[label] = value

    if not overview:
        # Fallback: flat-text "Label: value" on one line (covers themes that
        # render label-then-value in normal reading order).
        text_block = clean_text(container)
        for label in _OVERVIEW_LABELS:
            m = re.search(rf"{re.escape(label)}\s*:?\s*([^\n]+?)(?=\s+(?:{'|'.join(_OVERVIEW_LABELS)})\b|$)",
                          text_block, re.I)
            if m:
                overview[label] = m.group(1).strip(" -:")

    return overview

def _gamjobs_extract_apply(soup: BeautifulSoup, page_text: str):
    """
    Find the real public apply destination from a Gamjobs detail page: prefer
    an explicit URL/email under a 'How to Apply' heading, then fall back to
    any email/external link in the page body. Returns (apply_url, apply_email).
    """
    apply_url = ""
    apply_email = ""

    how_to_apply = soup.find(string=re.compile(r"how\s+to\s+apply", re.I))
    search_scope = soup
    if how_to_apply:
        # Search within the section following the "How to Apply" heading.
        parent = how_to_apply.find_parent(["h1", "h2", "h3", "h4", "strong", "p"])
        if parent:
            following = parent.find_all_next(limit=15)
            for el in following:
                if el.name == "a" and el.get("href"):
                    href = el.get("href", "")
                    if href.startswith("mailto:"):
                        apply_email = href[len("mailto:"):].split("?")[0]
                        break
                    if href.startswith("http") and "gamjobs.com" not in href:
                        apply_url = strip_tracking_params(href)
                        break

    if not apply_url and not apply_email:
        # Fall back: any external (non-gamjobs) link or mailto in the body.
        for a in soup.select("a[href]"):
            href = a.get("href", "")
            if href.startswith("mailto:"):
                apply_email = href[len("mailto:"):].split("?")[0]
                break
            if href.startswith("http") and "gamjobs.com" not in href and "facebook" not in href \
               and "twitter" not in href and "linkedin" not in href and "instagram" not in href \
               and "whatsapp" not in href.lower():
                apply_url = strip_tracking_params(href)
                break

    if not apply_email:
        apply_email = extract_email(page_text)

    return apply_url, apply_email

def _gamjobs_extract_logo(soup: BeautifulSoup) -> str:
    og = soup.select_one('meta[property="og:image"]')
    if og and og.get("content"):
        return og["content"]
    img = soup.select_one(".company_logo img, .company-logo img")
    if img and img.get("src"):
        return urljoin(GAMJOBS_BASE_URL, img["src"])
    return ""

def _gamjobs_extract_company(soup: BeautifulSoup) -> tuple:
    """Return (company_name, company_url) from the employer link near the title."""
    link = soup.select_one('a[href*="/employers/"]')
    if link:
        name = clean_text(link)
        url  = urljoin(GAMJOBS_BASE_URL, link.get("href", ""))
        return name or "Unknown Employer", url
    return "Unknown Employer", ""

def scrape_gamjobs_detail(job_url: str):
    """Fetch and parse a single Gamjobs job/tender detail page into a raw job dict."""
    try:
        soup = get_soup(job_url)
    except Exception as e:
        log(C_RED(f"    FAILED to fetch {job_url}: {e}"))
        return None

    page_text = clean_text(soup.body) if soup.body else ""

    title_el = soup.select_one("h1")
    title = clean_text(title_el) if title_el else ""
    # Strip the trailing "NNN views" counter some themes append to <h1>.
    title = re.sub(r"\s*\d[\d,]*\s+views\s*$", "", title, flags=re.I).strip()
    if not title:
        return None

    company_name, company_url = _gamjobs_extract_company(soup)

    job_type = ""
    type_link = soup.select_one('a[href*="/job-type/"]')
    if type_link:
        job_type = clean_text(type_link)

    location = DEFAULT_LOCATION
    loc_links = soup.select('a[href*="/job-location/"]')
    if loc_links:
        loc_text = ", ".join(clean_text(l) for l in loc_links if clean_text(l))
        # Drop the generic "The Gambia" country tag when a town is also present.
        parts = [p.strip() for p in loc_text.split(",") if p.strip()]
        specific = [p for p in parts if p.lower() != "the gambia"]
        location = specific[0] if specific else (parts[0] if parts else DEFAULT_LOCATION)

    categories = []
    for cat_link in soup.select('a[href*="/job-category/"]'):
        cname = clean_text(cat_link)
        if cname:
            categories.append(cname)
    job_field = categories[0] if categories else ""

    overview = _gamjobs_parse_overview(soup)

    # Deadline: prefer the explicit "Application Deadline" overview field,
    # then the publish-expiry date range shown next to the job type/location.
    deadline = ""
    if overview.get("Application Deadline"):
        deadline = parse_slash_date(overview["Application Deadline"]) or \
                   parse_gambia_date(overview["Application Deadline"])
    if not deadline:
        header_block = soup.select_one("h1")
        header_text = ""
        if header_block:
            following_siblings_text = clean_text(header_block.find_parent())
            header_text = following_siblings_text or page_text[:400]
        deadline = parse_slash_date(header_text) or parse_slash_date(page_text[:1000])
    if not deadline:
        deadline = (datetime.now() + timedelta(days=60)).strftime("%Y-%m-%d")

    salary = overview.get("Salary", "")

    # Main content: the body copy sits in the post/article content area.
    content_el = (soup.select_one(".job_description, .single_job_listing, "
                                   "article .entry-content, article")
                  or soup.select_one("main"))
    description = html_block_to_text(content_el) if content_el else page_text[:2000]
    # Trim trailing site furniture that sometimes gets swept into a loose
    # "article" selector (share links, related-jobs heading, etc.).
    description = re.split(r"\n(?:Share:|Related Jobs|Apply for this job)\b",
                           description, maxsplit=1)[0].strip()

    qual_match = re.search(
        r"Qualifications?\s*&?\s*Experience\s*\n(.+?)(?=\n\*\*|\n[A-Z][a-zA-Z ]{3,30}\n|\Z)",
        description, re.I | re.S)
    qualification = qual_match.group(1).strip() if qual_match else ""
    experience = extract_experience(qualification or description)

    apply_url, apply_email = _gamjobs_extract_apply(soup, page_text)
    logo = _gamjobs_extract_logo(soup)

    return {
        "source":          "Gamjobs",
        "title":           title,
        "job_url":         job_url,
        "company_name":    company_name,
        "location":        location,
        "qualification":   qualification,
        "experience":      experience,
        "salary":          salary,
        "grade":           "",
        "deadline":        deadline,
        "description":     description,
        "apply_url":       apply_url,
        "apply_email":     apply_email,
        "apply_text":      description[-500:],
        "company_logo":    logo,
        "company_website": company_url or GAMJOBS_WEBSITE,
        "source_page":     job_url,
        "job_type":        job_type,
    }

def scrape_gamjobs_vacancies(already_processed_urls: set):
    """Crawl Gamjobs pagination and parse every job/tender detail page found."""
    links = gamjobs_collect_links(already_processed_urls)

    jobs = []
    for i, link in enumerate(links, start=1):
        if link in already_processed_urls:
            continue
        log(f"  [{i}/{len(links)}] Fetching detail page: {link}")
        job = scrape_gamjobs_detail(link)
        if job:
            jobs.append(job)
        time.sleep(REQUEST_DELAY)

    log(f"  Parsed {len(jobs)} new Gamjobs record(s).")
    return jobs

# =============================================================================
#  STEP 2 — DEDUPLICATE + PARAPHRASE + APPLY-RULE GATING
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
    ministry = raw_job.get("company_name", "")
    location = raw_job.get("location", "")
    source   = raw_job.get("source", "Unknown")

    job_id = make_job_id(job_url, title, ministry)

    if job_id in processed_ids or job_url in processed_urls:
        log(C_DIM(f"  Already processed (tracker) — skipped: {title}"))
        return "duplicate", None

    fingerprint = (title.lower().strip(), ministry.lower().strip(), location.lower().strip())
    if fingerprint in seen_content:
        log(C_DIM(f"  Duplicate content this run — skipped: {title}"))
        return "duplicate", None
    seen_content.add(fingerprint)

    # ---- Public-apply rule -------------------------------------------------
    # PSC: every posting shares the same login-only portal, gated by
    # PSC_PORTAL_AS_APPLY. Gamjobs: each listing carries its own real public
    # apply email/URL already, so it qualifies whenever either is present —
    # no separate site-wide toggle needed.
    apply_email = raw_job.get("apply_email", "")
    apply_url   = raw_job.get("apply_url", "")

    if source == "PSC Gambia":
        qualifies = bool(apply_email) or (PSC_PORTAL_AS_APPLY and bool(apply_url))
        non_qualify_reason = ("login-only e-recruitment portal; PSC_PORTAL_AS_APPLY is off"
                              if apply_url else "no public apply email or URL")
    else:
        qualifies = bool(apply_email) or bool(apply_url)
        non_qualify_reason = "no public apply email or URL found on listing"

    if not qualifies:
        write_flagged(raw_job, non_qualify_reason, raw_job.get("apply_text", "")[:300])
        log(C_RED(f"  FLAGGED (non-qualifying apply) — {title}"))
        return "flagged", None

    # Record on scrape — before paraphrasing or posting.
    mark_scraped(job_id, job_url, title, ministry)
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
    apply_method = ("portal_url" if apply_url and not apply_email
                    else "description_email" if apply_email else "not_found")

    return "ok", {
        "jobTitle":          paraphrased_title,
        "jobDescription":    paraphrased_desc,
        "companyDetails":    "",
        "originalTitle":     title,
        "originalDesc":      description,
        "jobType":           raw_job.get("job_type", "") or "Full-time",
        "jobQualifications": raw_job.get("qualification", ""),
        "jobExperience":     raw_job.get("experience", ""),
        "jobLocation":       location,
        "jobField":          ministry if source == "PSC Gambia" else raw_job.get("job_type", "") or ministry,
        "datePosted":        datetime.now().strftime("%Y-%m-%d"),
        "deadline":          raw_job.get("deadline", ""),
        "application":       application,
        "companyUrl":        raw_job.get("company_website", ""),
        "companyName":       ministry,
        "companyLogo":       raw_job.get("company_logo", ""),
        "companyWebsite":    raw_job.get("company_website", ""),
        "companyAddress":    location,
        "jobUrl":            job_url,
        "salaryRange":       raw_job.get("salary", ""),
        "_source":           source,
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
    print(C_HEADER(f"  JOB #{index}  [{job.get('_source', '')}]"))
    print(C_DIVIDER())
    print(f"  {C_LABEL('Title (original)')}    : {C_VALUE(job.get('originalTitle',''))}")
    print(f"  {C_LABEL('Title (paraphrased)')} : {C_GREEN(job.get('jobTitle',''))}")
    print(f"  {C_LABEL('Job Type')}             : {job.get('jobType','') or C_DIM('—')}")
    print(f"  {C_LABEL('Qualification')}        : {(job.get('jobQualifications','')[:120] or C_DIM('—'))}")
    print(f"  {C_LABEL('Experience')}           : {job.get('jobExperience','') or C_DIM('—')}")
    print(f"  {C_LABEL('Location')}             : {job.get('jobLocation','') or C_DIM('—')}")
    print(f"  {C_LABEL('Ministry/Field')}       : {job.get('jobField','') or C_DIM('—')}")
    print(f"  {C_LABEL('Salary')}               : {job.get('salaryRange','') or C_DIM('—')}")
    print(f"  {C_LABEL('Posted')}               : {job.get('datePosted','') or C_DIM('—')}")
    print(f"  {C_LABEL('Deadline')}             : {job.get('deadline','') or C_DIM('—')}")

    application = job.get("application", "")
    print(f"  {C_LABEL('Apply')}                : {C_GREEN(application) if application else C_DIM('— not found —')}")
    print(f"  {C_LABEL('Apply Method')}         : {C_DIM(job.get('_apply_method',''))}")

    print()
    print(f"  {C_BLUE('── EMPLOYER ─────────────────────────────────────────')}")
    print(f"  {C_LABEL('Name')}      : {C_VALUE(job.get('companyName','') or C_DIM('—'))}")
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
    "Source", "Job Title", "Job Type", "Job Qualifications", "Job Experience",
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
            job.get("_source", ""), job["jobTitle"], job["jobType"], job["jobQualifications"],
            job["jobExperience"], job["jobLocation"], job["jobField"], job["datePosted"],
            job["deadline"], job["jobDescription"], job["application"], job["companyUrl"],
            job["companyName"], job["companyLogo"], job["companyWebsite"], job["companyAddress"],
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
    print(C_HEADER("  GAMBIA JOBS SCRAPER (PSC + GAMJOBS) + MISTRAL PARAPHRASE + WP POSTING"))
    print(C_HEADER("=" * 80))
    print(f"  PSC source page  : {VACANCIES_URL}")
    print(f"  PSC apply portal : {PSC_PORTAL_URL}")
    print(f"  PSC portal=apply : {'✅ yes (jobs post)' if PSC_PORTAL_AS_APPLY else '❌ no (flag to CSV)'}")
    print(f"  Gamjobs enabled  : {'✅ yes' if GAMJOBS_ENABLED else '❌ no'}")
    if GAMJOBS_ENABLED:
        print(f"  Gamjobs source   : {GAMJOBS_JOBS_URL}")
        print(f"  Gamjobs max pages: {'unlimited (until no new jobs)' if not GAMJOBS_MAX_PAGES else GAMJOBS_MAX_PAGES}")
    print(f"  Max new jobs     : {'unlimited' if not MAX_JOBS else MAX_JOBS}")
    print(f"  Paraphrase       : {'✅ enabled' if (ENABLE_PARAPHRASE and MISTRAL_API_KEY) else '❌ disabled'}")
    print(f"  WordPress post   : {'✅ enabled' if (WP_USER and WP_PASSWORD) else '❌ disabled'}")
    print(f"  Excel export     : {'✅ enabled' if _XLSX_AVAILABLE else '❌ disabled (pip install pandas openpyxl)'}")
    print(f"  NLP gating       : {'✅' if _NLP_AVAILABLE else '⚠️  no sentence-transformers / language-tool'}")
    print(f"  Started          : {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(C_HEADER("=" * 80))

    _init_tracker()
    _init_flagged()
    processed_ids, processed_urls = load_processed_ids()
    print(f"  Tracker loaded: {len(processed_ids)} previously processed job IDs\n")

    raw_jobs = []

    try:
        raw_jobs.extend(scrape_psc_vacancies())
    except Exception as e:
        log(C_RED(f"  ERROR scraping PSC vacancies page: {e}"))

    if GAMJOBS_ENABLED:
        try:
            raw_jobs.extend(scrape_gamjobs_vacancies(processed_urls))
        except Exception as e:
            log(C_RED(f"  ERROR scraping Gamjobs: {e}"))

    if not raw_jobs:
        log(C_RED("  FATAL: no vacancies parsed from any source."))
        return

    jobs_out = []
    seen_content = set()
    posted_count = 0
    flagged_count = 0
    dup_count = 0
    errors = 0
    source_counts = {}

    for raw_job in raw_jobs:
        try:
            status, job = process_job(raw_job, processed_ids, processed_urls, seen_content)
        except Exception as e:
            errors += 1
            log(C_RED(f"  ERROR processing job '{raw_job.get('title','')}' : {e}"))
            continue

        if status == "duplicate":
            dup_count += 1
            continue
        if status == "flagged":
            flagged_count += 1
            continue

        src = job.get("_source", "Unknown")
        source_counts[src] = source_counts.get(src, 0) + 1

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
    print(f"  {C_LABEL('Vacancies parsed (all sources)')} : {len(raw_jobs)}")
    for src, cnt in source_counts.items():
        print(f"  {C_LABEL('  -> ' + src)}{'':<{max(1, 20-len(src))}} : {cnt}")
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
        print(f"    URL          : {with_url}")
        print(f"    Email found  : {with_email}")

        para_count = sum(1 for j in jobs_out if j.get("jobTitle") != j.get("originalTitle"))
        print(f"\n  {C_LABEL('Paraphrased titles')} : {para_count}/{len(jobs_out)}")

        with_salary = sum(1 for j in jobs_out if j.get("salaryRange"))
        print(f"  {C_LABEL('Salary captured')}    : {with_salary}/{len(jobs_out)}")

        with_deadline = sum(1 for j in jobs_out if j.get("deadline"))
        print(f"  {C_LABEL('Deadline captured')}  : {with_deadline}/{len(jobs_out)}")

    print(C_HEADER("=" * 80))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)
