"""
IDN/HCO Parent Organization Identification
Streamlit UI — Single-step: Upload → Run → Download
"""

import subprocess
import streamlit as st
import pandas as pd
import time
import sys
import os
import random
import json
import re
import requests
import warnings
from urllib.parse import urlparse, urljoin

# Suppress Streamlit 'missing ScriptRunContext' warning
import logging
logging.getLogger("streamlit.runtime.scriptrunner.script_run_context").setLevel(logging.ERROR)
from bs4 import BeautifulSoup

subprocess.check_call(
    [sys.executable, "-m", "pip", "install", "duckduckgo-search", "ddgs", "groq", "openpyxl"],
    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
)

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="IDN/HCO Parent Identifier",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #ADD8E6; }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1a1f2e 0%, #ADD8E6 100%);
        border-right: 1px solid #2d3748;
    }

    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #4f8ef7, #7c5cbf);
    }

    .stButton > button {
        background: linear-gradient(135deg, #4f8ef7 0%, #7c5cbf 100%);
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.2s;
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 15px rgba(79, 142, 247, 0.4);
    }

    h1, h2, h3 { color: #e2e8f0 !important; }

    .upload-card {
        background: linear-gradient(135deg, #1a1f2e 0%, #1e2538 100%);
        border: 1px solid #2d3748;
        border-radius: 14px;
        padding: 28px 32px;
        margin: 16px 0;
    }

    .result-card {
        background: linear-gradient(135deg, #1a2e1a, #1a2e2e);
        border: 2px solid #48bb78;
        border-radius: 14px;
        padding: 24px 28px;
        margin-top: 20px;
    }

    .metric-box {
        background: #1a1f2e;
        border: 1px solid #2d3748;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
    }
    .metric-value {
        color: #4f8ef7;
        font-size: 26px;
        font-weight: 700;
    }
    .metric-label {
        color: #718096;
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-top: 4px;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────
EXCLUDE_DOMAINS = [
    "yelp.com", "linkedin.com", "facebook.com", "twitter.com", "instagram.com",
    "healthgrades.com", "webmd.com", "npino.com", "wikipedia.org",
    "yellowpages.com", "bbb.org", "mapquest.com", "vitals.com",
    "zocdoc.com", "doximity.com", "sharecare.com", "everydayhealth.com",
    "usnews.com", "medicare.gov", "medicaid.gov", "npidb.org",
    "countryoffice.org", "freeclinicdirectory.org", "manta.com",
    "chamberofcommerce.com", "cms.gov", "rxlist.com", "ratemds.com",
    "healthscores.com", "ehealthscores.com",
]

ABOUT_SLUGS = [
    "/about-us", "/about", "/who-we-are", "/our-story", "/about/",
    "/about-us/", "/our-organization", "/company", "/overview",
]

SSO_PATTERNS = [
    r"login", r"signin", r"sign-in", r"sso", r"auth", r"okta\.com",
    r"ping\.identity", r"microsoft\.com/adfs", r"onelogin", r"duo\.com",
    r"saml", r"oauth", r"idp\.", r"myworkday", r"workday\.com",
]

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/119.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/121.0",
]

IDN_KEYWORDS = [
    "health", "hospital", "medical", "system", "network", "care", "services",
    "centre", "center", "institute", "foundation", "regional", "community",
    "mercy", "memorial", "physicians", "wellness", "rehabilitation", "therapy",
    "children", "cancer", "heart", "clinic", "pharmacy", "infusion",
    "saint", "st.", "memorial", "general", "university",
]

_GENERIC = {
    "www", "portal", "patient", "app", "my", "online", "health",
    "care", "web", "secure", "login", "sso", "auth", "workforce",
    "mychartplus", "mychart", "myhealth", "myaccount",
}

TOP_N_URLS = 4
DELAY_MIN  = 1.5
DELAY_MAX  = 3.0


# ─────────────────────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────────────────────
def init_state():
    defaults = {
        "uploaded_df":     None,
        "final_df":        None,
        "pipeline_done":   False,
        "pipeline_running":False,
        "groq_keys":       [],
        "key_index":       0,
        "col_map":         {},
        "log_messages":    [],
        "output_filename": "idn_results.csv",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()


# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
def log(msg: str):
    ts = time.strftime("%H:%M:%S")
    st.session_state.log_messages.append(f"[{ts}] {msg}")

def get_headers():
    return {
        "User-Agent":                random.choice(USER_AGENTS),
        "Accept":                    "text/html,application/xhtml+xml,*/*;q=0.8",
        "Accept-Language":           "en-US,en;q=0.9",
        "Accept-Encoding":           "gzip, deflate, br",
        "Connection":                "keep-alive",
        "Upgrade-Insecure-Requests": "1",
    }

def extract_domain(url: str) -> str:
    return urlparse(url).netloc.lower().replace("www.", "")

def is_excluded(domain: str) -> bool:
    return any(excl in domain for excl in EXCLUDE_DOMAINS)

def get_groq_client():
    from groq import Groq
    keys = st.session_state.groq_keys
    idx  = st.session_state.key_index
    return Groq(api_key=keys[idx])

def rotate_key() -> bool:
    nxt = st.session_state.key_index + 1
    if nxt < len(st.session_state.groq_keys):
        st.session_state.key_index = nxt
        log(f"🔄 Switched to Groq API Key {nxt + 1}")
        return True
    log("❌ All Groq API keys exhausted — waiting 60s then retrying Key 1")
    time.sleep(60)
    st.session_state.key_index = 0
    return False

def is_rate_limit_error(e) -> bool:
    err = str(e).lower()
    return any(x in err for x in ["429", "rate limit", "quota", "too many", "exceeded"])


# ─────────────────────────────────────────────────────────────
# PHASE 1 — URL DISCOVERY (internal)
# ─────────────────────────────────────────────────────────────
def safe_search(query: str, num_results: int = 10, retries: int = 3):
    try:
        from ddgs import DDGS
        for attempt in range(retries):
            try:
                with DDGS() as ddgs:
                    return [r["href"] for r in ddgs.text(query, max_results=num_results)]
            except Exception as e:
                log(f"Search attempt {attempt+1} failed: {e}")
                time.sleep(5 * (attempt + 1))
        return []
    except ImportError:
        log("⚠️  ddgs not installed.")
        return []

def get_candidates(name, addr1, city, state, zip_code):
    queries = [
        f'"{name}" {city} {state} location site',
        f'"{name}" {addr1} {city} {state} location page',
        f"{name} {city} {state} find a location",
        f"{name} {addr1} {state} specific location page",
    ]
    all_urls, seen = [], set()
    for query in queries:
        results = safe_search(query, num_results=10)
        for url in results:
            domain = extract_domain(url)
            if not is_excluded(domain) and domain not in seen:
                seen.add(domain)
                all_urls.append(url)
        if len(all_urls) >= 15:
            break
        if not results:
            time.sleep(random.uniform(1, 2))
    return all_urls[:20]

def rank_urls(site_name, address, city, state, zip_code, candidate_urls):
    if not candidate_urls:
        return ["", "", "", ""]

    url_list = "\n".join([f"{i+1}. {u}" for i, u in enumerate(candidate_urls)])
    prompt = f"""You are an expert at identifying the SPECIFIC LOCATION PAGE for US healthcare facilities.

FACILITY:
- Name: {site_name}
- Address: {address}, {city}, {state} {zip_code}

CANDIDATE URLs:
{url_list}

Pick the top {TOP_N_URLS} URLs that best point to THIS SPECIFIC FACILITY's location page on its official website.

RULES:
1. PREFER specific sub-pages mentioning this exact location, address, city, or facility name
2. Must be from the facility's own official site or its parent health system — NOT a directory
3. If URL contains keywords like "loc", "location", "site", "office", "center" rank it best
4. REJECT: NPI lookup sites, {EXCLUDE_DOMAINS}
5. Return fewer than 4 if no good matches — do NOT return junk to fill slots
6. Strongly prefer .org domains

Respond with ONLY a JSON array, nothing else.
Example: ["https://example.org/locations/main-campus/"]
If nothing good found: []"""

    keys_tried = 0
    while keys_tried < len(st.session_state.groq_keys) + 1:
        try:
            client = get_groq_client()
            response = client.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=300,
                temperature=0.0,
            )
            raw = response.choices[0].message.content.strip()
            match = re.search(r'\[.*?\]', raw, re.DOTALL)
            if match:
                urls = json.loads(match.group())
                urls = [u for u in urls if isinstance(u, str) and u.startswith("http")]
                while len(urls) < TOP_N_URLS:
                    urls.append("")
                return urls[:TOP_N_URLS]
            else:
                return (candidate_urls + [""] * TOP_N_URLS)[:TOP_N_URLS]
        except Exception as e:
            if is_rate_limit_error(e):
                log(f"⚠️  Key {st.session_state.key_index + 1} rate limited")
                rotate_key()
                keys_tried += 1
            else:
                log(f"❌ AI ranking error: {e}")
                return (candidate_urls + [""] * TOP_N_URLS)[:TOP_N_URLS]

    return (candidate_urls + [""] * TOP_N_URLS)[:TOP_N_URLS]


# ─────────────────────────────────────────────────────────────
# PHASE 2 — IDN EXTRACTION (internal)
# ─────────────────────────────────────────────────────────────
def domain_name_hint(url: str) -> str:
    try:
        host  = urlparse(url).netloc.lower().split(":")[0]
        parts = host.split(".")
        sld   = parts[-2] if len(parts) >= 2 else parts[0]
        for prefix in ["my", "e", "i"]:
            if sld.startswith(prefix) and len(sld) > len(prefix) + 3:
                candidate = sld[len(prefix):]
                if candidate not in _GENERIC:
                    sld = candidate
                    break
        useful_parts = [p for p in parts[:-2] if p not in _GENERIC and len(p) > 2]
        if sld in _GENERIC and useful_parts:
            sld = useful_parts[-1]
        spaced = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", sld)
        spaced = re.sub(r"[-_]+", " ", spaced)
        hint   = spaced.title().strip()
        return hint if hint and hint.lower() not in _GENERIC else ""
    except Exception:
        return ""

def is_sso_url(url: str) -> bool:
    return any(re.search(p, url, re.I) for p in SSO_PATTERNS)

def is_sso_page(html: str, final_url: str) -> bool:
    if is_sso_url(final_url):
        return True
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text(" ", strip=True).lower()
    if len(text) < 400 and any(w in text for w in ["password", "sign in", "login", "username"]):
        return True
    return False

def fetch_html_requests(url: str, timeout: int = 12) -> tuple:
    if not url or str(url).strip() in ("", "nan"):
        return "", ""
    for attempt in range(2):
        try:
            resp = requests.get(
                url, headers=get_headers(),
                timeout=timeout, allow_redirects=True, verify=False,
            )
            if resp.status_code == 200:
                return resp.url, resp.text
            elif resp.status_code in (403, 429):
                time.sleep(2 * (attempt + 1))
            else:
                break
        except requests.exceptions.Timeout:
            time.sleep(2)
        except Exception:
            break
    return "", ""

def fetch_html(url: str) -> tuple:
    final_url, html = fetch_html_requests(url)
    if html and not is_sso_page(html, final_url):
        return final_url, html
    return final_url, html

def fetch_about_page(base_url: str) -> tuple:
    for slug in ABOUT_SLUGS:
        about_url  = base_url.rstrip("/") + slug
        final_url, html = fetch_html_requests(about_url, 8)
        if html and not is_sso_page(html, final_url):
            return final_url, html
    return "", ""

def get_base_url(url: str) -> str:
    p = urlparse(url)
    return f"{p.scheme}://{p.netloc}"

def extract_signals(html: str, url: str) -> dict:
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg", "iframe"]):
        tag.decompose()

    signals = {}

    title = soup.find("title")
    signals["title"] = title.get_text(strip=True) if title else ""

    def meta(attr, val):
        t = soup.find("meta", {attr: val})
        return t.get("content", "").strip() if t else ""

    signals["og_site_name"]     = meta("property", "og:site_name")
    signals["og_title"]         = meta("property", "og:title")
    signals["meta_description"] = meta("name", "description")
    signals["app_name"]         = meta("name", "application-name")
    signals["twitter_site"]     = meta("name", "twitter:site")

    jsonld_names = []
    org_types = {"HealthCare", "Organization", "Hospital", "MedicalClinic", "MedicalOrganization",
                 "LocalBusiness", "MedicalBusiness", "MedicalUniversity", "HealthcareOrganization"}
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data  = json.loads(script.string or "")
            items = data if isinstance(data, list) else [data]
            flat  = []
            for item in items:
                if isinstance(item, dict) and "@graph" in item:
                    flat.extend(item["@graph"])
                else:
                    flat.append(item)
            for item in flat:
                if not isinstance(item, dict):
                    continue
                t      = item.get("@type", "")
                types  = t if isinstance(t, list) else [t]
                if any(x in org_types for x in types):
                    name = item.get("name", "").strip()
                    if name:
                        jsonld_names.append(f"{name} (type={types})")
                parent = item.get("parentOrganization", {})
                if isinstance(parent, dict):
                    pname = parent.get("name", "").strip()
                    if pname:
                        jsonld_names.append(f"{pname} (parentOrganization)")
        except Exception:
            pass
    signals["jsonld_names"] = jsonld_names

    footer = soup.find("footer") or soup.find(attrs={"role": "contentinfo"})
    if footer:
        ft = footer.get_text(separator=" ", strip=True)
        m  = re.search(r'(©|copyright|&copy;).{0,250}', ft, re.IGNORECASE)
        signals["footer_copyright"] = m.group(0)[:300] if m else ""
        signals["footer_text"]      = ft[:500]
    else:
        page_text = soup.get_text(separator=" ", strip=True)
        m = re.search(r'(©|copyright).{0,250}', page_text, re.IGNORECASE)
        signals["footer_copyright"] = m.group(0)[:300] if m else ""
        signals["footer_text"]      = ""

    header    = soup.find("header") or soup.find(attrs={"role": "banner"})
    logo_alts = []
    area      = header if header else soup
    for img in area.find_all("img", alt=True):
        alt = img.get("alt", "").strip()
        src = img.get("src", "").lower()
        cls = " ".join(img.get("class", [])).lower()
        if any(x in src + cls for x in ["logo", "brand", "header"]) and alt:
            cleaned = re.sub(r'\s*(logo|icon|banner|image|svg|png|jpg|jpeg)\s*$', '', alt, flags=re.I).strip()
            if cleaned and cleaned.lower() not in _GENERIC:
                logo_alts.append(cleaned)
    signals["logo_alts"] = logo_alts

    svg_labels = []
    for svg_el in soup.find_all("svg"):
        lbl = svg_el.get("aria-label", "").strip()
        if lbl and len(lbl) > 3:
            svg_labels.append(lbl)
    signals["svg_aria_labels"] = svg_labels[:3]

    home_texts = []
    if header:
        for a in header.find_all("a", href=True):
            href = a.get("href", "")
            if re.match(r'^(https?://[^/]+/?|/|#)$', href):
                t = a.get_text(strip=True)
                if 3 < len(t) < 60:
                    home_texts.append(t)
    signals["home_link_texts"] = home_texts[:3]

    h1 = soup.find("h1")
    signals["h1"] = h1.get_text(strip=True)[:120] if h1 else ""

    main = soup.find("main") or soup.find(attrs={"role": "main"}) or soup.find("body")
    signals["body_snippet"] = " ".join(main.get_text(separator=" ", strip=True).split()[:150]) if main else ""

    url_hint = ""
    for pattern in [r"client_id=([a-z0-9_-]{4,40})", r"/idp/([a-z0-9_-]{4,40})",
                    r"realm=([a-z0-9_-]{4,40})", r"tenant=([a-z0-9_-]{4,40})"]:
        m = re.search(pattern, url, re.I)
        if m:
            raw = m.group(1)
            if raw.lower() not in _GENERIC:
                url_hint = raw.replace("-", " ").replace("_", " ").title()
                break
    signals["url_hint"]    = url_hint
    signals["domain_hint"] = domain_name_hint(url)
    return signals

def signals_to_block(signals: dict, url: str) -> str:
    lines = [f"URL: {url}", "", "=== META / STRUCTURED DATA ==="]
    for k, lbl in [("og_site_name", "og:site_name"), ("app_name", "app-name    "),
                   ("twitter_site", "twitter:site"), ("og_title", "og:title    "),
                   ("title", "<title>     "), ("meta_description", "description ")]:
        if signals.get(k):
            lines.append(f"  {lbl}: {signals[k][:180]}")
    if signals.get("jsonld_names"):
        lines += ["", "=== JSON-LD (most reliable) ==="]
        for n in signals["jsonld_names"]:
            lines.append(f"  {n}")
    lines += ["", "=== HEADER / LOGO ==="]
    if signals.get("logo_alts"):
        lines.append(f"  Logo alts       : {signals['logo_alts']}")
    if signals.get("svg_aria_labels"):
        lines.append(f"  SVG aria-labels : {signals['svg_aria_labels']}")
    if signals.get("home_link_texts"):
        lines.append(f"  Home-link texts : {signals['home_link_texts']}")
    lines += ["", "=== FOOTER / COPYRIGHT ==="]
    if signals.get("footer_copyright"):
        lines.append(f"  Copyright : {signals['footer_copyright']}")
    if signals.get("footer_text"):
        lines.append(f"  Footer    : {signals['footer_text'][:400]}")
    lines += ["", "=== PAGE CONTENT ==="]
    if signals.get("h1"):
        lines.append(f"  H1  : {signals['h1']}")
    if signals.get("body_snippet"):
        lines.append(f"  Body: {signals['body_snippet'][:400]}")
    lines += ["", "=== FALLBACK HINTS ==="]
    if signals.get("url_hint"):
        lines.append(f"  URL param hint : {signals['url_hint']}")
    if signals.get("domain_hint"):
        lines.append(f"  Domain hint    : {signals['domain_hint']}")
    return "\n".join(lines)

SYSTEM_PROMPT = """You are a senior healthcare data analyst identifying US IDNs/HCOs.

Signal priority (highest → lowest):
1. JSON-LD structured data  (parentOrganization first if present)
2. og:site_name meta tag
3. Footer copyright line
4. Logo alt text / SVG aria-label
5. Home-link anchor text
6. <title> tag
7. Body text
8. URL param hint
9. Domain hint

Rules:
- Return the PARENT health system name, not a department or individual facility.
- Strip generic suffixes: "Patient Portal", "MyChart", "Employee Login", "Login", "Sign In".
- Do NOT guess — only use what the signals clearly show.
- Confidence = "high" if JSON-LD or og:site_name confirmed it.
            = "medium" if copyright/logo/title confirms it.
            = "low" if only domain/URL hint used.
- Always populate idn_name — use the domain hint if nothing else is available.

Return ONLY valid JSON (no markdown fences):
{"idn_name": "Name", "confidence": "high|medium|low", "reasoning": "which signal used"}"""

def ask_groq_idn(signal_block: str) -> dict:
    try:
        client = get_groq_client()
        resp   = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": f"Identify the IDN/HCO from these signals:\n\n{signal_block}\n\nReturn ONLY JSON."},
            ],
            max_tokens=150,
            temperature=0.0,
        )
        raw  = resp.choices[0].message.content.strip()
        raw  = re.sub(r'^```(?:json)?\s*', '', raw, flags=re.M)
        raw  = re.sub(r'\s*```$', '', raw, flags=re.M)
        data = json.loads(raw)
        return {
            "idn_name":   data.get("idn_name", "").strip(),
            "confidence": data.get("confidence", "low"),
            "reasoning":  data.get("reasoning", ""),
        }
    except Exception as e:
        if is_rate_limit_error(e):
            rotate_key()
        return {"idn_name": "", "confidence": "low", "reasoning": str(e)}

def get_idn_for_url(url: str) -> dict:
    if not url or str(url).strip() in ("", "nan"):
        return {"idn_name": "", "confidence": "", "reasoning": "no URL"}

    final_url, html = fetch_html(url)

    if not html:
        domain_hint = domain_name_hint(url)
        best_name   = domain_hint or ""
        return {
            "idn_name":   best_name,
            "confidence": "low" if best_name else "none",
            "reasoning":  "page could not load; domain hint used",
        }

    signals = extract_signals(html, final_url or url)
    block   = signals_to_block(signals, final_url or url)
    result  = ask_groq_idn(block)

    if result["confidence"] in ("low", "none") or result["idn_name"] in ("", "UNKNOWN"):
        base           = get_base_url(final_url or url)
        _, about_html  = fetch_about_page(base)
        if about_html:
            about_signals = extract_signals(about_html, base + "/about")
            about_block   = signals_to_block(about_signals, base + "/about")
            combined      = block + "\n\n=== ABOUT PAGE ===\n" + about_block
            result        = ask_groq_idn(combined)

    if not result["idn_name"] or result["idn_name"].upper() == "UNKNOWN":
        fallback = signals.get("domain_hint") or domain_name_hint(url)
        if fallback:
            result["idn_name"]   = fallback
            result["confidence"] = "low"
            result["reasoning"]  = "domain heuristic (all signals exhausted)"

    return result


# ─────────────────────────────────────────────────────────────
# PHASE 3 — BEST IDN RANKING (internal)
# ─────────────────────────────────────────────────────────────
def score_idn_name(name: str) -> int:
    if not name or name.upper() == "UNKNOWN":
        return 0
    score      = 0
    name_lower = name.lower()
    for kw in IDN_KEYWORDS:
        if kw in name_lower:
            score += 2
    if len(name.split()) >= 2:
        score += 1
    if len(name) < 5:
        score -= 3
    return score

def pick_best_idn(site_name: str, candidates: list) -> str:
    valid = [c for c in candidates if c and c.strip() and c.upper() != "UNKNOWN"]
    if not valid:
        return ""
    if len(set(valid)) == 1:
        return valid[0]

    cand_str = "\n".join([f"  {i+1}. {c}" for i, c in enumerate(valid)])
    prompt   = f"""You are a healthcare data expert selecting the most accurate IDN/HCO name.

FACILITY: {site_name}

CANDIDATE NAMES (extracted from up to 4 different URLs for this facility):
{cand_str}

TASK: Pick the single BEST name that represents the PARENT integrated delivery network (IDN)
or healthcare organization (HCO) for this facility.

SELECTION RULES (in priority order):
1. Prefer names containing IDN keywords: health, hospital, medical, system, network,
   care, services, centre, center, institute, foundation, regional, community, mercy,
   memorial, physicians, wellness, rehabilitation, therapy, children, cancer, heart
2. Prefer the most complete/formal name (e.g. "MetroHealth System" over "MetroHealth")
3. Prefer names that are clearly a health SYSTEM over individual clinic names
4. If multiple are equally good, pick the one that appears most often in the list
5. Never pick a generic word alone like "Health" or "Care"

Return ONLY the best name as a plain string — no JSON, no explanation, just the name."""

    try:
        client = get_groq_client()
        resp   = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=30,
            temperature=0.0,
        )
        best = resp.choices[0].message.content.strip().strip('"\'')
        if best and len(best) > 3:
            return best
    except Exception as e:
        if is_rate_limit_error(e):
            rotate_key()

    scored = sorted(valid, key=lambda n: score_idn_name(n), reverse=True)
    return scored[0]


# ─────────────────────────────────────────────────────────────
# FULL PIPELINE RUNNER (all 3 phases chained)
# ─────────────────────────────────────────────────────────────
def run_full_pipeline(df_input, col_map, progress_bar, status_area, log_area):
    NAME  = col_map.get("Site Name", "Site Name")
    ADDR1 = col_map.get("Site Address 1", "")
    CITY  = col_map.get("Site City", "")
    STATE = col_map.get("Site State", "")
    ZIP   = col_map.get("Site Zip", "")

    total   = len(df_input)
    records = df_input.to_dict("records")

    # ── Phase 1: URL Discovery ──────────────────────────────
    log("🚀 Phase 1 — URL Discovery")
    for i, record in enumerate(records):
        name  = str(record.get(NAME, "")).strip()
        addr1 = str(record.get(ADDR1, "")) if ADDR1 and pd.notna(record.get(ADDR1)) else ""
        city  = str(record.get(CITY,  "")) if CITY  and pd.notna(record.get(CITY))  else ""
        state = str(record.get(STATE, "")) if STATE and pd.notna(record.get(STATE)) else ""
        zip_c = str(record.get(ZIP,   "")) if ZIP   and pd.notna(record.get(ZIP))   else ""

        pct = (i / total) * 0.33
        progress_bar.progress(pct, text=f"Phase 1 — Searching {i+1}/{total}: {name[:40]}...")
        status_area.markdown(
            f'<div style="color:#f6ad55;font-size:13px;">🔍 <strong>Phase 1/3</strong> — URL Discovery: <strong>{name}</strong></div>',
            unsafe_allow_html=True,
        )

        log(f"[{i+1}/{total}] 🏥 {name} — {city}, {state}")
        candidates = get_candidates(name, addr1, city, state, zip_c)
        top_urls   = rank_urls(name, addr1, city, state, zip_c, candidates)
        log(f"  🏆 Top URLs: {[u[:60] for u in top_urls if u]}")

        for j, url in enumerate(top_urls, 1):
            record[f"url_{j}"] = url

        recent = st.session_state.log_messages[-8:]
        log_area.code("\n".join(recent), language=None)
        time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

    # ── Phase 2: IDN Extraction ─────────────────────────────
    log("🚀 Phase 2 — IDN Extraction")
    for i, record in enumerate(records):
        site_name = str(record.get(NAME, "")).strip()

        pct = 0.33 + (i / total) * 0.33
        progress_bar.progress(pct, text=f"Phase 2 — Extracting IDN {i+1}/{total}: {site_name[:40]}...")
        status_area.markdown(
            f'<div style="color:#f6ad55;font-size:13px;">🔎 <strong>Phase 2/3</strong> — IDN Extraction: <strong>{site_name}</strong></div>',
            unsafe_allow_html=True,
        )

        log(f"[{i+1}/{total}] 🏥 {site_name}")
        for j in range(1, 5):
            url = str(record.get(f"url_{j}", "")).strip()
            if url and url.lower() != "nan":
                log(f"  ▶ URL {j}: {url[:70]}")
                info = get_idn_for_url(url)
            else:
                info = {"idn_name": "", "confidence": "", "reasoning": "no URL"}

            record[f"idn_name_{j}"]     = info["idn_name"]
            record[f"idn_conf_{j}"]     = info["confidence"]
            record[f"idn_reasoning_{j}"]= info["reasoning"]
            if info["idn_name"]:
                log(f"    🏢 IDN {j}: '{info['idn_name']}' [{info['confidence']}]")
            time.sleep(0.3)

        recent = st.session_state.log_messages[-10:]
        log_area.code("\n".join(recent), language=None)

    # ── Phase 3: Best IDN Ranking ───────────────────────────
    log("🚀 Phase 3 — Best IDN Ranking")
    for i, record in enumerate(records):
        site_name  = str(record.get(NAME, "")).strip()
        candidates = [str(record.get(f"idn_name_{j}", "")).strip() for j in range(1, 5)]

        pct = 0.66 + (i / total) * 0.34
        progress_bar.progress(pct, text=f"Phase 3 — Ranking IDN {i+1}/{total}: {site_name[:40]}...")
        status_area.markdown(
            f'<div style="color:#f6ad55;font-size:13px;">🏆 <strong>Phase 3/3</strong> — Selecting best IDN: <strong>{site_name}</strong></div>',
            unsafe_allow_html=True,
        )

        log(f"[{i+1}/{total}] {site_name} — candidates: {[c for c in candidates if c]}")
        best = pick_best_idn(site_name, candidates)
        record["final_idn"] = best
        log(f"  ✅ Final IDN: '{best}'")

        recent = st.session_state.log_messages[-8:]
        log_area.code("\n".join(recent), language=None)
        time.sleep(0.2)

    progress_bar.progress(1.0, text="✅ Pipeline Complete!")
    return pd.DataFrame(records)


# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="text-align:center;padding:16px 0 8px;">
        <div style="font-size:32px;">🏥</div>
        <div style="color:#e2e8f0;font-size:18px;font-weight:700;margin-top:4px;">IDN Pipeline</div>
        <div style="color:#718096;font-size:12px;">Parent Org Identifier</div>
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    # Pipeline status
    st.markdown("**Pipeline Status**")
    steps = [
        ("File Uploaded",       st.session_state.uploaded_df is not None),
        ("Pipeline Complete",   st.session_state.pipeline_done),
    ]
    for name, done in steps:
        icon  = "✅" if done else "⬜"
        color = "#48bb78" if done else "#718096"
        st.markdown(f'<div style="color:{color};font-size:13px;padding:3px 0;">{icon} {name}</div>', unsafe_allow_html=True)

    st.divider()

    # API Keys
    st.markdown("**🔑 Groq API Keys**")
    key1 = st.text_input("Key 1", type="password", placeholder="gsk_...", key="key1_input")
    key2 = st.text_input("Key 2 (optional)", type="password", placeholder="gsk_...", key="key2_input")

    if st.button("💾 Save API Keys", use_container_width=True):
        keys = [k.strip() for k in [key1, key2] if k.strip()]
        if keys:
            st.session_state.groq_keys = keys
            st.session_state.key_index = 0
            st.success(f"✅ {len(keys)} key(s) saved!")
        else:
            st.error("Please enter at least one API key")

    if st.session_state.groq_keys:
        st.markdown(f'<div style="color:#48bb78;font-size:12px;">✅ {len(st.session_state.groq_keys)} key(s) active</div>', unsafe_allow_html=True)

    st.divider()

    # Settings
    st.markdown("**⚙️ Settings**")
    max_rows  = st.number_input("Max rows to process", min_value=1, max_value=10000, value=200, step=10)
    output_fn = st.text_input("Output filename", value="idn_results.csv")
    st.session_state.output_filename = output_fn

    if st.button("🔄 Reset Pipeline", use_container_width=True):
        for key in ["uploaded_df", "final_df", "pipeline_done", "pipeline_running", "log_messages"]:
            if key in ["pipeline_done", "pipeline_running"]:
                st.session_state[key] = False
            elif key == "log_messages":
                st.session_state[key] = []
            else:
                st.session_state[key] = None
        st.rerun()


# ─────────────────────────────────────────────────────────────
# MAIN CONTENT
# ─────────────────────────────────────────────────────────────
st.markdown("""
<h1 style="color:#e2e8f0;font-size:28px;font-weight:700;margin-bottom:4px;">
    🏥 IDN / HCO Parent Organization Pipeline
</h1>
<p style="color:#718096;font-size:14px;margin-bottom:24px;">
    Upload your facilities file, run the AI pipeline, and download results with identified parent health systems.
</p>
""", unsafe_allow_html=True)


# ─── STEP 1: UPLOAD ─────────────────────────────────────────
with st.expander("📂 Upload Facilities File", expanded=(st.session_state.uploaded_df is None)):
    st.markdown(
        '<p style="color:#718096;font-size:13px;">Upload a CSV or Excel file. '
        'Expected columns: Site Name, Site Address 1, Site City, Site State, Site Zip</p>',
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader(
        "Choose file", type=["csv", "xlsx", "xls"], label_visibility="collapsed"
    )

    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv") else pd.read_excel(uploaded_file)
            df = df.head(max_rows)
            st.session_state.uploaded_df = df

            # Column mapping
            st.markdown("**📋 Column Mapping**")
            cols    = list(df.columns)
            defaults = {
                "Site Name":      next((c for c in cols if "name"  in c.lower()), cols[0]),
                "Site Address 1": next((c for c in cols if "addr"  in c.lower()), cols[min(1, len(cols)-1)]),
                "Site City":      next((c for c in cols if "city"  in c.lower()), cols[min(2, len(cols)-1)]),
                "Site State":     next((c for c in cols if "state" in c.lower()), cols[min(3, len(cols)-1)]),
                "Site Zip":       next((c for c in cols if "zip"   in c.lower() or "postal" in c.lower()), cols[min(4, len(cols)-1)]),
            }

            col1, col2, col3 = st.columns(3)
            col_map = {}
            with col1:
                col_map["Site Name"]      = st.selectbox("Facility Name *", cols, index=cols.index(defaults["Site Name"]) if defaults["Site Name"] in cols else 0)
                col_map["Site Address 1"] = st.selectbox("Address Line 1",  cols, index=cols.index(defaults["Site Address 1"]) if defaults["Site Address 1"] in cols else 0)
            with col2:
                col_map["Site City"]  = st.selectbox("City *",  cols, index=cols.index(defaults["Site City"])  if defaults["Site City"]  in cols else 0)
                col_map["Site State"] = st.selectbox("State *", cols, index=cols.index(defaults["Site State"]) if defaults["Site State"] in cols else 0)
            with col3:
                col_map["Site Zip"] = st.selectbox("Zip Code", cols, index=cols.index(defaults["Site Zip"]) if defaults["Site Zip"] in cols else 0)
                addr2_opts = ["(none)"] + cols
                addr2_sel  = st.selectbox("Address Line 2", addr2_opts, index=0)
                col_map["Site Address 2"] = "" if addr2_sel == "(none)" else addr2_sel

            st.session_state.col_map = col_map

            st.markdown(
                f'<div style="color:#48bb78;font-size:13px;margin:8px 0;">✅ Loaded <strong>{len(df)}</strong> rows × <strong>{len(df.columns)}</strong> columns</div>',
                unsafe_allow_html=True,
            )
            st.markdown("**Preview (first 5 rows):**")
            st.dataframe(df.head(5), use_container_width=True)

        except Exception as e:
            st.error(f"Error reading file: {e}")


# ─── RUN PIPELINE ────────────────────────────────────────────
st.divider()

if st.session_state.uploaded_df is None:
    st.info("⬆️ Upload a file above to get started.")
elif not st.session_state.groq_keys:
    st.warning("🔑 Add your Groq API key(s) in the sidebar, then click Run.")
elif not st.session_state.pipeline_done:
    if st.button("🚀 Run Pipeline", use_container_width=True, type="primary"):
        st.session_state.pipeline_running = True
        log("🚀 Pipeline started")

        progress_bar = st.progress(0, text="Initialising...")
        status_area  = st.empty()
        log_area     = st.empty()

        try:
            final_df = run_full_pipeline(
                st.session_state.uploaded_df.copy(),
                st.session_state.col_map,
                progress_bar,
                status_area,
                log_area,
            )
            st.session_state.final_df        = final_df
            st.session_state.pipeline_done   = True
            st.session_state.pipeline_running = False
            log(f"✅ Pipeline complete — {len(final_df)} rows")
        except Exception as e:
            st.error(f"Pipeline error: {e}")
            log(f"❌ Pipeline error: {e}")
            st.session_state.pipeline_running = False

        st.rerun()


# ─── RESULTS & DOWNLOAD ──────────────────────────────────────
if st.session_state.pipeline_done and st.session_state.final_df is not None:
    df3     = st.session_state.final_df
    col_map = st.session_state.col_map
    NAME    = col_map.get("Site Name", "Site Name")

    filled  = df3["final_idn"].astype(str).str.strip().ne("").sum() if "final_idn" in df3.columns else 0
    pct_cov = 100 * filled // max(len(df3), 1)

    st.markdown("""
    <div style="background:linear-gradient(135deg,#1a2e1a,#1a2e2e);border:2px solid #48bb78;
    border-radius:14px;padding:24px 28px;margin-top:8px;">
        <div style="color:#48bb78;font-size:13px;font-weight:700;text-transform:uppercase;
        letter-spacing:0.8px;margin-bottom:12px;">✅ Pipeline Complete — Results Ready</div>
    """, unsafe_allow_html=True)

    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown(f'<div class="metric-box"><div class="metric-value">{len(df3)}</div><div class="metric-label">Total Rows</div></div>', unsafe_allow_html=True)
    with m2:
        st.markdown(f'<div class="metric-box"><div class="metric-value">{int(filled)}</div><div class="metric-label">IDNs Identified</div></div>', unsafe_allow_html=True)
    with m3:
        st.markdown(f'<div class="metric-box"><div class="metric-value">{pct_cov}%</div><div class="metric-label">Coverage</div></div>', unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("**🏆 Final Results — Best IDN per Facility:**")
    if "final_idn" in df3.columns:
        idn_cols_present = [f"idn_name_{j}" for j in range(1, 5) if f"idn_name_{j}" in df3.columns]
        show = df3[[NAME, "final_idn"] + idn_cols_present].copy()
        show.columns = [NAME, "✅ Final IDN"] + [f"IDN Candidate {j}" for j in range(1, len(idn_cols_present) + 1)]
        st.dataframe(show, use_container_width=True)

    # Build output columns — interleaved url/idn pairs then final_idn
    base_cols    = [col_map.get(c, c) for c in ["Site Name", "Site Address 1", "Site Address 2", "Site City", "Site State", "Site Zip"]]
    interleaved  = [col for j in range(1, 5) for col in (f"url_{j}", f"idn_name_{j}")]
    keep         = [c for c in base_cols + interleaved + ["final_idn"] if c in df3.columns]
    out_df    = df3[keep]

    st.markdown(f'<div style="color:#718096;font-size:12px;margin:8px 0;">{len(out_df)} rows · {len(keep)} columns</div>', unsafe_allow_html=True)

    csv_out = out_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        f"⬇️ Download {st.session_state.output_filename}",
        csv_out,
        st.session_state.output_filename,
        "text/csv",
        use_container_width=True,
        type="primary",
    )


# ─── ACTIVITY LOG ────────────────────────────────────────────
st.divider()
with st.expander("📋 Activity Log", expanded=False):
    if st.session_state.log_messages:
        st.code("\n".join(st.session_state.log_messages[-60:]), language=None)
        if st.button("🗑️ Clear Log"):
            st.session_state.log_messages = []
            st.rerun()
    else:
        st.markdown('<p style="color:#718096;">No activity yet.</p>', unsafe_allow_html=True)
