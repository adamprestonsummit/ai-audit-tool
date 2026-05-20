"""
Summit AI Visibility Audit Tool v2
Powered by Google Gemini 2.5 Flash
"""

import streamlit as st
import requests
from bs4 import BeautifulSoup
import json
import re
import base64
import io
import os
import time
from datetime import datetime
from urllib.parse import urlparse

# --- Page config must be first Streamlit call ---
st.set_page_config(
    page_title="Summit AI Visibility Audit",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

import plotly.graph_objects as go
import google.genai as genai
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from fpdf import FPDF
from PIL import Image, ImageDraw, ImageFont

# =============================================================================
# SUMMIT BRAND PALETTE  (extracted from summit.co.uk + logo)
# =============================================================================
# Primary brand colours
S_RED       = "#D0380B"   # Logo red-orange (dominant)
S_CHARCOAL  = "#3E405B"   # Site theme-color (dark blue-grey)
S_WHITE     = "#FFFFFF"
S_OFFWHITE  = "#F7F7F8"   # Page background tint
S_LIGHT     = "#EDEDF0"   # Card border / divider

# Text colours
S_INK       = "#1C1C2E"   # Near-black body text
S_MUTED     = "#6B6B80"   # Muted / secondary text
S_CAPTION   = "#9999AA"   # Caption

# Score band colours — backgrounds and foregrounds
BG_RED      = "#FDE9E5"
FG_RED      = "#B52D0A"
BG_AMBER    = "#FFF4E0"
FG_AMBER    = "#8A4800"
BG_GREEN    = "#E6F5EC"
FG_GREEN    = "#1E6E3C"

# RGB tuples for docx / fpdf
S_RED_RGB       = (208, 56, 11)
S_CHARCOAL_RGB  = (62, 64, 91)
S_WHITE_RGB     = (255, 255, 255)
FG_RED_RGB      = (181, 45, 10)
FG_AMBER_RGB    = (138, 72, 0)
FG_GREEN_RGB    = (30, 110, 60)
BG_RED_HEX      = "FDE9E5"
BG_AMBER_HEX    = "FFF4E0"
BG_GREEN_HEX    = "E6F5EC"

# Dimension configuration — weighted by AI impact
DIMENSION_CONFIG = {
    "Crawlability & Bot Access":    {"weight": 0.20, "icon": "🤖", "color": S_CHARCOAL,  "description": "Can AI crawlers access and parse the page content?"},
    "Structured Data / Schema":     {"weight": 0.18, "icon": "🏷️",  "color": "#7C3AED",  "description": "Schema.org JSON-LD implementation quality and completeness"},
    "LLM Content Signals":          {"weight": 0.15, "icon": "🧠",  "color": S_RED,      "description": "Content clarity, factual density and E-E-A-T signals"},
    "Meta & SEO Signals":           {"weight": 0.12, "icon": "🔍",  "color": "#0E7C4A",  "description": "Title tags, meta descriptions, canonical tags and Open Graph"},
    "Heading Structure":            {"weight": 0.10, "icon": "📋",  "color": "#0369A1",  "description": "Semantic H1–H6 hierarchy and logical content flow"},
    "ARIA Implementation":          {"weight": 0.10, "icon": "♿",  "color": "#BE185D",  "description": "Accessibility attributes that also aid AI content parsing"},
    "Link Quality":                 {"weight": 0.08, "icon": "🔗",  "color": "#78350F",  "description": "Internal/external link quality, anchor text and nofollow usage"},
    "Image Alt Text":               {"weight": 0.07, "icon": "🖼️",  "color": "#374151",  "description": "Alt text coverage, quality and keyword relevance"},
    "AI Search Health":             {"weight": 0.05, "icon": "📡",  "color": "#B45309",  "description": "llms.txt, AI bot directives in robots.txt"},
    "Duplicate Content & Tags":     {"weight": 0.05, "icon": "📄",  "color": "#6B6B80",  "description": "Duplicate titles, canonical consistency, thin content"},
}


# =============================================================================
# LOGO HELPERS
# =============================================================================
LOGO_PATH = os.path.join(os.path.dirname(__file__), "summit_logo.png")

def load_logo_bytes() -> bytes:
    """Load the Summit logo PNG from disk."""
    with open(LOGO_PATH, "rb") as f:
        return f.read()

def logo_b64(logo_bytes: bytes) -> str:
    return base64.b64encode(logo_bytes).decode()


# =============================================================================
# CSS
# =============================================================================
def inject_css():
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
    .main {{ background: {S_OFFWHITE}; }}

    /* ---- Sidebar ---- */
    [data-testid="stSidebar"] {{ background: {S_CHARCOAL} !important; }}
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] div,
    [data-testid="stSidebar"] small {{ color: rgba(255,255,255,0.90) !important; }}
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {{ color: #FFFFFF !important; }}
    [data-testid="stSidebar"] .stTextInput input {{
        background: rgba(255,255,255,0.10) !important;
        color: #FFFFFF !important;
        border: 1px solid rgba(255,255,255,0.25) !important;
        border-radius: 6px;
    }}
    [data-testid="stSidebar"] .stTextInput input::placeholder {{ color: rgba(255,255,255,0.45) !important; }}
    [data-testid="stSidebar"] .stButton > button {{
        background: {S_RED} !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        border: none !important;
        border-radius: 6px !important;
        width: 100%;
        padding: 0.65rem 1rem;
        font-size: 0.95rem;
        letter-spacing: 0.02em;
    }}
    [data-testid="stSidebar"] .stButton > button:hover {{ background: #B52D0A !important; }}
    [data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,0.15) !important; }}

    /* ---- Metric cards ---- */
    .metric-card {{
        background: {S_WHITE};
        border-radius: 12px;
        padding: 18px 22px;
        box-shadow: 0 1px 6px rgba(0,0,0,0.08);
        margin-bottom: 14px;
        border-top: 4px solid {S_RED};
    }}
    .metric-card.red   {{ border-top-color: {FG_RED}; background: {BG_RED}; }}
    .metric-card.amber {{ border-top-color: {FG_AMBER}; background: {BG_AMBER}; }}
    .metric-card.green {{ border-top-color: {FG_GREEN}; background: {BG_GREEN}; }}
    .metric-number {{ font-size: 2.6rem; font-weight: 800; line-height: 1.05; color: {S_INK}; }}
    .metric-label  {{ font-size: 0.75rem; font-weight: 600; text-transform: uppercase;
                      letter-spacing: 0.06em; color: {S_MUTED}; margin-top: 4px; }}

    /* ---- Score hero ---- */
    .score-hero {{
        background: linear-gradient(135deg, {S_CHARCOAL} 0%, #52547A 100%);
        border-radius: 14px;
        padding: 28px 32px;
        color: {S_WHITE};
        text-align: center;
        box-shadow: 0 4px 18px rgba(62,64,91,0.35);
    }}
    .score-hero .number {{ font-size: 4.8rem; font-weight: 800; line-height: 1; color: {S_WHITE}; }}
    .score-hero .denom  {{ font-size: 1.5rem; color: rgba(255,255,255,0.55); }}
    .score-hero .lbl    {{ font-size: 0.78rem; font-weight: 600; text-transform: uppercase;
                           letter-spacing: 0.08em; color: rgba(255,255,255,0.65); margin-bottom: 8px; }}

    /* ---- Issues list ---- */
    .issue-row {{
        background: {S_WHITE};
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 8px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
        border-left: 4px solid {S_RED};
        display: flex; align-items: flex-start; gap: 12px;
    }}
    .issue-row.critical {{ border-left-color: {FG_RED}; }}
    .issue-row.warning  {{ border-left-color: {FG_AMBER}; }}
    .issue-row.info     {{ border-left-color: #1D5FA6; }}

    /* ---- Badges ---- */
    .badge-red   {{ background: {BG_RED};   color: {FG_RED};   padding: 3px 10px; border-radius: 12px; font-weight: 600; font-size: 0.8rem; white-space: nowrap; }}
    .badge-amber {{ background: {BG_AMBER}; color: {FG_AMBER}; padding: 3px 10px; border-radius: 12px; font-weight: 600; font-size: 0.8rem; white-space: nowrap; }}
    .badge-green {{ background: {BG_GREEN}; color: {FG_GREEN}; padding: 3px 10px; border-radius: 12px; font-weight: 600; font-size: 0.8rem; white-space: nowrap; }}

    /* ---- Table ---- */
    .audit-table {{ width:100%; border-collapse:collapse; background:{S_WHITE};
                    border-radius:10px; overflow:hidden; box-shadow:0 2px 8px rgba(0,0,0,0.07); }}
    .audit-table th {{ background:{S_CHARCOAL}; color:{S_WHITE}; padding:11px 14px;
                       text-align:left; font-size:0.78rem; text-transform:uppercase; letter-spacing:0.05em; }}
    .audit-table td {{ padding:10px 14px; border-bottom:1px solid {S_LIGHT}; font-size:0.86rem; color:{S_INK}; }}
    .audit-table tr:last-child td {{ border-bottom:none; }}
    .audit-table tr:hover td {{ background:{S_OFFWHITE}; }}

    /* ---- Misc ---- */
    #MainMenu {{ visibility: hidden; }}
    footer {{ visibility: hidden; }}
    .stDeployButton {{ display: none; }}
    </style>
    """, unsafe_allow_html=True)


# =============================================================================
# SCORING HELPERS
# =============================================================================
def score_band(score: float) -> str:
    if score <= 2:   return "red"
    if score <= 5:   return "amber"
    return "green"

def score_fg(score: float) -> str:
    return {" red": FG_RED, "red": FG_RED, "amber": FG_AMBER, "green": FG_GREEN}[score_band(score)]

def score_bg(score: float) -> str:
    return {"red": BG_RED, "amber": BG_AMBER, "green": BG_GREEN}[score_band(score)]

def score_fg_hex(score: float) -> str:
    return {"red": FG_RED, "amber": FG_AMBER, "green": FG_GREEN}[score_band(score)]

def score_fg_rgb(score: float) -> tuple:
    return {"red": FG_RED_RGB, "amber": FG_AMBER_RGB, "green": FG_GREEN_RGB}[score_band(score)]

def score_bg_hex(score: float) -> str:
    return {"red": BG_RED_HEX, "amber": BG_AMBER_HEX, "green": BG_GREEN_HEX}[score_band(score)]

def weighted_overall(scores: dict) -> float:
    total_w, total_ws = 0, 0
    for dim, score in scores.items():
        w = DIMENSION_CONFIG.get(dim, {}).get("weight", 0.05)
        total_ws += score * w
        total_w  += w
    return round(total_ws / total_w, 1) if total_w else 0.0


# =============================================================================
# PAGE FETCHER
# =============================================================================
def fetch_page(url: str) -> dict:
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; SummitAuditBot/2.0; +https://summit.co.uk)",
        "Accept": "text/html,application/xhtml+xml",
        "Accept-Language": "en-GB,en;q=0.9",
    }
    result = {
        "url": url, "status_code": None, "html": "", "text": "",
        "error": None, "load_time": None, "is_https": url.startswith("https://"),
        "redirect_chain": [], "html_raw_length": 0, "text_length": 0,
        "text_to_html_ratio": 0,
    }
    try:
        t0 = time.time()
        r = requests.get(url, headers=headers, timeout=15, allow_redirects=True)
        result["load_time"]     = round(time.time() - t0, 2)
        result["status_code"]   = r.status_code
        result["html"]          = r.text
        result["final_url"]     = r.url
        result["redirect_chain"]= [resp.url for resp in r.history]
        result["response_headers"] = dict(r.headers)
        soup = BeautifulSoup(r.text, "html.parser")
        for tag in soup(["script", "style", "noscript"]):
            tag.decompose()
        result["text"]              = soup.get_text(" ", strip=True)[:8000]
        result["html_raw_length"]   = len(r.text)
        result["text_length"]       = len(result["text"])
        result["text_to_html_ratio"]= round(result["text_length"] / max(len(r.text), 1), 3)
    except Exception as e:
        result["error"] = str(e)
    return result


def check_robots(url: str) -> dict:
    parsed  = urlparse(url)
    rob_url = f"{parsed.scheme}://{parsed.netloc}/robots.txt"
    ai_bots = ["GPTBot","ChatGPT-User","OAI-SearchBot","Google-Extended",
               "Googlebot","PerplexityBot","ClaudeBot","anthropic-ai",
               "Amazonbot","YouBot","CCBot"]
    result  = {"robots_url": rob_url, "found": False, "raw": "", "ai_bots": {}}
    try:
        r = requests.get(rob_url, timeout=8, headers={"User-Agent":"Mozilla/5.0"})
        if r.status_code == 200:
            result["found"] = True
            result["raw"]   = r.text[:3000]
            text = r.text.lower()
            for bot in ai_bots:
                pat   = rf"user-agent:\s*{re.escape(bot.lower())}(.*?)(?=user-agent:|$)"
                match = re.search(pat, text, re.DOTALL | re.IGNORECASE)
                blocked = False
                if match:
                    block = match.group(1)
                    if re.search(r"disallow:\s*/\s*$", block, re.MULTILINE):
                        blocked = True
                    elif re.search(r"disallow:\s*$", block, re.MULTILINE):
                        blocked = False
                result["ai_bots"][bot] = "BLOCKED" if blocked else "ALLOWED"
    except Exception as e:
        result["error"] = str(e)
    return result


def check_llms_txt(url: str) -> dict:
    parsed   = urlparse(url)
    llms_url = f"{parsed.scheme}://{parsed.netloc}/llms.txt"
    result   = {"found": False, "url": llms_url, "content": ""}
    try:
        r = requests.get(llms_url, timeout=8, headers={"User-Agent":"Mozilla/5.0"})
        if r.status_code == 200 and len(r.text) > 10:
            result["found"]   = True
            result["content"] = r.text[:1000]
    except Exception:
        pass
    return result


# =============================================================================
# GEMINI ANALYSIS
# =============================================================================
def analyse_dimension(client, page_data: dict, dimension: str) -> dict:
    html_snip = page_data.get("html", "")[:6000]
    txt_snip  = page_data.get("text", "")[:3000]

    prompts = {
        "ARIA Implementation": f"""Analyse the ARIA implementation on this webpage for AI visibility.

HTML (first 6000 chars):
{html_snip}

Evaluate: ARIA roles/labels/landmarks, form label associations, aria-live regions, structural clarity for AI parsers.

Return ONLY a JSON object — no markdown fences:
{{"score": <1-10>, "summary": "<2 sentences>", "findings": ["<str>"], "issues": [{{"severity": "critical|warning|info", "issue": "<str>", "recommendation": "<str>"}}], "positive": ["<str>"]}}""",

        "Structured Data / Schema": f"""Analyse the schema.org structured data on this webpage.

HTML:
{html_snip}

Evaluate: JSON-LD types present, completeness, missing recommended types (Product, BreadcrumbList, FAQPage, Organization, etc.), required field population.

Return ONLY a JSON object — no markdown fences:
{{"score": <1-10>, "summary": "<2 sentences>", "schemas_found": ["<str>"], "schemas_missing": ["<str>"], "findings": ["<str>"], "issues": [{{"severity": "critical|warning|info", "issue": "<str>", "recommendation": "<str>"}}], "positive": ["<str>"]}}""",

        "Heading Structure": f"""Analyse the heading structure (H1-H6) on this webpage.

HTML:
{html_snip}
Text:
{txt_snip}

Evaluate: H1 count, hierarchy logic, skipped levels, heading descriptiveness, content structure clarity for AI.

Return ONLY a JSON object — no markdown fences:
{{"score": <1-10>, "summary": "<2 sentences>", "h1_count": <int>, "findings": ["<str>"], "issues": [{{"severity": "critical|warning|info", "issue": "<str>", "recommendation": "<str>"}}], "positive": ["<str>"]}}""",

        "Meta & SEO Signals": f"""Analyse meta tags and SEO signals on this webpage.

HTML:
{html_snip}

Evaluate: title tag (presence, length 50-60 chars), meta description (presence, length 150-160 chars), canonical tag, Open Graph tags, Twitter card, robots meta tag.

Return ONLY a JSON object — no markdown fences:
{{"score": <1-10>, "summary": "<2 sentences>", "title_length": <int>, "meta_desc_present": true/false, "canonical_present": true/false, "og_present": true/false, "findings": ["<str>"], "issues": [{{"severity": "critical|warning|info", "issue": "<str>", "recommendation": "<str>"}}], "positive": ["<str>"]}}""",

        "Link Quality": f"""Analyse link quality on this webpage for AI visibility.

HTML:
{html_snip}

Evaluate: descriptive vs generic anchor text ("click here" etc.), external link authority, nofollow usage, image links with no text, anchor text helping AI content graph.

Return ONLY a JSON object — no markdown fences:
{{"score": <1-10>, "summary": "<2 sentences>", "findings": ["<str>"], "issues": [{{"severity": "critical|warning|info", "issue": "<str>", "recommendation": "<str>"}}], "positive": ["<str>"]}}""",

        "Image Alt Text": f"""Analyse image alt text quality on this webpage.

HTML:
{html_snip}

Evaluate: percentage of images with alt text, alt text quality (descriptive, not stuffed), decorative images with empty alt="", complex images needing longer descriptions.

Return ONLY a JSON object — no markdown fences:
{{"score": <1-10>, "summary": "<2 sentences>", "images_found": <int>, "images_with_alt": <int>, "findings": ["<str>"], "issues": [{{"severity": "critical|warning|info", "issue": "<str>", "recommendation": "<str>"}}], "positive": ["<str>"]}}""",

        "Crawlability & Bot Access": f"""Analyse crawlability and bot access for this webpage.

HTML length: {page_data.get('html_raw_length', 0)} chars
Text extracted: {page_data.get('text_length', 0)} chars
Text-to-HTML ratio: {page_data.get('text_to_html_ratio', 0):.3f}
Status code: {page_data.get('status_code')}
HTTPS: {page_data.get('is_https')}
Load time: {page_data.get('load_time')}s
Redirects: {page_data.get('redirect_chain', [])}
Robots.txt data: {json.dumps(page_data.get('robots_data', {}))}
HTML snippet: {html_snip[:2000]}

Evaluate: SSR vs CSR rendering (low text-to-HTML ratio < 0.05 suggests JS dependency), bot blocking, HTTPS, redirect chains, load time, AI bot access in robots.txt.

Return ONLY a JSON object — no markdown fences:
{{"score": <1-10>, "summary": "<2 sentences>", "rendering_type": "SSR|CSR|Mixed", "js_dependent": true/false, "findings": ["<str>"], "issues": [{{"severity": "critical|warning|info", "issue": "<str>", "recommendation": "<str>"}}], "positive": ["<str>"]}}""",

        "LLM Content Signals": f"""Analyse content quality and LLM/AI visibility signals on this webpage.

Text:
{txt_snip}
HTML snippet:
{html_snip[:2000]}

Evaluate: E-E-A-T signals (expertise, experience, authoritativeness, trustworthiness), content depth and factual density, clarity for LLM extraction, trust signals (author names, dates, credentials), brand/entity clarity, FAQ-style extractable content.

Return ONLY a JSON object — no markdown fences:
{{"score": <1-10>, "summary": "<2 sentences>", "eeat_signals": ["<str>"], "findings": ["<str>"], "issues": [{{"severity": "critical|warning|info", "issue": "<str>", "recommendation": "<str>"}}], "positive": ["<str>"]}}""",

        "AI Search Health": f"""Analyse AI search health signals for this webpage.

llms.txt found: {page_data.get('llms_txt', {}).get('found', False)}
llms.txt content: {page_data.get('llms_txt', {}).get('content', 'N/A')}
Robots.txt AI bots: {json.dumps(page_data.get('robots_data', {}).get('ai_bots', {}))}
HTML: {html_snip[:2000]}

Evaluate: llms.txt presence and quality, AI crawler access (GPTBot, ChatGPT-User, OAI-SearchBot, ClaudeBot etc.) in robots.txt, any AI-specific meta tags or directives.

Return ONLY a JSON object — no markdown fences:
{{"score": <1-10>, "summary": "<2 sentences>", "llms_txt_present": true/false, "ai_bots_blocked": ["<str>"], "findings": ["<str>"], "issues": [{{"severity": "critical|warning|info", "issue": "<str>", "recommendation": "<str>"}}], "positive": ["<str>"]}}""",

        "Duplicate Content & Tags": f"""Analyse duplicate content and tag issues on this webpage.

HTML:
{html_snip}
Text-to-HTML ratio: {page_data.get('text_to_html_ratio', 0):.3f}

Evaluate: canonical tag presence, duplicate/default CMS title risk, boilerplate content proportion, thin content signals, duplicate meta description risk.

Return ONLY a JSON object — no markdown fences:
{{"score": <1-10>, "summary": "<2 sentences>", "canonical_present": true/false, "findings": ["<str>"], "issues": [{{"severity": "critical|warning|info", "issue": "<str>", "recommendation": "<str>"}}], "positive": ["<str>"]}}"""
    }

    prompt = prompts.get(dimension,
        f'Analyse {dimension}. Return ONLY JSON: {{"score": 5, "summary": "", "findings": [], "issues": [], "positive": []}}')

    try:
        resp = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        raw  = resp.text.strip()
        raw  = re.sub(r"^```(?:json)?\n?", "", raw)
        raw  = re.sub(r"\n?```$", "", raw)
        return json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r'\{.*\}', raw if 'raw' in dir() else '{}', re.DOTALL)
        if m:
            try: return json.loads(m.group(0))
            except Exception: pass
    except Exception:
        pass
    return {"score": 0, "summary": "Analysis could not be completed.", "findings": [], "issues": [], "positive": []}


def gen_exec_summary(client, url: str, scores: dict, all_results: dict) -> str:
    overall      = weighted_overall(scores)
    score_lines  = "\n".join(f"- {d}: {s}/10" for d, s in scores.items())
    critical_iss = [f"{d}: {i['issue']}" for d, r in all_results.items()
                    for i in r.get("issues", []) if i.get("severity") == "critical"][:8]

    prompt = f"""You are an expert AI visibility consultant at Summit, a performance marketing agency based in Hull.

Write a professional executive summary for an AI visibility technical audit of {url}.

Overall weighted score: {overall}/10
Dimension scores:
{score_lines}

Critical issues:
{chr(10).join(critical_iss)}

Write 3–4 concise, professional paragraphs covering:
1. Overall assessment and score context
2. Key strengths identified
3. Priority areas for improvement
4. Business impact — why this matters for AI search visibility in 2025/26

Tone: consultancy-grade, direct, no bullet points. UK English."""

    try:
        resp = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
        return resp.text.strip()
    except Exception as e:
        return f"Executive summary could not be generated: {e}"


# =============================================================================
# PLOTLY CHARTS
# =============================================================================
def gauge_chart(overall: float) -> go.Figure:
    color = score_fg_hex(overall)
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=overall,
        number={"suffix": "/10", "font": {"size": 40, "color": color, "family": "Inter"}},
        gauge={
            "axis": {"range": [0, 10], "tickwidth": 1, "tickcolor": S_MUTED,
                     "tickfont": {"size": 10, "color": S_MUTED}},
            "bar":  {"color": color, "thickness": 0.28},
            "bgcolor": S_WHITE,
            "borderwidth": 0,
            "steps": [
                {"range": [0,   3.3], "color": BG_RED},
                {"range": [3.3, 6.6], "color": BG_AMBER},
                {"range": [6.6, 10],  "color": BG_GREEN},
            ],
        }
    ))
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=20, b=10), height=210,
        font=dict(family="Inter"),
    )
    return fig


def bar_chart(scores: dict) -> go.Figure:
    dims   = list(scores.keys())
    vals   = [scores[d] for d in dims]
    colors = [score_fg_hex(v) for v in vals]
    labels = [d.replace(" & ", " &\n").replace(" / ", "/\n") for d in dims]

    fig = go.Figure(go.Bar(
        x=vals, y=labels, orientation="h",
        marker_color=colors,
        text=[f"  {v}/10" for v in vals],
        textposition="outside",
        textfont=dict(size=11, color=S_INK),
        hovertemplate="%{y}: %{x}/10<extra></extra>",
    ))
    fig.add_vline(x=3.3, line_dash="dot", line_color=FG_RED,   opacity=0.5,
                  annotation_text="Critical", annotation_font=dict(size=9, color=FG_RED))
    fig.add_vline(x=6.6, line_dash="dot", line_color=FG_GREEN, opacity=0.5,
                  annotation_text="Good", annotation_font=dict(size=9, color=FG_GREEN))
    fig.update_layout(
        xaxis=dict(range=[0, 12.5], showgrid=True, gridcolor=S_LIGHT, title="Score / 10"),
        yaxis=dict(autorange="reversed", tickfont=dict(size=10, color=S_INK)),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10, r=70, t=20, b=30), height=400,
        font=dict(family="Inter"), showlegend=False,
    )
    return fig


def radar_chart(scores: dict) -> go.Figure:
    dims = list(scores.keys())
    vals = [scores[d] for d in dims]
    fig  = go.Figure(go.Scatterpolar(
        r=vals + [vals[0]], theta=dims + [dims[0]],
        fill="toself",
        fillcolor=f"rgba(208,56,11,0.12)",
        line=dict(color=S_RED, width=2.5),
        marker=dict(size=6, color=S_RED),
    ))
    fig.update_layout(
        polar=dict(
            bgcolor=S_WHITE,
            radialaxis=dict(visible=True, range=[0,10], tickfont=dict(size=8), gridcolor=S_LIGHT),
            angularaxis=dict(tickfont=dict(size=9, color=S_INK)),
        ),
        showlegend=False,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=55, r=55, t=30, b=30), height=400,
        font=dict(family="Inter"),
    )
    return fig


def weight_donut(scores: dict) -> go.Figure:
    dims    = list(DIMENSION_CONFIG.keys())
    weights = [DIMENSION_CONFIG[d]["weight"] * 100 for d in dims]
    colors  = [DIMENSION_CONFIG[d]["color"] for d in dims]
    fig = go.Figure(go.Pie(
        labels=dims, values=weights, hole=0.52,
        marker=dict(colors=colors),
        textinfo="percent",
        textfont=dict(size=9),
        hovertemplate="%{label}<br>Weight: %{value:.0f}%<extra></extra>",
    ))
    fig.update_layout(
        showlegend=True,
        legend=dict(font=dict(size=9, color=S_INK), orientation="v", x=1.02),
        paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=5, r=5, t=20, b=10), height=360,
        font=dict(family="Inter"),
    )
    return fig


# =============================================================================
# DOCX EXPORT
# =============================================================================
def _set_cell_bg(cell, hex6: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex6)
    tcPr.append(shd)


def _heading_para(doc, text: str, size: int = 14):
    p  = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot  = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single"); bot.set(qn("w:sz"), "8")
    bot.set(qn("w:space"), "4");    bot.set(qn("w:color"), "D0380B")
    pBdr.append(bot); pPr.append(pBdr)
    r = p.add_run(text)
    r.bold = True; r.font.size = Pt(size)
    r.font.color.rgb = RGBColor(*S_CHARCOAL_RGB)
    r.font.name = "Arial"
    return p


def build_docx(url: str, scores: dict, all_results: dict,
               exec_summary: str, logo_bytes: bytes) -> bytes:
    doc = Document()
    for sec in doc.sections:
        sec.top_margin    = Inches(0.7)
        sec.bottom_margin = Inches(0.7)
        sec.left_margin   = Inches(0.8)
        sec.right_margin  = Inches(0.8)

    # Styles
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(10)

    # — Logo —
    logo_para = doc.add_paragraph()
    logo_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    logo_para.add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(1.4))

    # — Title —
    _heading_para(doc, "AI Visibility Audit Report", size=24)
    meta = doc.add_paragraph()
    for txt, col in [(f"{url}  |  ", (120,120,130)), (datetime.now().strftime("%B %Y"), (120,120,130))]:
        r = meta.add_run(txt); r.font.size = Pt(9); r.font.color.rgb = RGBColor(*col)
    doc.add_paragraph()

    # — Overall score table —
    overall = weighted_overall(scores)
    _heading_para(doc, "Overall AI Visibility Score")

    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = "Table Grid"; tbl.autofit = False
    widths = [1800, 1800, 6200]
    cells  = tbl.rows[0].cells

    # Cell 0 — big score
    cells[0].width = Emu(widths[0] * 914)
    _set_cell_bg(cells[0], "3E405B")
    p = cells[0].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(str(overall)); r.bold = True; r.font.size = Pt(42); r.font.color.rgb = RGBColor(255,255,255)
    p2 = cells[0].add_paragraph("/10"); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.runs[0].font.color.rgb = RGBColor(200,200,210); p2.runs[0].font.size = Pt(14)

    # Cell 1 — rating label
    cells[1].width = Emu(widths[1] * 914)
    band = score_band(overall)
    bg_hex_map = {"red": BG_RED_HEX, "amber": BG_AMBER_HEX, "green": BG_GREEN_HEX}
    fg_rgb_map  = {"red": FG_RED_RGB, "amber": FG_AMBER_RGB, "green": FG_GREEN_RGB}
    lbl_map     = {"red": "Needs Attention", "amber": "Developing", "green": "Good"}
    _set_cell_bg(cells[1], bg_hex_map[band])
    p = cells[1].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(lbl_map[band]); r.bold = True; r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(*fg_rgb_map[band])

    # Cell 2 — dimension breakdown
    cells[2].width = Emu(widths[2] * 914)
    _set_cell_bg(cells[2], "F7F7F8")
    p = cells[2].paragraphs[0]
    r = p.add_run("Score breakdown"); r.bold = True; r.font.size = Pt(9); r.font.color.rgb = RGBColor(*S_CHARCOAL_RGB)
    for dim, s in sorted(scores.items(), key=lambda x: x[1]):
        p2  = cells[2].add_paragraph()
        r1  = p2.add_run(f"{dim}: "); r1.font.size = Pt(8.5)
        r2  = p2.add_run(f"{s}/10"); r2.bold = True; r2.font.size = Pt(8.5)
        r2.font.color.rgb = RGBColor(*fg_rgb_map[score_band(s)])
    doc.add_paragraph()

    # — Executive Summary —
    _heading_para(doc, "Executive Summary")
    for para_txt in exec_summary.split("\n\n"):
        if para_txt.strip():
            p = doc.add_paragraph(para_txt.strip())
            p.style.font.size = Pt(10)
    doc.add_paragraph()

    # — Dimension scores table —
    _heading_para(doc, "Dimension Scores")
    dt = doc.add_table(rows=1, cols=4); dt.style = "Table Grid"; dt.autofit = False
    dw = [3800, 900, 1000, 4100]
    for i, (cell, hdr) in enumerate(zip(dt.rows[0].cells, ["Dimension","Weight","Score","Rating"])):
        cell.width = Emu(dw[i] * 914); _set_cell_bg(cell, "3E405B")
        r = cell.paragraphs[0].add_run(hdr); r.bold = True
        r.font.color.rgb = RGBColor(255,255,255); r.font.size = Pt(9)
    for dim, s in scores.items():
        row = dt.add_row(); band_d = score_band(s)
        w_pct = int(DIMENSION_CONFIG.get(dim,{}).get("weight",0.05)*100)
        row.cells[0].width = Emu(dw[0]*914); row.cells[1].width = Emu(dw[1]*914)
        row.cells[2].width = Emu(dw[2]*914); row.cells[3].width = Emu(dw[3]*914)
        row.cells[0].paragraphs[0].add_run(dim).font.size = Pt(9)
        row.cells[1].paragraphs[0].add_run(f"{w_pct}%").font.size = Pt(9)
        _set_cell_bg(row.cells[2], bg_hex_map[band_d])
        sr = row.cells[2].paragraphs[0].add_run(f"{s}/10")
        sr.bold = True; sr.font.size = Pt(10); sr.font.color.rgb = RGBColor(*fg_rgb_map[band_d])
        _set_cell_bg(row.cells[3], bg_hex_map[band_d])
        lbl_strs = {"red":"❌  Needs Attention","amber":"⚠  Needs Work","green":"✓  Good"}
        lr = row.cells[3].paragraphs[0].add_run(lbl_strs[band_d])
        lr.font.size = Pt(9); lr.font.color.rgb = RGBColor(*fg_rgb_map[band_d])
    doc.add_paragraph()

    # — Detailed findings —
    _heading_para(doc, "Detailed Findings by Dimension")
    sev_bg  = {"critical": BG_RED_HEX, "warning": BG_AMBER_HEX, "info": "EBF0FB"}
    sev_fg  = {"critical": FG_RED_RGB, "warning": FG_AMBER_RGB, "info": (21, 80, 175)}
    for dim, result in all_results.items():
        s    = scores.get(dim, 0); band_d = score_band(s)
        icon = DIMENSION_CONFIG.get(dim, {}).get("icon", "•")
        dp   = doc.add_paragraph()
        r1   = dp.add_run(f"{icon}  {dim}   "); r1.bold = True; r1.font.size = Pt(12)
        r1.font.color.rgb = RGBColor(*S_CHARCOAL_RGB)
        r2   = dp.add_run(f"[{s}/10]"); r2.bold = True; r2.font.size = Pt(12)
        r2.font.color.rgb = RGBColor(*fg_rgb_map[band_d])
        if result.get("summary"):
            sp = doc.add_paragraph(result["summary"])
            sp.runs[0].italic = True; sp.runs[0].font.size = Pt(9.5)
        issues = result.get("issues", [])
        if issues:
            it = doc.add_table(rows=1, cols=3); it.style = "Table Grid"; it.autofit = False
            iw = [1000, 3500, 5300]
            for i, (cell, hdr) in enumerate(zip(it.rows[0].cells, ["Severity","Issue","Recommendation"])):
                cell.width = Emu(iw[i]*914); _set_cell_bg(cell, "3E405B")
                r = cell.paragraphs[0].add_run(hdr); r.bold = True
                r.font.color.rgb = RGBColor(255,255,255); r.font.size = Pt(8.5)
            for iss in issues:
                sev = iss.get("severity","info")
                row = it.add_row()
                row.cells[0].width = Emu(iw[0]*914)
                row.cells[1].width = Emu(iw[1]*914)
                row.cells[2].width = Emu(iw[2]*914)
                _set_cell_bg(row.cells[0], sev_bg.get(sev,"F5F5F5"))
                sr = row.cells[0].paragraphs[0].add_run(sev.upper())
                sr.bold = True; sr.font.size = Pt(8); sr.font.color.rgb = RGBColor(*sev_fg.get(sev,(80,80,80)))
                row.cells[1].paragraphs[0].add_run(iss.get("issue","")).font.size = Pt(8.5)
                row.cells[2].paragraphs[0].add_run(iss.get("recommendation","")).font.size = Pt(8.5)
        positives = result.get("positive", [])
        if positives:
            pp = doc.add_paragraph(); r = pp.add_run("✓ Strengths:  ")
            r.bold = True; r.font.size = Pt(9); r.font.color.rgb = RGBColor(*FG_GREEN_RGB)
            r2 = pp.add_run("  |  ".join(positives[:3]))
            r2.font.size = Pt(9); r2.font.color.rgb = RGBColor(*FG_GREEN_RGB)
        doc.add_paragraph()

    # — Priority recommendations —
    _heading_para(doc, "Priority Recommendations")
    all_recs = []
    for dim, result in all_results.items():
        w = DIMENSION_CONFIG.get(dim,{}).get("weight",0.05)
        for iss in result.get("issues",[]):
            sev = iss.get("severity","info")
            all_recs.append({"priority": {"critical":1,"warning":2,"info":3}.get(sev,3),
                             "weight": w, "dim": dim,
                             "issue": iss.get("issue",""), "rec": iss.get("recommendation",""), "sev": sev})
    all_recs.sort(key=lambda x: (x["priority"], -x["weight"]))
    rt = doc.add_table(rows=1, cols=4); rt.style = "Table Grid"; rt.autofit = False
    rw = [600, 2000, 3400, 3800]
    for i, (cell, hdr) in enumerate(zip(rt.rows[0].cells, ["#","Dimension","Issue","Recommendation"])):
        cell.width = Emu(rw[i]*914); _set_cell_bg(cell, "3E405B")
        r = cell.paragraphs[0].add_run(hdr); r.bold = True
        r.font.color.rgb = RGBColor(255,255,255); r.font.size = Pt(9)
    for rank, rec in enumerate(all_recs[:15], 1):
        row = rt.add_row(); sev = rec["sev"]
        row.cells[0].width = Emu(rw[0]*914); row.cells[1].width = Emu(rw[1]*914)
        row.cells[2].width = Emu(rw[2]*914); row.cells[3].width = Emu(rw[3]*914)
        _set_cell_bg(row.cells[0], sev_bg.get(sev,"F5F5F5"))
        nr = row.cells[0].paragraphs[0].add_run(str(rank)); nr.bold = True; nr.font.size = Pt(9)
        nr.font.color.rgb = RGBColor(*sev_fg.get(sev,(80,80,80)))
        row.cells[1].paragraphs[0].add_run(rec["dim"]).font.size = Pt(8.5)
        row.cells[2].paragraphs[0].add_run(rec["issue"]).font.size = Pt(8.5)
        row.cells[3].paragraphs[0].add_run(rec["rec"]).font.size = Pt(8.5)
    doc.add_paragraph()

    # — Footer —
    fp = doc.add_paragraph(); fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = fp._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    top  = OxmlElement("w:top"); top.set(qn("w:val"),"single"); top.set(qn("w:sz"),"4")
    top.set(qn("w:space"),"4"); top.set(qn("w:color"),"D0380B")
    pBdr.append(top); pPr.append(pBdr)
    r = fp.add_run("Prepared by Summit  |  AI Visibility Practice  |  summit.co.uk")
    r.font.size = Pt(8); r.italic = True; r.font.color.rgb = RGBColor(130,130,140)

    buf = io.BytesIO()
    doc.save(buf); buf.seek(0)
    return buf.getvalue()


# =============================================================================
# PDF EXPORT
# =============================================================================
def build_pdf(url: str, scores: dict, exec_summary: str, logo_bytes: bytes) -> bytes:
    logo_tmp = "/tmp/summit_logo_pdf.png"
    with open(logo_tmp, "wb") as f:
        f.write(logo_bytes)

    overall = weighted_overall(scores)
    band    = score_band(overall)
    bg_map_rgb  = {"red":(253,233,229), "amber":(255,244,224), "green":(230,245,236)}
    fg_map_rgb2 = {"red": FG_RED_RGB, "amber": FG_AMBER_RGB, "green": FG_GREEN_RGB}
    lbl_map     = {"red":"Needs Attention","amber":"Developing","green":"Good"}

    class SummitPDF(FPDF):
        def header(self):
            self.image(logo_tmp, 12, 8, 24)
            self.set_font("Helvetica", "B", 16)
            self.set_text_color(*S_CHARCOAL_RGB)
            self.set_xy(40, 10)
            self.cell(0, 8, "AI Visibility Audit Report", ln=True)
            self.set_font("Helvetica", "", 9)
            self.set_text_color(120, 120, 130)
            self.set_x(40)
            self.cell(0, 6, f"{url}  |  {datetime.now().strftime('%B %Y')}", ln=True)
            self.ln(3)
            self.set_draw_color(*S_RED_RGB)
            self.set_line_width(1.0)
            self.line(12, 28, 198, 28)
            self.ln(4)

        def footer(self):
            self.set_y(-12)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(150, 150, 160)
            self.cell(0, 6, f"Summit AI Visibility Audit  |  Confidential  |  Page {self.page_no()}", align="C")

    pdf = SummitPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # Overall score block
    y0 = pdf.get_y()
    pdf.set_fill_color(*S_CHARCOAL_RGB)
    pdf.rect(12, y0, 88, 38, "F")
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 34)
    pdf.set_xy(12, y0 + 4)
    pdf.cell(88, 16, f"{overall}/10", align="C", ln=False)
    pdf.set_xy(12, y0 + 22)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(200, 200, 215)
    pdf.cell(88, 10, "Overall AI Visibility Score", align="C", ln=False)

    pdf.set_fill_color(*bg_map_rgb[band])
    pdf.rect(104, y0, 90, 38, "F")
    pdf.set_text_color(*fg_map_rgb2[band])
    pdf.set_font("Helvetica", "B", 17)
    pdf.set_xy(104, y0 + 8)
    pdf.cell(90, 12, lbl_map[band], align="C", ln=False)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_xy(104, y0 + 22)
    pdf.cell(90, 10, f"Weighted across {len(scores)} dimensions", align="C")
    pdf.ln(y0 + 46 - pdf.get_y())

    # Scores table
    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 11); pdf.set_text_color(*S_CHARCOAL_RGB)
    pdf.cell(0, 7, "Dimension Scores", ln=True)
    pdf.set_draw_color(*S_RED_RGB); pdf.set_line_width(0.7)
    pdf.line(12, pdf.get_y(), 198, pdf.get_y()); pdf.ln(3)

    pdf.set_fill_color(*S_CHARCOAL_RGB); pdf.set_text_color(255,255,255)
    pdf.set_font("Helvetica","B",8.5)
    pdf.cell(88,6.5,"Dimension",fill=True,border=0)
    pdf.cell(18,6.5,"Weight",fill=True,border=0,align="C")
    pdf.cell(18,6.5,"Score",fill=True,border=0,align="C")
    pdf.cell(58,6.5,"Rating",fill=True,border=0); pdf.ln()

    sev_lbl_pdf = {"red":"Needs Attention","amber":"Needs Work","green":"Good"}
    for i, (dim, s) in enumerate(sorted(scores.items(), key=lambda x: x[1])):
        bd   = score_band(s); w_pct = int(DIMENSION_CONFIG.get(dim,{}).get("weight",0.05)*100)
        fill = bg_map_rgb[bd] if i % 2 == 0 else (250,250,252)
        pdf.set_fill_color(*fill); pdf.set_text_color(40,40,50)
        pdf.set_font("Helvetica","",8.5)
        pdf.cell(88,6,"  "+dim,fill=True,border=0)
        pdf.cell(18,6,f"{w_pct}%",fill=True,border=0,align="C")
        pdf.set_fill_color(*bg_map_rgb[bd]); pdf.set_text_color(*fg_map_rgb2[bd])
        pdf.set_font("Helvetica","B",9)
        pdf.cell(18,6,f"{s}/10",fill=True,border=0,align="C")
        pdf.set_font("Helvetica","",8.5)
        pdf.cell(58,6,f"  {sev_lbl_pdf[bd]}",fill=True,border=0); pdf.ln()

    pdf.ln(6)
    pdf.set_font("Helvetica","B",11); pdf.set_text_color(*S_CHARCOAL_RGB)
    pdf.cell(0,7,"Executive Summary",ln=True)
    pdf.set_draw_color(*S_RED_RGB); pdf.set_line_width(0.7)
    pdf.line(12, pdf.get_y(), 198, pdf.get_y()); pdf.ln(3)
    pdf.set_font("Helvetica","",9.5); pdf.set_text_color(40,40,50)
    for para in exec_summary.split("\n\n"):
        if para.strip():
            pdf.multi_cell(0, 5.5, para.strip()); pdf.ln(2)

    raw = pdf.output()
    return bytes(raw) if not isinstance(raw, bytes) else raw


# =============================================================================
# MAIN
# =============================================================================
def main():
    inject_css()

    # Load real Summit logo
    logo_bytes = load_logo_bytes()
    logo_b64_str = logo_b64(logo_bytes)

    # — Sidebar —
    with st.sidebar:
        st.markdown(
            f'<img src="data:image/png;base64,{logo_b64_str}" '
            f'width="100" style="margin-bottom:20px;display:block">',
            unsafe_allow_html=True
        )
        st.markdown("## AI Visibility Audit")
        st.markdown("---")

        # API key: prefer Streamlit secrets, fall back to text input
        api_key = ""
        try:
            api_key = st.secrets["GEMINI_API_KEY"]
            st.success("✅ API key loaded from secrets", icon=None)
        except Exception:
            api_key = st.text_input(
                "Gemini API Key",
                type="password",
                help="Or add GEMINI_API_KEY to .streamlit/secrets.toml"
            )

        url_input = st.text_input(
            "URL to Audit",
            placeholder="https://example.com",
            help="Full URL including https://"
        )

        st.markdown("---")
        st.markdown("**Top dimensions by weight**")
        for dim, cfg in list(DIMENSION_CONFIG.items())[:5]:
            st.caption(f"{cfg['icon']} {dim}: {int(cfg['weight']*100)}%")

        st.markdown("---")
        run_audit = st.button("🚀 Run Audit", use_container_width=True)

        st.markdown("---")
        st.caption("Summit AI Visibility Audit v2")
        st.caption("Powered by Gemini 2.5 Flash")
        st.caption("© Summit Performance Marketing")

    # — Header —
    hcol1, hcol2 = st.columns([3, 1])
    with hcol1:
        st.markdown(
            f'<div style="display:flex;align-items:center;gap:16px;margin-bottom:6px">'
            f'<img src="data:image/png;base64,{logo_b64_str}" width="64" style="border-radius:8px">'
            f'<div>'
            f'<h1 style="margin:0;color:{S_CHARCOAL};font-size:1.75rem;font-weight:800">AI Visibility Audit Dashboard</h1>'
            f'<p style="margin:0;color:{S_MUTED};font-size:0.88rem">Technical audit for AI crawler access &amp; content extraction quality</p>'
            f'</div></div>',
            unsafe_allow_html=True
        )
    with hcol2:
        st.markdown(
            f'<p style="text-align:right;color:{S_CAPTION};margin-top:22px;font-size:0.82rem">'
            f'{datetime.now().strftime("%d %b %Y")}</p>',
            unsafe_allow_html=True
        )
    st.markdown("---")

    # — Landing state —
    if not run_audit:
        st.markdown(f"""
        <div style="background:{S_WHITE};border-radius:12px;padding:36px;text-align:center;
        box-shadow:0 2px 8px rgba(0,0,0,0.07);margin-bottom:28px">
            <h2 style="color:{S_CHARCOAL};margin-top:0">Enter a URL in the sidebar to begin</h2>
            <p style="color:{S_MUTED};max-width:480px;margin:12px auto 0">
            Audits 10 dimensions weighted by AI search impact. Generates a Semrush-style
            dashboard plus branded Word and PDF exports.</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### Audit Dimensions")
        cols = st.columns(2)
        for i, (dim, cfg) in enumerate(DIMENSION_CONFIG.items()):
            with cols[i % 2]:
                st.markdown(f"""
                <div style="background:{S_WHITE};border-radius:8px;padding:13px 16px;
                margin-bottom:9px;box-shadow:0 1px 4px rgba(0,0,0,0.05);
                border-left:4px solid {cfg['color']}">
                    <div style="display:flex;justify-content:space-between;align-items:center">
                        <span style="font-weight:600;font-size:0.9rem;color:{S_CHARCOAL}">{cfg['icon']} {dim}</span>
                        <span style="background:{S_OFFWHITE};color:{S_CHARCOAL};padding:2px 8px;
                        border-radius:10px;font-size:0.76rem;font-weight:700">{int(cfg['weight']*100)}%</span>
                    </div>
                    <p style="margin:4px 0 0;color:{S_MUTED};font-size:0.8rem">{cfg['description']}</p>
                </div>
                """, unsafe_allow_html=True)
        return

    # — Validation —
    if not api_key:
        st.error("Please enter your Gemini API key in the sidebar (or add it to .streamlit/secrets.toml).")
        return
    if not url_input:
        st.error("Please enter a URL to audit.")
        return

    url = url_input.strip()
    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    # — Run audit —
    st.markdown(f"### Auditing `{url}`")
    prog  = st.progress(0)
    stat  = st.empty()

    try:
        client = genai.Client(api_key=api_key)
    except Exception as e:
        st.error(f"Gemini API error: {e}"); return

    stat.text("📡 Fetching page...")
    prog.progress(5)
    page_data = fetch_page(url)
    if page_data.get("error"):
        st.error(f"Failed to fetch page: {page_data['error']}"); return

    stat.text("🤖 Checking robots.txt and llms.txt...")
    prog.progress(10)
    page_data["robots_data"] = check_robots(url)
    page_data["llms_txt"]    = check_llms_txt(url)

    scores, all_results = {}, {}
    dims = list(DIMENSION_CONFIG.keys())
    for i, dim in enumerate(dims):
        stat.text(f"🔍 Analysing: {dim}...")
        prog.progress(15 + int(68 * i / len(dims)))
        res = analyse_dimension(client, page_data, dim)
        scores[dim]       = res.get("score", 0)
        all_results[dim]  = res
        time.sleep(0.25)

    stat.text("✍️ Generating executive summary...")
    prog.progress(88)
    exec_summary = gen_exec_summary(client, url, scores, all_results)

    prog.progress(100); time.sleep(0.4)
    stat.empty(); prog.empty()

    overall = weighted_overall(scores)
    band    = score_band(overall)
    bg_map  = {"red": BG_RED, "amber": BG_AMBER, "green": BG_GREEN}
    fg_map  = {"red": FG_RED, "amber": FG_AMBER, "green": FG_GREEN}
    lbl_map = {"red": "Needs Attention", "amber": "Developing", "green": "Good"}

    # ── Dashboard ──────────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns([1.6, 1.4, 1, 1])

    with c1:
        st.markdown(f"""
        <div class="score-hero">
            <div class="lbl">Overall AI Visibility Score</div>
            <div>
                <span class="number">{overall}</span>
                <span class="denom"> /10</span>
            </div>
            <div style="margin-top:10px;background:rgba(255,255,255,0.12);border-radius:8px;
                        padding:6px 18px;display:inline-block;font-weight:700;font-size:1rem">
                {lbl_map[band]}
            </div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.plotly_chart(gauge_chart(overall), use_container_width=True,
                        config={"displayModeBar": False})

    with c3:
        n_crit = sum(1 for r in all_results.values()
                     for iss in r.get("issues",[]) if iss.get("severity")=="critical")
        n_warn = sum(1 for r in all_results.values()
                     for iss in r.get("issues",[]) if iss.get("severity")=="warning")
        st.markdown(f"""
        <div class="metric-card red" style="text-align:center">
            <div class="metric-number" style="color:{FG_RED}">{n_crit}</div>
            <div class="metric-label">Critical Issues</div>
        </div>
        <div class="metric-card amber" style="text-align:center">
            <div class="metric-number" style="color:{FG_AMBER}">{n_warn}</div>
            <div class="metric-label">Warnings</div>
        </div>
        """, unsafe_allow_html=True)

    with c4:
        n_pass = sum(1 for s in scores.values() if s > 6.5)
        lt     = page_data.get("load_time","—")
        st.markdown(f"""
        <div class="metric-card green" style="text-align:center">
            <div class="metric-number" style="color:{FG_GREEN}">{n_pass}</div>
            <div class="metric-label">Dimensions Passing</div>
        </div>
        <div class="metric-card" style="text-align:center">
            <div class="metric-number" style="color:{S_CHARCOAL}">{lt}s</div>
            <div class="metric-label">Page Load Time</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    # Tabs
    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["📊 Scores", "🕸️ Radar", "⚖️ Weighting", "📋 Issues", "📝 Summary"])

    with tab1:
        ba, bb = st.columns([2, 1])
        with ba:
            st.markdown("#### Score by Dimension")
            st.plotly_chart(bar_chart(scores), use_container_width=True,
                            config={"displayModeBar": False})
        with bb:
            st.markdown("#### Breakdown")
            rows = ""
            for dim, s in sorted(scores.items(), key=lambda x: x[1]):
                bc = score_band(s); icon = DIMENSION_CONFIG.get(dim,{}).get("icon","•")
                rows += f"""<tr>
                    <td>{icon} {dim}</td>
                    <td style="text-align:right"><span class="badge-{bc}">{s}/10</span></td>
                </tr>"""
            st.markdown(f'<table class="audit-table"><thead><tr>'
                        f'<th>Dimension</th><th>Score</th></tr></thead>'
                        f'<tbody>{rows}</tbody></table>', unsafe_allow_html=True)

    with tab2:
        st.markdown("#### AI Visibility Radar Profile")
        st.plotly_chart(radar_chart(scores), use_container_width=True,
                        config={"displayModeBar": False})

    with tab3:
        wa, wb = st.columns(2)
        with wa:
            st.markdown("#### Dimension Weighting")
            st.plotly_chart(weight_donut(scores), use_container_width=True,
                            config={"displayModeBar": False})
        with wb:
            st.markdown("#### Weight & Score Table")
            rows = ""
            for dim, cfg in DIMENSION_CONFIG.items():
                s = scores.get(dim, 0); bc = score_band(s)
                ws = round(s * cfg["weight"], 2)
                rows += f"""<tr>
                    <td>{cfg['icon']} {dim}</td>
                    <td style="text-align:center">{int(cfg['weight']*100)}%</td>
                    <td style="text-align:center"><span class="badge-{bc}">{s}/10</span></td>
                    <td style="text-align:center;font-weight:600;color:{S_CHARCOAL}">{ws}</td>
                </tr>"""
            st.markdown(f'<table class="audit-table"><thead><tr>'
                        f'<th>Dimension</th><th>Weight</th><th>Score</th><th>Weighted</th>'
                        f'</tr></thead><tbody>{rows}</tbody></table>',
                        unsafe_allow_html=True)

    with tab4:
        st.markdown("#### Issues")
        sev_filter = st.selectbox("Filter by severity", ["All","Critical","Warning","Info"])
        flat_issues = sorted(
            [{"sev": i.get("severity","info"), "dim": d, **i}
             for d, r in all_results.items() for i in r.get("issues",[])],
            key=lambda x: {"critical":0,"warning":1,"info":2}.get(x["sev"],3)
        )
        shown = 0
        for iss in flat_issues:
            if sev_filter != "All" and iss["sev"].lower() != sev_filter.lower():
                continue
            shown += 1
            icons = {"critical":"🔴","warning":"🟡","info":"🔵"}
            icon  = icons.get(iss["sev"],"⚪")
            st.markdown(f"""
            <div class="issue-row {iss['sev']}">
                <div style="min-width:24px;font-size:1rem">{icon}</div>
                <div style="flex:1">
                    <div style="display:flex;align-items:center;gap:8px;margin-bottom:3px">
                        <span style="font-weight:600;font-size:0.87rem;color:{S_INK}">{iss.get('issue','')}</span>
                        <span style="background:{S_LIGHT};color:{S_MUTED};padding:1px 7px;
                              border-radius:10px;font-size:0.74rem">{iss['dim']}</span>
                    </div>
                    <p style="margin:0;color:{S_MUTED};font-size:0.82rem">
                        💡 {iss.get('recommendation','')}</p>
                </div>
            </div>
            """, unsafe_allow_html=True)
        if shown == 0:
            st.info("No issues found for this filter.")

    with tab5:
        st.markdown("#### Executive Summary")
        st.markdown(f"""
        <div style="background:{S_WHITE};border-radius:12px;padding:26px;
        box-shadow:0 2px 8px rgba(0,0,0,0.07);color:{S_INK};line-height:1.65;font-size:0.93rem">
            {exec_summary.replace(chr(10), "<br>")}
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    # Detailed accordion
    st.markdown("### 🔎 Findings by Dimension")
    for dim, result in all_results.items():
        s    = scores.get(dim, 0); bd = score_band(s)
        icon = DIMENSION_CONFIG.get(dim,{}).get("icon","•")
        with st.expander(f"{icon} {dim}  —  {s}/10", expanded=(s <= 3)):
            da, db = st.columns([1, 2])
            with da:
                st.markdown(f"""
                <div style="text-align:center;padding:18px;background:{bg_map[bd]};
                border-radius:10px;border:1px solid {S_LIGHT}">
                    <div style="font-size:2.8rem;font-weight:800;color:{fg_map[bd]}">{s}</div>
                    <div style="color:{fg_map[bd]};font-size:0.76rem;font-weight:600;
                    text-transform:uppercase;letter-spacing:0.05em">/10 — {lbl_map[bd]}</div>
                </div>
                <p style="text-align:center;color:{S_CAPTION};font-size:0.78rem;margin-top:8px">
                Weight: {int(DIMENSION_CONFIG.get(dim,{}).get('weight',0.05)*100)}%</p>
                """, unsafe_allow_html=True)
            with db:
                if result.get("summary"):
                    st.markdown(f"**{result['summary']}**")
                for f_txt in result.get("findings", []):
                    st.markdown(f"- {f_txt}")
            issues = result.get("issues", [])
            if issues:
                st.markdown("**Issues**")
                for iss in issues:
                    sev = iss.get("severity","info")
                    ico = {"critical":"🔴","warning":"🟡","info":"🔵"}.get(sev,"⚪")
                    st.markdown(f"{ico} **{iss.get('issue','')}**")
                    st.caption(f"Recommendation: {iss.get('recommendation','')}")
            positives = result.get("positive",[])
            if positives:
                st.markdown("**Strengths**")
                for p in positives:
                    st.markdown(f"✅ {p}")

    st.markdown("---")

    # ── Exports ────────────────────────────────────────────────────────────
    st.markdown("### 📥 Export Reports")
    e1, e2, e3 = st.columns(3)

    domain = urlparse(url).netloc.replace("www.","")
    stamp  = datetime.now().strftime("%Y%m")

    with e1:
        st.markdown(f"""
        <div style="background:{S_WHITE};border-radius:10px;padding:18px;
        box-shadow:0 2px 6px rgba(0,0,0,0.07);border-top:4px solid {S_CHARCOAL}">
            <h4 style="color:{S_CHARCOAL};margin-top:0">📄 Word Document</h4>
            <p style="color:{S_MUTED};font-size:0.83rem">Full Summit-branded audit report. Includes executive summary, dimension scores, detailed findings and priority recommendations.</p>
        </div>
        """, unsafe_allow_html=True)
        with st.spinner("Building .docx..."):
            docx_bytes = build_docx(url, scores, all_results, exec_summary, logo_bytes)
        st.download_button("⬇️ Download Word Report", docx_bytes,
                           file_name=f"summit_ai_audit_{domain}_{stamp}.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           use_container_width=True)

    with e2:
        st.markdown(f"""
        <div style="background:{S_WHITE};border-radius:10px;padding:18px;
        box-shadow:0 2px 6px rgba(0,0,0,0.07);border-top:4px solid {S_RED}">
            <h4 style="color:{S_RED};margin-top:0">📋 PDF Summary</h4>
            <p style="color:{S_MUTED};font-size:0.83rem">Dashboard summary as PDF. Includes overall score, dimension breakdown and executive summary. Ideal for client presentations.</p>
        </div>
        """, unsafe_allow_html=True)
        with st.spinner("Building PDF..."):
            pdf_bytes = build_pdf(url, scores, exec_summary, logo_bytes)
        st.download_button("⬇️ Download PDF", pdf_bytes,
                           file_name=f"summit_ai_audit_{domain}_{stamp}.pdf",
                           mime="application/pdf",
                           use_container_width=True)

    with e3:
        st.markdown(f"""
        <div style="background:{S_WHITE};border-radius:10px;padding:18px;
        box-shadow:0 2px 6px rgba(0,0,0,0.07);border-top:4px solid {FG_GREEN}">
            <h4 style="color:{FG_GREEN};margin-top:0">📊 JSON Data</h4>
            <p style="color:{S_MUTED};font-size:0.83rem">Raw audit data including all scores, findings and analysis output. Useful for integration with other reporting tools or CRMs.</p>
        </div>
        """, unsafe_allow_html=True)
        export = {"url": url, "audit_date": datetime.now().isoformat(),
                  "overall_score": overall, "scores": scores,
                  "results": {k: {kk:vv for kk,vv in v.items() if kk!="soup"}
                               for k,v in all_results.items()},
                  "executive_summary": exec_summary}
        st.download_button("⬇️ Download JSON", json.dumps(export, indent=2, default=str),
                           file_name=f"summit_ai_audit_{domain}_{stamp}.json",
                           mime="application/json",
                           use_container_width=True)


if __name__ == "__main__":
    main()
