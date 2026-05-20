"""
Summit AI Visibility Audit Tool v2
Powered by Google Gemini
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
from urllib.parse import urlparse, urljoin
import urllib.robotparser

# --- Page config first ---
st.set_page_config(
    page_title="Summit AI Visibility Audit",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Imports that require packages ---
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import google.genai as genai
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Emu, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from fpdf import FPDF
from PIL import Image, ImageDraw, ImageFont


# =============================================================================
# SUMMIT BRAND CONSTANTS
# =============================================================================
SUMMIT_ORANGE = "#C25200"
SUMMIT_NAVY = "#1A3055"
SUMMIT_RED_SCORE = "#C0392B"
SUMMIT_AMBER_SCORE = "#7B3800"
SUMMIT_GREEN_SCORE = "#276749"
SUMMIT_BG_RED = "#FDECEA"
SUMMIT_BG_AMBER = "#FFF8E1"
SUMMIT_BG_GREEN = "#EBF7EF"
SUMMIT_GREY = "#F5F5F5"
SUMMIT_DARK = "#1A1A1A"
SUMMIT_MID_GREY = "#666666"

# Score weight configuration — highest impact first
DIMENSION_CONFIG = {
    "Crawlability & Bot Access": {
        "weight": 0.20,
        "icon": "🤖",
        "description": "Can AI crawlers access and parse the page content?",
        "color": "#2196F3"
    },
    "Structured Data / Schema": {
        "weight": 0.18,
        "icon": "🏷️",
        "description": "Schema.org JSON-LD implementation quality and completeness",
        "color": "#9C27B0"
    },
    "LLM Content Signals": {
        "weight": 0.15,
        "icon": "🧠",
        "description": "Content clarity, factual density and E-E-A-T signals",
        "color": "#FF9800"
    },
    "Meta & SEO Signals": {
        "weight": 0.12,
        "icon": "🔍",
        "description": "Title tags, meta descriptions, canonical tags and Open Graph",
        "color": "#4CAF50"
    },
    "Heading Structure": {
        "weight": 0.10,
        "icon": "📋",
        "description": "Semantic H1-H6 hierarchy and logical content flow",
        "color": "#00BCD4"
    },
    "ARIA Implementation": {
        "weight": 0.10,
        "icon": "♿",
        "description": "Accessibility attributes that also aid AI content parsing",
        "color": "#E91E63"
    },
    "Link Quality": {
        "weight": 0.08,
        "icon": "🔗",
        "description": "Internal and external link quality, anchor text and no-follow usage",
        "color": "#795548"
    },
    "Image Alt Text": {
        "weight": 0.07,
        "icon": "🖼️",
        "description": "Alt text coverage, quality and keyword relevance",
        "color": "#607D8B"
    },
    "AI Search Health": {
        "weight": 0.05,
        "icon": "🤖",
        "description": "llms.txt presence, AI-specific directives and bot accessibility",
        "color": "#FF5722"
    },
    "Duplicate Content & Tags": {
        "weight": 0.05,
        "icon": "📄",
        "description": "Duplicate title tags, content duplication and canonical consistency",
        "color": "#9E9E9E"
    }
}

TOTAL_WEIGHTS = sum(d["weight"] for d in DIMENSION_CONFIG.values())


# =============================================================================
# LOGO GENERATION
# =============================================================================
def create_summit_logo(width=280, height=70):
    """Generate Summit branded logo as PIL Image."""
    img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    # Orange-red background
    draw.rounded_rectangle([0, 0, width - 1, height - 1], radius=6, fill=(194, 82, 0))
    try:
        font_bold = ImageFont.truetype("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf", 32)
        font_small = ImageFont.truetype("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", 11)
    except Exception:
        font_bold = ImageFont.load_default()
        font_small = font_bold
    draw.text((18, 10), "SUMMIT", fill="white", font=font_bold)
    draw.text((18, 48), "Performance Marketing", fill=(255, 200, 160), font=font_small)
    return img


def logo_to_base64(img: Image.Image) -> str:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()


def logo_to_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# =============================================================================
# CUSTOM CSS
# =============================================================================
def inject_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    .main { background: #F8F9FA; }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: #1A3055 !important;
    }
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    [data-testid="stSidebar"] .stTextInput input {
        background: rgba(255,255,255,0.1) !important;
        color: white !important;
        border: 1px solid rgba(255,255,255,0.3) !important;
        border-radius: 6px;
    }
    [data-testid="stSidebar"] .stButton button {
        background: #C25200 !important;
        color: white !important;
        font-weight: 600 !important;
        border: none !important;
        border-radius: 6px !important;
        width: 100%;
        padding: 0.6rem 1rem;
        font-size: 1rem;
    }

    /* Score cards */
    .score-card {
        background: white;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        margin-bottom: 16px;
        border-top: 4px solid #C25200;
    }
    .score-card-red { border-top-color: #C0392B !important; background: #FDECEA; }
    .score-card-amber { border-top-color: #E65100 !important; background: #FFF8E1; }
    .score-card-green { border-top-color: #276749 !important; background: #EBF7EF; }

    .metric-big {
        font-size: 3rem;
        font-weight: 700;
        line-height: 1;
    }
    .metric-label {
        font-size: 0.8rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #666;
        margin-top: 4px;
    }

    /* Issue rows */
    .issue-row {
        background: white;
        border-radius: 8px;
        padding: 14px 18px;
        margin-bottom: 8px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
        border-left: 4px solid #C25200;
        display: flex;
        align-items: flex-start;
        gap: 12px;
    }
    .issue-row.critical { border-left-color: #C0392B; }
    .issue-row.warning { border-left-color: #E65100; }
    .issue-row.info { border-left-color: #1565C0; }

    /* Overall score hero */
    .score-hero {
        background: linear-gradient(135deg, #1A3055 0%, #2C4E7A 100%);
        border-radius: 16px;
        padding: 32px;
        color: white;
        text-align: center;
        margin-bottom: 24px;
        box-shadow: 0 4px 20px rgba(26,48,85,0.3);
    }
    .score-hero h1 { color: white; font-size: 5rem; margin: 0; font-weight: 800; }
    .score-hero p { color: rgba(255,255,255,0.8); margin: 0; }

    /* Section headers */
    .section-header {
        background: #1A3055;
        color: white;
        padding: 10px 18px;
        border-radius: 8px;
        font-weight: 600;
        font-size: 1rem;
        margin-bottom: 16px;
        margin-top: 24px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    /* Semrush-style table */
    .audit-table {
        width: 100%;
        border-collapse: collapse;
        background: white;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }
    .audit-table th {
        background: #1A3055;
        color: white;
        padding: 12px 16px;
        text-align: left;
        font-size: 0.82rem;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }
    .audit-table td {
        padding: 11px 16px;
        border-bottom: 1px solid #f0f0f0;
        font-size: 0.88rem;
    }
    .audit-table tr:last-child td { border-bottom: none; }
    .audit-table tr:hover td { background: #f9f9f9; }

    /* Score badge */
    .badge-red { background: #FDECEA; color: #C0392B; padding: 3px 10px; border-radius: 12px; font-weight: 600; font-size: 0.82rem; }
    .badge-amber { background: #FFF8E1; color: #7B3800; padding: 3px 10px; border-radius: 12px; font-weight: 600; font-size: 0.82rem; }
    .badge-green { background: #EBF7EF; color: #276749; padding: 3px 10px; border-radius: 12px; font-weight: 600; font-size: 0.82rem; }

    /* Hide streamlit default elements */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    .stDeployButton { display: none; }
    </style>
    """, unsafe_allow_html=True)


# =============================================================================
# SCORING HELPERS
# =============================================================================
def score_color_class(score: float) -> str:
    if score <= 2:
        return "red"
    elif score <= 5:
        return "amber"
    else:
        return "green"


def score_hex(score: float) -> str:
    if score <= 2:
        return SUMMIT_RED_SCORE
    elif score <= 5:
        return SUMMIT_AMBER_SCORE
    else:
        return SUMMIT_GREEN_SCORE


def score_bg_hex(score: float) -> str:
    if score <= 2:
        return SUMMIT_BG_RED
    elif score <= 5:
        return SUMMIT_BG_AMBER
    else:
        return SUMMIT_BG_GREEN


def weighted_overall_score(scores: dict) -> float:
    """Compute weighted overall score (0-10) from dimension scores."""
    total_weight = 0
    weighted_sum = 0
    for dim, score in scores.items():
        w = DIMENSION_CONFIG.get(dim, {}).get("weight", 0.05)
        weighted_sum += score * w
        total_weight += w
    if total_weight == 0:
        return 0
    return round(weighted_sum / total_weight, 1)


# =============================================================================
# PAGE FETCHER
# =============================================================================
def fetch_page(url: str, timeout: int = 15) -> dict:
    """Fetch a URL and return content + metadata."""
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; SummitAuditBot/2.0; +https://summit.co.uk)",
        "Accept": "text/html,application/xhtml+xml,application/xhtml+xml",
        "Accept-Language": "en-GB,en;q=0.9",
    }
    result = {
        "url": url,
        "status_code": None,
        "html": "",
        "text": "",
        "error": None,
        "load_time": None,
        "headers": {},
        "is_https": url.startswith("https://"),
    }
    try:
        start = time.time()
        r = requests.get(url, headers=headers, timeout=timeout, allow_redirects=True)
        result["load_time"] = round(time.time() - start, 2)
        result["status_code"] = r.status_code
        result["headers"] = dict(r.headers)
        result["html"] = r.text
        result["final_url"] = r.url
        result["redirect_chain"] = [resp.url for resp in r.history]

        soup = BeautifulSoup(r.text, "html.parser")
        # Remove scripts and styles for clean text
        for tag in soup(["script", "style", "noscript"]):
            tag.decompose()
        result["text"] = soup.get_text(separator=" ", strip=True)[:8000]
        result["soup"] = soup
        result["html_raw_length"] = len(r.text)
        result["text_length"] = len(result["text"])
        result["text_to_html_ratio"] = round(result["text_length"] / max(result["html_raw_length"], 1), 3)
    except Exception as e:
        result["error"] = str(e)
    return result


def check_robots_txt(url: str) -> dict:
    """Check robots.txt for AI bot directives."""
    parsed = urlparse(url)
    robots_url = f"{parsed.scheme}://{parsed.netloc}/robots.txt"
    result = {
        "robots_url": robots_url,
        "found": False,
        "raw": "",
        "ai_bots": {},
        "disallowed_paths": [],
    }
    ai_bots = [
        "GPTBot", "ChatGPT-User", "OAI-SearchBot", "Google-Extended",
        "Googlebot", "PerplexityBot", "ClaudeBot", "anthropic-ai",
        "Amazonbot", "YouBot", "CCBot", "ia_archiver"
    ]
    try:
        r = requests.get(robots_url, timeout=8, headers={"User-Agent": "Mozilla/5.0"})
        if r.status_code == 200:
            result["found"] = True
            result["raw"] = r.text[:3000]
            text = r.text.lower()
            for bot in ai_bots:
                blocked = False
                # Look for disallow rules for this bot
                pattern = rf"user-agent:\s*{re.escape(bot.lower())}.*?(?=user-agent:|$)"
                match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
                if match:
                    block = match.group(0)
                    blocked = "disallow: /" in block and "disallow: /\n" not in block.replace("disallow: /\n", "OK")
                    # More precise: check if disallow: / with no path means full block
                    if re.search(r"disallow:\s*/\s*$", block, re.MULTILINE):
                        blocked = True
                    elif re.search(r"disallow:\s*$", block, re.MULTILINE):
                        blocked = False  # empty disallow = allow all
                result["ai_bots"][bot] = "BLOCKED" if blocked else "ALLOWED"
    except Exception as e:
        result["error"] = str(e)
    return result


def check_llms_txt(url: str) -> dict:
    """Check for llms.txt presence."""
    parsed = urlparse(url)
    llms_url = f"{parsed.scheme}://{parsed.netloc}/llms.txt"
    result = {"found": False, "url": llms_url, "content": ""}
    try:
        r = requests.get(llms_url, timeout=8, headers={"User-Agent": "Mozilla/5.0"})
        if r.status_code == 200 and len(r.text) > 10:
            result["found"] = True
            result["content"] = r.text[:1000]
    except Exception:
        pass
    return result


# =============================================================================
# GEMINI ANALYSIS
# =============================================================================
def analyse_with_gemini(client, page_data: dict, dimension: str) -> dict:
    """Use Gemini to analyse a specific dimension for a page."""
    config = DIMENSION_CONFIG[dimension]
    html_snippet = page_data.get("html", "")[:6000]
    text_snippet = page_data.get("text", "")[:3000]

    prompts = {
        "ARIA Implementation": f"""Analyse the ARIA implementation on this webpage for AI visibility.

HTML (first 6000 chars): {html_snippet}

Evaluate:
1. Are ARIA roles, labels and landmarks used correctly (role="main", role="navigation", aria-label, etc.)?
2. Are form inputs properly labelled?
3. Are dynamic content regions marked with aria-live?
4. Do aria attributes help AI systems understand page structure?
5. Any missing or incorrect ARIA usage?

Return JSON only:
{{
  "score": <1-10>,
  "summary": "<2 sentence summary>",
  "findings": ["<finding 1>", "<finding 2>", ...],
  "issues": [
    {{"severity": "critical|warning|info", "issue": "<issue>", "recommendation": "<fix>"}}
  ],
  "positive": ["<what's done well>"]
}}""",

        "Structured Data / Schema": f"""Analyse the structured data / schema markup on this webpage.

HTML (first 6000 chars): {html_snippet}

Evaluate:
1. What schema.org types are present (JSON-LD, microdata, RDFa)?
2. Are they complete and correctly implemented?
3. What additional schema types would be valuable (Product, BreadcrumbList, FAQPage, Organization, etc.)?
4. Are required fields populated?
5. Rate the overall structured data quality for AI visibility.

Return JSON only:
{{
  "score": <1-10>,
  "summary": "<2 sentence summary>",
  "schemas_found": ["<schema type>", ...],
  "schemas_missing": ["<recommended schema>", ...],
  "findings": ["<finding>", ...],
  "issues": [
    {{"severity": "critical|warning|info", "issue": "<issue>", "recommendation": "<fix>"}}
  ],
  "positive": ["<what's done well>"]
}}""",

        "Heading Structure": f"""Analyse the heading structure (H1-H6) on this webpage.

HTML: {html_snippet}
Text: {text_snippet}

Evaluate:
1. Is there exactly one H1?
2. Is the heading hierarchy logical (no skipped levels)?
3. Are headings descriptive and keyword-rich?
4. Do headings help AI systems understand content structure?
5. Is there an appropriate number of headings for the content length?

Return JSON only:
{{
  "score": <1-10>,
  "summary": "<2 sentence summary>",
  "h1_count": <number>,
  "heading_hierarchy": "<description>",
  "findings": ["<finding>", ...],
  "issues": [
    {{"severity": "critical|warning|info", "issue": "<issue>", "recommendation": "<fix>"}}
  ],
  "positive": ["<what's done well>"]
}}""",

        "Meta & SEO Signals": f"""Analyse the meta tags and SEO signals on this webpage.

HTML: {html_snippet}

Evaluate:
1. Title tag: present, length (50-60 chars ideal), keyword relevance
2. Meta description: present, length (150-160 chars ideal), compelling
3. Canonical tag: present and correct
4. Open Graph tags: og:title, og:description, og:image
5. Twitter card tags
6. hreflang if international
7. Robots meta tag
8. Any duplicate or conflicting signals

Return JSON only:
{{
  "score": <1-10>,
  "summary": "<2 sentence summary>",
  "title": "<page title or null>",
  "title_length": <number>,
  "meta_description_present": true/false,
  "canonical_present": true/false,
  "og_tags_present": true/false,
  "findings": ["<finding>", ...],
  "issues": [
    {{"severity": "critical|warning|info", "issue": "<issue>", "recommendation": "<fix>"}}
  ],
  "positive": ["<what's done well>"]
}}""",

        "Link Quality": f"""Analyse the link quality on this webpage for AI visibility.

HTML: {html_snippet}
Text: {text_snippet}

Evaluate:
1. Internal link quality: descriptive anchor text vs generic ("click here", "read more")
2. External links: are they to authoritative sources?
3. Nofollow usage: appropriate or excessive?
4. Links with no anchor text (image links etc.)
5. Broken link indicators
6. Links that aid AI content graph understanding

Return JSON only:
{{
  "score": <1-10>,
  "summary": "<2 sentence summary>",
  "findings": ["<finding>", ...],
  "issues": [
    {{"severity": "critical|warning|info", "issue": "<issue>", "recommendation": "<fix>"}}
  ],
  "positive": ["<what's done well>"]
}}""",

        "Image Alt Text": f"""Analyse image alt text quality on this webpage.

HTML: {html_snippet}

Evaluate:
1. What percentage of images have alt text?
2. Quality of alt text: descriptive, keyword-relevant, not stuffed
3. Decorative images using empty alt=""
4. Complex images (charts, infographics) with adequate descriptions
5. AI systems' ability to understand page imagery

Return JSON only:
{{
  "score": <1-10>,
  "summary": "<2 sentence summary>",
  "images_found": <number>,
  "images_with_alt": <number>,
  "findings": ["<finding>", ...],
  "issues": [
    {{"severity": "critical|warning|info", "issue": "<issue>", "recommendation": "<fix>"}}
  ],
  "positive": ["<what's done well>"]
}}""",

        "Crawlability & Bot Access": f"""Analyse the crawlability and bot access for this webpage.

HTML length: {page_data.get('html_raw_length', 0)} chars
Text extracted: {page_data.get('text_length', 0)} chars
Text-to-HTML ratio: {page_data.get('text_to_html_ratio', 0)}
Status code: {page_data.get('status_code')}
Load time: {page_data.get('load_time')} seconds
Is HTTPS: {page_data.get('is_https')}
Redirect chain: {page_data.get('redirect_chain', [])}
Robots.txt data: {json.dumps(page_data.get('robots_data', {}))}

HTML snippet: {html_snippet[:2000]}

Evaluate:
1. Is the page server-side rendered (SSR) or client-side rendered (CSR)? A low text-to-HTML ratio (<0.05) suggests JS dependency.
2. Are there bot blocking mechanisms (Cloudflare, CAPTCHA, etc.)?
3. Is the page accessible over HTTPS?
4. Are there excessive redirects?
5. Is load time acceptable for crawlers?
6. Are AI bots blocked in robots.txt?

Return JSON only:
{{
  "score": <1-10>,
  "summary": "<2 sentence summary>",
  "rendering_type": "SSR|CSR|Mixed",
  "js_dependent": true/false,
  "findings": ["<finding>", ...],
  "issues": [
    {{"severity": "critical|warning|info", "issue": "<issue>", "recommendation": "<fix>"}}
  ],
  "positive": ["<what's done well>"]
}}""",

        "LLM Content Signals": f"""Analyse the content quality and LLM/AI visibility signals on this webpage.

Text content: {text_snippet}
HTML snippet: {html_snippet[:2000]}

Evaluate:
1. E-E-A-T signals: expertise, experience, authoritativeness, trustworthiness
2. Content depth and factual density: is it substantive or thin?
3. Clarity and structure: would an LLM extract clear facts from this?
4. Trust signals: author names, dates, citations, credentials, awards
5. Brand/entity clarity: is the brand identity clearly expressed?
6. Content freshness signals
7. FAQ-style content that LLMs can extract for featured snippets/answers

Return JSON only:
{{
  "score": <1-10>,
  "summary": "<2 sentence summary>",
  "eeat_signals": ["<signal>", ...],
  "findings": ["<finding>", ...],
  "issues": [
    {{"severity": "critical|warning|info", "issue": "<issue>", "recommendation": "<fix>"}}
  ],
  "positive": ["<what's done well>"]
}}""",

        "AI Search Health": f"""Analyse the AI search health signals for this webpage.

HTML: {html_snippet}
llms.txt found: {page_data.get('llms_txt', {}).get('found', False)}
llms.txt content: {page_data.get('llms_txt', {}).get('content', 'N/A')}
Robots.txt AI bots: {json.dumps(page_data.get('robots_data', {}).get('ai_bots', {}))}

Evaluate:
1. Is llms.txt present and well-structured?
2. Are AI crawlers (GPTBot, ChatGPT-User, OAI-SearchBot, ClaudeBot etc.) allowed in robots.txt?
3. Are there any AI-specific directives or meta tags?
4. Is the site structure conducive to AI search indexing?
5. Any proactive AI search optimisation signals?

Return JSON only:
{{
  "score": <1-10>,
  "summary": "<2 sentence summary>",
  "llms_txt_present": true/false,
  "ai_bots_blocked": ["<bot>", ...],
  "findings": ["<finding>", ...],
  "issues": [
    {{"severity": "critical|warning|info", "issue": "<issue>", "recommendation": "<fix>"}}
  ],
  "positive": ["<what's done well>"]
}}""",

        "Duplicate Content & Tags": f"""Analyse duplicate content and tag issues on this webpage.

HTML: {html_snippet}

Evaluate:
1. Is the title tag unique and not a CMS default?
2. Is there a canonical tag preventing duplicate content issues?
3. Are there signs of boilerplate/templated content that reduces page value?
4. Text-to-HTML ratio (low ratio can indicate content-thin pages)
5. Duplicate meta descriptions risk

Return JSON only:
{{
  "score": <1-10>,
  "summary": "<2 sentence summary>",
  "canonical_present": true/false,
  "findings": ["<finding>", ...],
  "issues": [
    {{"severity": "critical|warning|info", "issue": "<issue>", "recommendation": "<fix>"}}
  ],
  "positive": ["<what's done well>"]
}}"""
    }

    prompt = prompts.get(dimension, f"Analyse the {dimension} dimension. Return JSON with score (1-10), summary, findings, issues, positive fields.")

    try:
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=prompt
        )
        text = response.text.strip()
        # Strip markdown fences if present
        if text.startswith("```"):
            text = re.sub(r"^```(?:json)?\n?", "", text)
            text = re.sub(r"\n?```$", "", text)
        return json.loads(text)
    except json.JSONDecodeError:
        # Try to extract JSON from response
        match = re.search(r'\{.*\}', text, re.DOTALL)
        if match:
            try:
                return json.loads(match.group(0))
            except Exception:
                pass
        return {
            "score": 5,
            "summary": "Analysis completed.",
            "findings": [],
            "issues": [],
            "positive": []
        }
    except Exception as e:
        return {
            "score": 0,
            "summary": f"Analysis error: {str(e)}",
            "findings": [],
            "issues": [],
            "positive": []
        }


def generate_executive_summary(client, url: str, scores: dict, all_results: dict) -> str:
    """Generate an executive summary using Gemini."""
    overall = weighted_overall_score(scores)
    score_overview = "\n".join([f"- {dim}: {score}/10" for dim, score in scores.items()])

    # Gather top issues
    all_issues = []
    for dim, result in all_results.items():
        for issue in result.get("issues", []):
            if issue.get("severity") == "critical":
                all_issues.append(f"{dim}: {issue.get('issue', '')}")

    prompt = f"""You are an expert AI visibility consultant at Summit, a performance marketing agency.

Write a professional executive summary for an AI visibility audit of {url}.

Overall weighted score: {overall}/10
Dimension scores:
{score_overview}

Critical issues found:
{chr(10).join(all_issues[:8])}

Write 3-4 concise, professional paragraphs:
1. Overall assessment and score context
2. Key strengths
3. Priority areas for improvement
4. Business impact / why this matters for AI search visibility

Tone: Professional, direct, consultancy-grade. UK English. No bullet points in this summary."""

    try:
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=prompt
        )
        return response.text.strip()
    except Exception as e:
        return f"Executive summary generation failed: {str(e)}"


# =============================================================================
# PLOTLY CHARTS
# =============================================================================
def create_radar_chart(scores: dict) -> go.Figure:
    dims = list(scores.keys())
    vals = [scores[d] for d in dims]
    vals_closed = vals + [vals[0]]
    dims_closed = dims + [dims[0]]

    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=vals_closed,
        theta=dims_closed,
        fill="toself",
        fillcolor="rgba(194,82,0,0.15)",
        line=dict(color=SUMMIT_ORANGE, width=2.5),
        marker=dict(size=7, color=SUMMIT_ORANGE),
        name="Score"
    ))
    fig.update_layout(
        polar=dict(
            bgcolor="white",
            radialaxis=dict(visible=True, range=[0, 10], tickfont=dict(size=9), gridcolor="#E0E0E0"),
            angularaxis=dict(tickfont=dict(size=10, color="#333"))
        ),
        showlegend=False,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=60, r=60, t=40, b=40),
        height=420
    )
    return fig


def create_gauge_chart(overall_score: float) -> go.Figure:
    color = score_hex(overall_score)
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=overall_score,
        number={"suffix": "/10", "font": {"size": 44, "color": color, "family": "Inter"}},
        gauge={
            "axis": {"range": [0, 10], "tickwidth": 1, "tickcolor": "#666", "tickfont": {"size": 11}},
            "bar": {"color": color, "thickness": 0.3},
            "bgcolor": "white",
            "borderwidth": 0,
            "steps": [
                {"range": [0, 3.3], "color": "#FDECEA"},
                {"range": [3.3, 6.6], "color": "#FFF8E1"},
                {"range": [6.6, 10], "color": "#EBF7EF"}
            ],
            "threshold": {
                "line": {"color": color, "width": 4},
                "thickness": 0.75,
                "value": overall_score
            }
        }
    ))
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=30, b=20),
        height=250,
        font=dict(family="Inter")
    )
    return fig


def create_bar_chart(scores: dict) -> go.Figure:
    dims = list(scores.keys())
    vals = [scores[d] for d in dims]
    colors = [score_hex(v) for v in vals]
    bg_colors = [score_bg_hex(v) for v in vals]

    fig = go.Figure(go.Bar(
        x=vals,
        y=[d.replace(" & ", " &\n").replace(" / ", " /\n") for d in dims],
        orientation="h",
        marker_color=colors,
        text=[f"{v}/10" for v in vals],
        textposition="outside",
        textfont=dict(size=12, color="#333"),
        hovertemplate="%{y}: %{x}/10<extra></extra>"
    ))
    fig.update_layout(
        xaxis=dict(range=[0, 12], showgrid=True, gridcolor="#f0f0f0", title="Score (out of 10)"),
        yaxis=dict(autorange="reversed", tickfont=dict(size=11)),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10, r=80, t=20, b=30),
        height=420,
        font=dict(family="Inter"),
        showlegend=False
    )
    # Add reference lines
    for threshold, color, label in [(3.3, "#C0392B", "Critical"), (6.6, "#276749", "Good")]:
        fig.add_vline(x=threshold, line_dash="dash", line_color=color, opacity=0.5,
                      annotation_text=label, annotation_position="top",
                      annotation_font=dict(size=10, color=color))
    return fig


def create_weight_chart() -> go.Figure:
    """Show dimension weights as a donut chart."""
    dims = list(DIMENSION_CONFIG.keys())
    weights = [DIMENSION_CONFIG[d]["weight"] * 100 for d in dims]
    colors = [DIMENSION_CONFIG[d]["color"] for d in dims]

    fig = go.Figure(go.Pie(
        labels=dims,
        values=weights,
        hole=0.5,
        marker=dict(colors=colors),
        textinfo="label+percent",
        textfont=dict(size=10),
        hovertemplate="%{label}: %{value:.0f}% weight<extra></extra>"
    ))
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        showlegend=False,
        margin=dict(l=10, r=10, t=20, b=10),
        height=380,
        font=dict(family="Inter")
    )
    return fig


# =============================================================================
# DOCX EXPORT
# =============================================================================
def set_cell_bg(cell, hex_color: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def add_paragraph_border_bottom(para, color="C25200", size=8):
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def create_docx_report(url: str, scores: dict, all_results: dict, executive_summary: str,
                       robots_data: dict, llms_txt: dict, logo_bytes: bytes) -> bytes:
    """Generate a Summit-branded .docx report."""
    doc = Document()

    # Page setup — A4
    for section in doc.sections:
        section.page_width = Emu(11906 * 914)
        section.page_height = Emu(16838 * 914)
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Styles
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(10)

    # ==========  COVER / HEADER  ==========
    # Logo
    logo_stream = io.BytesIO(logo_bytes)
    logo_para = doc.add_paragraph()
    logo_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = logo_para.add_run()
    run.add_picture(logo_stream, width=Inches(1.8))

    # Title
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_paragraph_border_bottom(title_para, color="C25200", size=12)
    r = title_para.add_run("AI Visibility Audit Report")
    r.bold = True
    r.font.size = Pt(26)
    r.font.color.rgb = RGBColor(0x1A, 0x30, 0x55)
    r.font.name = "Arial"

    # URL and date
    meta_para = doc.add_paragraph()
    r1 = meta_para.add_run(f"{url}  |  ")
    r1.font.size = Pt(10)
    r1.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    r2 = meta_para.add_run(datetime.now().strftime("%B %Y"))
    r2.font.size = Pt(10)
    r2.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    doc.add_paragraph()

    # ==========  OVERALL SCORE  ==========
    overall = weighted_overall_score(scores)
    score_section = doc.add_paragraph()
    add_paragraph_border_bottom(score_section, color="C25200", size=8)
    r = score_section.add_run("Overall AI Visibility Score")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0x1A, 0x30, 0x55)
    r.font.name = "Arial"

    # Score table
    score_tbl = doc.add_table(rows=1, cols=3)
    score_tbl.style = "Table Grid"
    score_tbl.autofit = False
    col_widths = [2400, 2400, 5000]
    for i, cell in enumerate(score_tbl.rows[0].cells):
        cell.width = Emu(col_widths[i] * 914)

    cells = score_tbl.rows[0].cells

    # Cell 1: overall score
    set_cell_bg(cells[0], "1A3055")
    p = cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"{overall}")
    r.bold = True
    r.font.size = Pt(48)
    r.font.color.rgb = RGBColor(255, 255, 255)
    p2 = cells[0].add_paragraph("/10")
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.runs[0].font.color.rgb = RGBColor(255, 200, 160)
    p2.runs[0].font.size = Pt(16)

    # Cell 2: rating label
    color_class = score_color_class(overall)
    bg_map = {"red": "FDECEA", "amber": "FFF8E1", "green": "EBF7EF"}
    fg_map = {"red": (0xC0, 0x39, 0x2B), "amber": (0x7B, 0x38, 0x00), "green": (0x27, 0x67, 0x49)}
    label_map = {"red": "Needs Attention", "amber": "Developing", "green": "Good"}
    set_cell_bg(cells[1], bg_map[color_class])
    p = cells[1].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(label_map[color_class])
    r.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(*fg_map[color_class])

    # Cell 3: score breakdown
    set_cell_bg(cells[2], "F5F5F5")
    p = cells[2].paragraphs[0]
    r = p.add_run("Weighted Score Breakdown")
    r.bold = True
    r.font.size = Pt(10)
    for dim, score in sorted(scores.items(), key=lambda x: x[1]):
        p2 = cells[2].add_paragraph()
        r_dim = p2.add_run(f"{dim}: ")
        r_dim.font.size = Pt(9)
        r_score = p2.add_run(f"{score}/10")
        r_score.bold = True
        r_score.font.size = Pt(9)
        fg = fg_map[score_color_class(score)]
        r_score.font.color.rgb = RGBColor(*fg)

    doc.add_paragraph()

    # ==========  EXECUTIVE SUMMARY  ==========
    es_heading = doc.add_paragraph()
    add_paragraph_border_bottom(es_heading, color="C25200", size=8)
    r = es_heading.add_run("Executive Summary")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0x1A, 0x30, 0x55)

    for para_text in executive_summary.split("\n\n"):
        if para_text.strip():
            p = doc.add_paragraph(para_text.strip())
            p.style.font.size = Pt(10)
    doc.add_paragraph()

    # ==========  DIMENSION SCORES TABLE  ==========
    dim_heading = doc.add_paragraph()
    add_paragraph_border_bottom(dim_heading, color="C25200", size=8)
    r = dim_heading.add_run("Dimension Scores")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0x1A, 0x30, 0x55)

    dim_tbl = doc.add_table(rows=1, cols=4)
    dim_tbl.style = "Table Grid"
    dim_tbl.autofit = False

    headers = ["Dimension", "Weight", "Score", "Rating"]
    header_widths = [4000, 1200, 1200, 3400]
    header_row = dim_tbl.rows[0]
    for i, (cell, hdr) in enumerate(zip(header_row.cells, headers)):
        set_cell_bg(cell, "1A3055")
        cell.width = Emu(header_widths[i] * 914)
        p = cell.paragraphs[0]
        r = p.add_run(hdr)
        r.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(9)

    for dim, score in scores.items():
        row = dim_tbl.add_row()
        cc = score_color_class(score)
        bg = bg_map[cc]
        fg = fg_map[cc]
        weight_pct = int(DIMENSION_CONFIG.get(dim, {}).get("weight", 0.05) * 100)
        label_map2 = {"red": "❌ Critical", "amber": "⚠ Needs Work", "green": "✓ Good"}

        row.cells[0].width = Emu(4000 * 914)
        row.cells[1].width = Emu(1200 * 914)
        row.cells[2].width = Emu(1200 * 914)
        row.cells[3].width = Emu(3400 * 914)

        row.cells[0].paragraphs[0].add_run(dim).font.size = Pt(9)
        row.cells[1].paragraphs[0].add_run(f"{weight_pct}%").font.size = Pt(9)

        set_cell_bg(row.cells[2], bg)
        score_run = row.cells[2].paragraphs[0].add_run(f"{score}/10")
        score_run.bold = True
        score_run.font.size = Pt(10)
        score_run.font.color.rgb = RGBColor(*fg)

        set_cell_bg(row.cells[3], bg)
        label_run = row.cells[3].paragraphs[0].add_run(label_map2[cc])
        label_run.font.size = Pt(9)
        label_run.font.color.rgb = RGBColor(*fg)

    doc.add_paragraph()

    # ==========  DETAILED FINDINGS PER DIMENSION  ==========
    findings_heading = doc.add_paragraph()
    add_paragraph_border_bottom(findings_heading, color="C25200", size=8)
    r = findings_heading.add_run("Detailed Findings")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0x1A, 0x30, 0x55)

    for dim, result in all_results.items():
        score = scores.get(dim, 0)
        cc = score_color_class(score)
        fg = fg_map[cc]
        bg = bg_map[cc]

        # Dimension heading
        dim_para = doc.add_paragraph()
        r_icon = dim_para.add_run(f"{DIMENSION_CONFIG.get(dim, {}).get('icon', '•')} {dim}  ")
        r_icon.bold = True
        r_icon.font.size = Pt(12)
        r_icon.font.color.rgb = RGBColor(0x1A, 0x30, 0x55)
        r_score = dim_para.add_run(f"[{score}/10]")
        r_score.bold = True
        r_score.font.color.rgb = RGBColor(*fg)
        r_score.font.size = Pt(12)

        # Summary
        if result.get("summary"):
            p = doc.add_paragraph(result["summary"])
            p.runs[0].font.size = Pt(10)
            p.runs[0].italic = True

        # Issues table
        issues = result.get("issues", [])
        if issues:
            sev_heading = doc.add_paragraph()
            r = sev_heading.add_run("Issues & Recommendations")
            r.bold = True
            r.font.size = Pt(10)
            r.font.color.rgb = RGBColor(0x1A, 0x30, 0x55)

            iss_tbl = doc.add_table(rows=1, cols=3)
            iss_tbl.style = "Table Grid"
            iss_tbl.autofit = False
            i_headers = ["Severity", "Issue", "Recommendation"]
            i_widths = [1200, 3500, 5100]
            for i, (cell, hdr) in enumerate(zip(iss_tbl.rows[0].cells, i_headers)):
                set_cell_bg(cell, "1A3055")
                cell.width = Emu(i_widths[i] * 914)
                r = cell.paragraphs[0].add_run(hdr)
                r.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
                r.font.size = Pt(9)

            sev_colors = {"critical": ("FDECEA", (192, 57, 43)), "warning": ("FFF8E1", (123, 56, 0)), "info": ("E3F2FD", (21, 101, 192))}
            for issue in issues:
                sev = issue.get("severity", "info")
                bg_i, fg_i = sev_colors.get(sev, ("F5F5F5", (0, 0, 0)))
                row = iss_tbl.add_row()
                row.cells[0].width = Emu(1200 * 914)
                row.cells[1].width = Emu(3500 * 914)
                row.cells[2].width = Emu(5100 * 914)

                set_cell_bg(row.cells[0], bg_i)
                sev_run = row.cells[0].paragraphs[0].add_run(sev.upper())
                sev_run.bold = True
                sev_run.font.size = Pt(8)
                sev_run.font.color.rgb = RGBColor(*fg_i)

                row.cells[1].paragraphs[0].add_run(issue.get("issue", "")).font.size = Pt(9)
                row.cells[2].paragraphs[0].add_run(issue.get("recommendation", "")).font.size = Pt(9)

        # Positive findings
        positives = result.get("positive", [])
        if positives:
            pos_para = doc.add_paragraph()
            r = pos_para.add_run("✓ Strengths:  ")
            r.bold = True
            r.font.color.rgb = RGBColor(0x27, 0x67, 0x49)
            r.font.size = Pt(9)
            r2 = pos_para.add_run("  |  ".join(positives[:3]))
            r2.font.size = Pt(9)
            r2.font.color.rgb = RGBColor(0x27, 0x67, 0x49)

        doc.add_paragraph()

    # ==========  PRIORITY RECOMMENDATIONS  ==========
    rec_heading = doc.add_paragraph()
    add_paragraph_border_bottom(rec_heading, color="C25200", size=8)
    r = rec_heading.add_run("Priority Recommendations")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0x1A, 0x30, 0x55)

    # Collect all critical/warning issues sorted by dimension weight
    all_recs = []
    for dim, result in all_results.items():
        weight = DIMENSION_CONFIG.get(dim, {}).get("weight", 0.05)
        for issue in result.get("issues", []):
            sev = issue.get("severity", "info")
            priority = 1 if sev == "critical" else (2 if sev == "warning" else 3)
            all_recs.append({
                "priority": priority,
                "weight": weight,
                "dimension": dim,
                "issue": issue.get("issue", ""),
                "recommendation": issue.get("recommendation", ""),
                "severity": sev
            })
    all_recs.sort(key=lambda x: (x["priority"], -x["weight"]))

    rec_tbl = doc.add_table(rows=1, cols=4)
    rec_tbl.style = "Table Grid"
    rec_tbl.autofit = False
    r_headers = ["Priority", "Dimension", "Issue", "Recommendation"]
    r_widths = [700, 2000, 3300, 3800]
    for i, (cell, hdr) in enumerate(zip(rec_tbl.rows[0].cells, r_headers)):
        set_cell_bg(cell, "1A3055")
        cell.width = Emu(r_widths[i] * 914)
        r = cell.paragraphs[0].add_run(hdr)
        r.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(9)

    sev_colors = {"critical": ("FDECEA", (192, 57, 43)), "warning": ("FFF8E1", (123, 56, 0)), "info": ("E3F2FD", (21, 101, 192))}
    for rank, rec in enumerate(all_recs[:15], 1):
        row = rec_tbl.add_row()
        sev = rec["severity"]
        bg_r, fg_r = sev_colors.get(sev, ("F5F5F5", (0, 0, 0)))
        row.cells[0].width = Emu(r_widths[0] * 914)
        row.cells[1].width = Emu(r_widths[1] * 914)
        row.cells[2].width = Emu(r_widths[2] * 914)
        row.cells[3].width = Emu(r_widths[3] * 914)

        set_cell_bg(row.cells[0], bg_r)
        p_run = row.cells[0].paragraphs[0].add_run(f"P{rank}")
        p_run.bold = True
        p_run.font.size = Pt(10)
        p_run.font.color.rgb = RGBColor(*fg_r)

        row.cells[1].paragraphs[0].add_run(rec["dimension"]).font.size = Pt(9)
        row.cells[2].paragraphs[0].add_run(rec["issue"]).font.size = Pt(9)
        row.cells[3].paragraphs[0].add_run(rec["recommendation"]).font.size = Pt(9)

    doc.add_paragraph()

    # ==========  FOOTER  ==========
    footer_para = doc.add_paragraph()
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_paragraph_border_bottom(footer_para, color="C25200", size=4)
    r1 = footer_para.add_run("Prepared by Summit  |  AI Visibility Practice  |  summit.co.uk")
    r1.font.size = Pt(9)
    r1.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    r1.italic = True

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# =============================================================================
# PDF EXPORT (dashboard summary)
# =============================================================================
def create_pdf_report(url: str, scores: dict, executive_summary: str, logo_bytes: bytes) -> bytes:
    """Create a PDF summary of the dashboard."""
    from fpdf import FPDF

    class SummitPDF(FPDF):
        def header(self):
            # Logo
            logo_path = "/tmp/summit_logo_pdf.png"
            with open(logo_path, "wb") as f:
                f.write(logo_bytes)
            self.image(logo_path, 12, 8, 40)
            self.set_font("Helvetica", "B", 16)
            self.set_text_color(26, 48, 85)
            self.set_xy(55, 10)
            self.cell(0, 8, "AI Visibility Audit Report", ln=True)
            self.set_font("Helvetica", "", 9)
            self.set_text_color(102, 102, 102)
            self.set_x(55)
            self.cell(0, 6, f"{url}  |  {datetime.now().strftime('%B %Y')}", ln=True)
            self.ln(4)
            # Orange line
            self.set_draw_color(194, 82, 0)
            self.set_line_width(1.2)
            self.line(12, 28, 198, 28)
            self.ln(6)

        def footer(self):
            self.set_y(-12)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(150, 150, 150)
            self.cell(0, 6, f"Summit AI Visibility Audit  |  Confidential  |  Page {self.page_no()}", align="C")

    pdf = SummitPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    overall = weighted_overall_score(scores)

    # Overall score hero block
    cc = score_color_class(overall)
    bg_map = {"red": (253, 236, 234), "amber": (255, 248, 225), "green": (235, 247, 239)}
    fg_map_rgb = {"red": (192, 57, 43), "amber": (123, 56, 0), "green": (39, 103, 73)}
    label_map = {"red": "Needs Attention", "amber": "Developing", "green": "Good"}

    pdf.set_fill_color(26, 48, 85)
    pdf.rect(12, pdf.get_y(), 90, 40, "F")
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 36)
    pdf.set_xy(12, pdf.get_y() + 4)
    pdf.cell(90, 20, f"{overall}/10", align="C", ln=False)
    pdf.set_xy(12, pdf.get_y() + 20)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(255, 200, 160)
    pdf.cell(90, 10, "Overall AI Visibility Score", align="C", ln=False)

    # Rating badge
    pdf.set_fill_color(*bg_map[cc])
    pdf.rect(106, pdf.get_y() - 30, 90, 40, "F")
    pdf.set_text_color(*fg_map_rgb[cc])
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_xy(106, pdf.get_y() - 24)
    pdf.cell(90, 14, label_map[cc], align="C", ln=False)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_xy(106, pdf.get_y() + 10)
    pdf.cell(90, 8, f"Weighted across {len(scores)} dimensions", align="C")
    pdf.ln(18)

    # Dimension scores table
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(26, 48, 85)
    pdf.cell(0, 7, "Dimension Scores", ln=True)
    pdf.set_draw_color(194, 82, 0)
    pdf.set_line_width(0.8)
    pdf.line(12, pdf.get_y(), 198, pdf.get_y())
    pdf.ln(3)

    # Table header
    pdf.set_fill_color(26, 48, 85)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(90, 7, "Dimension", fill=True, border=0)
    pdf.cell(20, 7, "Weight", fill=True, border=0, align="C")
    pdf.cell(20, 7, "Score", fill=True, border=0, align="C")
    pdf.cell(58, 7, "Rating", fill=True, border=0)
    pdf.ln()

    sev_label = {"red": "Needs Attention", "amber": "Needs Work", "green": "Good"}
    for i, (dim, score) in enumerate(sorted(scores.items(), key=lambda x: x[1])):
        cc_d = score_color_class(score)
        bg = bg_map[cc_d]
        fg = fg_map_rgb[cc_d]
        w_pct = int(DIMENSION_CONFIG.get(dim, {}).get("weight", 0.05) * 100)

        fill_bg = bg if i % 2 == 0 else (252, 252, 252)
        pdf.set_fill_color(*fill_bg)
        pdf.set_text_color(30, 30, 30)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.cell(90, 6.5, dim, fill=True, border=0)
        pdf.cell(20, 6.5, f"{w_pct}%", fill=True, border=0, align="C")

        pdf.set_fill_color(*bg)
        pdf.set_text_color(*fg)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(20, 6.5, f"{score}/10", fill=True, border=0, align="C")

        pdf.set_fill_color(*bg)
        pdf.set_text_color(*fg)
        pdf.set_font("Helvetica", "", 8.5)
        pdf.cell(58, 6.5, sev_label[cc_d], fill=True, border=0)
        pdf.ln()

    pdf.ln(6)

    # Executive Summary
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(26, 48, 85)
    pdf.cell(0, 7, "Executive Summary", ln=True)
    pdf.set_draw_color(194, 82, 0)
    pdf.line(12, pdf.get_y(), 198, pdf.get_y())
    pdf.ln(3)

    pdf.set_font("Helvetica", "", 9.5)
    pdf.set_text_color(40, 40, 40)
    for para in executive_summary.split("\n\n"):
        if para.strip():
            pdf.multi_cell(0, 5.5, para.strip())
            pdf.ln(2)

    buf = io.BytesIO()
    pdf_bytes = pdf.output()
    if isinstance(pdf_bytes, str):
        return pdf_bytes.encode("latin-1")
    return bytes(pdf_bytes)


# =============================================================================
# MAIN STREAMLIT APP
# =============================================================================
def main():
    inject_css()

    # Logo in sidebar
    logo_img = create_summit_logo()
    logo_b64 = logo_to_base64(logo_img)
    logo_bytes = logo_to_bytes(logo_img)

    with st.sidebar:
        st.markdown(
            f'<img src="data:image/png;base64,{logo_b64}" width="200" style="margin-bottom:24px">',
            unsafe_allow_html=True
        )
        st.markdown("### AI Visibility Audit")
        st.markdown("---")
        st.markdown("**Configuration**")

        api_key = st.text_input("Gemini API Key", type="password", help="Enter your Google Gemini API key")

        url_input = st.text_input("URL to Audit", placeholder="https://example.com", help="Enter the full URL of the page to audit")

        st.markdown("---")
        st.markdown("**Scoring Weights**")
        st.caption("Dimensions are pre-weighted by AI impact. Highest weights:")
        for dim, config in list(DIMENSION_CONFIG.items())[:4]:
            st.caption(f"{config['icon']} {dim}: {int(config['weight']*100)}%")

        st.markdown("---")
        run_audit = st.button("🚀 Run Audit", use_container_width=True)

        st.markdown("---")
        st.markdown("**About**")
        st.caption("Summit AI Visibility Audit Tool v2. Powered by Google Gemini 2.0 Flash.")
        st.caption("© Summit Performance Marketing")

    # Main content
    col_header_l, col_header_r = st.columns([3, 1])
    with col_header_l:
        st.markdown(f"""
        <div style="display:flex;align-items:center;gap:16px;margin-bottom:8px">
            <img src="data:image/png;base64,{logo_b64}" width="140">
            <div>
                <h1 style="margin:0;color:#1A3055;font-size:1.8rem">AI Visibility Audit Dashboard</h1>
                <p style="margin:0;color:#666;font-size:0.9rem">Technical audit for AI crawler access & content extraction</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col_header_r:
        st.markdown(f"<p style='text-align:right;color:#666;margin-top:20px;font-size:0.85rem'>Last run: {datetime.now().strftime('%d %b %Y')}</p>", unsafe_allow_html=True)

    st.markdown("---")

    if not run_audit:
        # Landing state
        st.markdown("""
        <div style="background:white;border-radius:12px;padding:40px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,0.08)">
            <h2 style="color:#1A3055">Enter a URL to begin your audit</h2>
            <p style="color:#666;max-width:500px;margin:16px auto">This tool analyses your website across 10 AI visibility dimensions,
            weighted by impact. Get a Semrush-style dashboard with exportable reports.</p>
        </div>
        """, unsafe_allow_html=True)

        # Show dimension overview
        st.markdown("### Audit Dimensions")
        cols = st.columns(2)
        for i, (dim, config) in enumerate(DIMENSION_CONFIG.items()):
            with cols[i % 2]:
                st.markdown(f"""
                <div style="background:white;border-radius:8px;padding:14px 18px;margin-bottom:10px;
                box-shadow:0 1px 4px rgba(0,0,0,0.06);border-left:4px solid {config['color']}">
                    <div style="display:flex;justify-content:space-between;align-items:center">
                        <span style="font-weight:600;font-size:0.92rem;color:#1A3055">{config['icon']} {dim}</span>
                        <span style="background:#F0F4FF;color:#1A3055;padding:2px 8px;border-radius:10px;font-size:0.78rem;font-weight:600">{int(config['weight']*100)}%</span>
                    </div>
                    <p style="margin:4px 0 0;color:#666;font-size:0.82rem">{config['description']}</p>
                </div>
                """, unsafe_allow_html=True)
        return

    # Validation
    if not api_key:
        st.error("Please enter your Gemini API key in the sidebar.")
        return
    if not url_input:
        st.error("Please enter a URL to audit.")
        return

    url = url_input.strip()
    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    # ==========  RUN AUDIT  ==========
    st.markdown(f"### Auditing: `{url}`")
    progress_bar = st.progress(0)
    status_text = st.empty()

    # Initialise Gemini
    try:
        client = genai.Client(api_key=api_key)
    except Exception as e:
        st.error(f"Gemini API error: {e}")
        return

    # Step 1: Fetch page
    status_text.text("📡 Fetching page...")
    progress_bar.progress(5)
    page_data = fetch_page(url)

    if page_data.get("error"):
        st.error(f"Failed to fetch page: {page_data['error']}")
        return

    # Step 2: Robots & llms.txt
    status_text.text("🤖 Checking robots.txt and llms.txt...")
    progress_bar.progress(10)
    robots_data = check_robots_txt(url)
    llms_txt = check_llms_txt(url)
    page_data["robots_data"] = robots_data
    page_data["llms_txt"] = llms_txt

    # Step 3: Analyse each dimension
    scores = {}
    all_results = {}
    dimensions = list(DIMENSION_CONFIG.keys())
    n_dims = len(dimensions)

    for i, dim in enumerate(dimensions):
        status_text.text(f"🔍 Analysing: {dim}...")
        progress_bar.progress(15 + int(70 * (i / n_dims)))
        result = analyse_with_gemini(client, page_data, dim)
        scores[dim] = result.get("score", 0)
        all_results[dim] = result
        time.sleep(0.3)  # Rate limit buffer

    # Step 4: Executive summary
    status_text.text("✍️ Generating executive summary...")
    progress_bar.progress(88)
    exec_summary = generate_executive_summary(client, url, scores, all_results)

    # Step 5: Done
    progress_bar.progress(100)
    status_text.text("✅ Audit complete!")
    time.sleep(0.5)
    status_text.empty()
    progress_bar.empty()

    overall = weighted_overall_score(scores)
    cc = score_color_class(overall)
    bg_map = {"red": SUMMIT_BG_RED, "amber": SUMMIT_BG_AMBER, "green": SUMMIT_BG_GREEN}
    fg_map = {"red": SUMMIT_RED_SCORE, "amber": SUMMIT_AMBER_SCORE, "green": SUMMIT_GREEN_SCORE}
    label_map = {"red": "Needs Attention", "amber": "Developing", "green": "Good"}

    # ==========  DASHBOARD  ==========

    # Hero row: overall score + gauge + stats
    col1, col2, col3, col4 = st.columns([1.5, 1.5, 1, 1])

    with col1:
        st.markdown(f"""
        <div class="score-hero">
            <p style="font-size:0.85rem;font-weight:600;text-transform:uppercase;letter-spacing:0.08em;opacity:0.7;margin-bottom:8px">Overall AI Visibility Score</p>
            <h1>{overall}</h1>
            <p style="font-size:1rem;opacity:0.7">/10 weighted score</p>
            <div style="margin-top:12px;background:rgba(255,255,255,0.15);border-radius:8px;padding:8px 16px;display:inline-block">
                <span style="font-weight:700;font-size:1.1rem">{label_map[cc]}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        fig_gauge = create_gauge_chart(overall)
        st.plotly_chart(fig_gauge, use_container_width=True, config={"displayModeBar": False})

    with col3:
        critical_count = sum(1 for r in all_results.values() for iss in r.get("issues", []) if iss.get("severity") == "critical")
        warning_count = sum(1 for r in all_results.values() for iss in r.get("issues", []) if iss.get("severity") == "warning")
        st.markdown(f"""
        <div class="score-card score-card-red" style="text-align:center">
            <div class="metric-big" style="color:{SUMMIT_RED_SCORE}">{critical_count}</div>
            <div class="metric-label">Critical Issues</div>
        </div>
        <div class="score-card score-card-amber" style="text-align:center">
            <div class="metric-big" style="color:{SUMMIT_AMBER_SCORE}">{warning_count}</div>
            <div class="metric-label">Warnings</div>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        green_count = sum(1 for s in scores.values() if s > 6.5)
        load_time = page_data.get("load_time", "—")
        st.markdown(f"""
        <div class="score-card score-card-green" style="text-align:center">
            <div class="metric-big" style="color:{SUMMIT_GREEN_SCORE}">{green_count}</div>
            <div class="metric-label">Dimensions Passing</div>
        </div>
        <div class="score-card" style="text-align:center">
            <div class="metric-big" style="color:#1A3055">{load_time}s</div>
            <div class="metric-label">Page Load Time</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    # Charts row
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Score Overview", "🕸️ Radar Chart", "⚖️ Weighting", "📋 Issues", "📝 Executive Summary"])

    with tab1:
        col_bar, col_cards = st.columns([2, 1])
        with col_bar:
            st.markdown("#### Score by Dimension")
            fig_bar = create_bar_chart(scores)
            st.plotly_chart(fig_bar, use_container_width=True, config={"displayModeBar": False})

        with col_cards:
            st.markdown("#### Score Breakdown")
            for dim, score in sorted(scores.items(), key=lambda x: x[1]):
                cc_d = score_color_class(score)
                color = fg_map[cc_d]
                icon = DIMENSION_CONFIG.get(dim, {}).get("icon", "•")
                badge_cls = f"badge-{cc_d}"
                st.markdown(f"""
                <div style="display:flex;justify-content:space-between;align-items:center;
                padding:8px 12px;background:white;border-radius:6px;margin-bottom:6px;
                box-shadow:0 1px 3px rgba(0,0,0,0.06)">
                    <span style="font-size:0.85rem;color:#333">{icon} {dim}</span>
                    <span class="{badge_cls}">{score}/10</span>
                </div>
                """, unsafe_allow_html=True)

    with tab2:
        st.markdown("#### Radar Chart — AI Visibility Profile")
        fig_radar = create_radar_chart(scores)
        st.plotly_chart(fig_radar, use_container_width=True, config={"displayModeBar": False})

    with tab3:
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            st.markdown("#### Dimension Weighting")
            st.caption("Weights reflect relative impact on AI crawler visibility and content extraction quality.")
            fig_weight = create_weight_chart()
            st.plotly_chart(fig_weight, use_container_width=True, config={"displayModeBar": False})
        with col_w2:
            st.markdown("#### Weight Table")
            weight_rows = ""
            for dim, config in DIMENSION_CONFIG.items():
                score = scores.get(dim, 0)
                cc_d = score_color_class(score)
                badge_cls = f"badge-{cc_d}"
                weight_rows += f"""
                <tr>
                    <td>{config['icon']} {dim}</td>
                    <td style="text-align:center">{int(config['weight']*100)}%</td>
                    <td style="text-align:center"><span class="{badge_cls}">{score}/10</span></td>
                    <td style="text-align:center;font-weight:600">{round(score * config['weight'], 2)}</td>
                </tr>"""
            st.markdown(f"""
            <table class="audit-table">
                <thead><tr><th>Dimension</th><th>Weight</th><th>Score</th><th>Weighted</th></tr></thead>
                <tbody>{weight_rows}</tbody>
            </table>
            """, unsafe_allow_html=True)

    with tab4:
        st.markdown("#### All Issues")
        sev_filter = st.selectbox("Filter by Severity", ["All", "Critical", "Warning", "Info"])

        sev_icon = {"critical": "🔴", "warning": "🟡", "info": "🔵"}
        sev_class = {"critical": "critical", "warning": "warning", "info": "info"}

        all_issues_flat = []
        for dim, result in all_results.items():
            for iss in result.get("issues", []):
                all_issues_flat.append({**iss, "dimension": dim})
        all_issues_flat.sort(key=lambda x: {"critical": 0, "warning": 1, "info": 2}.get(x.get("severity", "info"), 3))

        displayed = 0
        for iss in all_issues_flat:
            sev = iss.get("severity", "info")
            if sev_filter != "All" and sev.lower() != sev_filter.lower():
                continue
            displayed += 1
            icon = sev_icon.get(sev, "⚪")
            cls = sev_class.get(sev, "info")
            st.markdown(f"""
            <div class="issue-row {cls}">
                <div style="min-width:28px">{icon}</div>
                <div style="flex:1">
                    <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px">
                        <span style="font-weight:600;font-size:0.88rem">{iss.get('issue', '')}</span>
                        <span style="background:#f0f0f0;color:#555;padding:1px 8px;border-radius:10px;font-size:0.76rem">{iss.get('dimension','')}</span>
                    </div>
                    <p style="margin:0;color:#555;font-size:0.83rem">💡 {iss.get('recommendation', '')}</p>
                </div>
            </div>
            """, unsafe_allow_html=True)

        if displayed == 0:
            st.info("No issues found for the selected filter.")

    with tab5:
        st.markdown("#### Executive Summary")
        st.markdown(f"""
        <div style="background:white;border-radius:12px;padding:28px;box-shadow:0 2px 8px rgba(0,0,0,0.08)">
            {exec_summary.replace(chr(10), '<br>')}
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    # ==========  EXPORTS  ==========
    st.markdown("### 📥 Export Reports")
    col_exp1, col_exp2, col_exp3 = st.columns(3)

    with col_exp1:
        st.markdown("""
        <div style="background:white;border-radius:10px;padding:20px;box-shadow:0 2px 6px rgba(0,0,0,0.08);border-top:4px solid #1A3055">
            <h4 style="color:#1A3055;margin-top:0">📄 Word Document</h4>
            <p style="color:#666;font-size:0.85rem">Full audit report with Summit branding. Includes executive summary, dimension scores, detailed findings, and priority recommendations table.</p>
        </div>
        """, unsafe_allow_html=True)
        with st.spinner("Generating .docx..."):
            docx_bytes = create_docx_report(url, scores, all_results, exec_summary, robots_data, llms_txt, logo_bytes)
        domain = urlparse(url).netloc.replace("www.", "")
        filename_docx = f"summit_ai_audit_{domain}_{datetime.now().strftime('%Y%m')}.docx"
        st.download_button(
            label="⬇️ Download Word Report",
            data=docx_bytes,
            file_name=filename_docx,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    with col_exp2:
        st.markdown("""
        <div style="background:white;border-radius:10px;padding:20px;box-shadow:0 2px 6px rgba(0,0,0,0.08);border-top:4px solid #C25200">
            <h4 style="color:#C25200;margin-top:0">📋 PDF Summary</h4>
            <p style="color:#666;font-size:0.85rem">Dashboard summary as a PDF. Includes overall score, dimension breakdown table, and executive summary. Ideal for client presentations.</p>
        </div>
        """, unsafe_allow_html=True)
        with st.spinner("Generating PDF..."):
            pdf_bytes = create_pdf_report(url, scores, exec_summary, logo_bytes)
        filename_pdf = f"summit_ai_audit_{domain}_{datetime.now().strftime('%Y%m')}.pdf"
        st.download_button(
            label="⬇️ Download PDF Report",
            data=pdf_bytes,
            file_name=filename_pdf,
            mime="application/pdf",
            use_container_width=True
        )

    with col_exp3:
        st.markdown("""
        <div style="background:white;border-radius:10px;padding:20px;box-shadow:0 2px 6px rgba(0,0,0,0.08);border-top:4px solid #276749">
            <h4 style="color:#276749;margin-top:0">📊 JSON Data</h4>
            <p style="color:#666;font-size:0.85rem">Raw audit data in JSON format. Includes all scores, findings, and Gemini analysis output. Useful for integration with other tools.</p>
        </div>
        """, unsafe_allow_html=True)
        export_data = {
            "url": url,
            "audit_date": datetime.now().isoformat(),
            "overall_score": overall,
            "scores": scores,
            "results": {k: {kk: vv for kk, vv in v.items() if kk != "soup"} for k, v in all_results.items()},
            "executive_summary": exec_summary
        }
        json_str = json.dumps(export_data, indent=2, default=str)
        filename_json = f"summit_ai_audit_{domain}_{datetime.now().strftime('%Y%m')}.json"
        st.download_button(
            label="⬇️ Download JSON Data",
            data=json_str,
            file_name=filename_json,
            mime="application/json",
            use_container_width=True
        )

    # ==========  DETAILED ACCORDION  ==========
    st.markdown("---")
    st.markdown("### 🔎 Detailed Findings by Dimension")

    for dim, result in all_results.items():
        score = scores.get(dim, 0)
        cc_d = score_color_class(score)
        icon = DIMENSION_CONFIG.get(dim, {}).get("icon", "•")
        with st.expander(f"{icon} {dim}  —  Score: {score}/10", expanded=(score <= 3)):
            col_a, col_b = st.columns([1, 2])
            with col_a:
                st.markdown(f"""
                <div style="text-align:center;padding:20px;background:{bg_map[cc_d]};border-radius:10px">
                    <div style="font-size:3rem;font-weight:800;color:{fg_map[cc_d]}">{score}</div>
                    <div style="color:{fg_map[cc_d]};font-size:0.8rem;font-weight:600;text-transform:uppercase">/10 — {label_map[cc_d]}</div>
                </div>
                """, unsafe_allow_html=True)
                weight = int(DIMENSION_CONFIG.get(dim, {}).get("weight", 0.05) * 100)
                st.caption(f"Weight in overall score: {weight}%")

            with col_b:
                st.markdown(f"**{result.get('summary', '')}**")
                findings = result.get("findings", [])
                if findings:
                    st.markdown("**Findings:**")
                    for f in findings:
                        st.markdown(f"- {f}")

            issues = result.get("issues", [])
            if issues:
                st.markdown("**Issues:**")
                for iss in issues:
                    sev = iss.get("severity", "info")
                    icon_s = {"critical": "🔴", "warning": "🟡", "info": "🔵"}.get(sev, "⚪")
                    st.markdown(f"{icon_s} **{iss.get('issue', '')}**")
                    st.caption(f"  Recommendation: {iss.get('recommendation', '')}")

            positives = result.get("positive", [])
            if positives:
                st.markdown("**Strengths:**")
                for p in positives:
                    st.markdown(f"✅ {p}")


if __name__ == "__main__":
    main()
