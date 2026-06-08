"""
app.py — Document Intelligence Annotator
Mobius Knowledge Services

Login-protected Streamlit app deployable on Streamlit Community Cloud (free tier).
Users are defined in .streamlit/secrets.toml — no database needed.
"""

from __future__ import annotations
import os, sys, io, tempfile, hashlib
from pathlib import Path

import streamlit as st
from PIL import Image

sys.path.insert(0, str(Path(__file__).parent))
from core.kyc_extractor import DocumentExtractor, load_image_b64, SUPPORTED_TYPES
from core.excel_builder import ExcelBuilder

# ── Page config (must be first Streamlit call) ────────────────────────────────
st.set_page_config(
    page_title="Document Annotator | Mobius",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Helpers ───────────────────────────────────────────────────────────────────
def _secret(k, fb=""):
    try:    return st.secrets[k]
    except: return os.environ.get(k, fb)

def _hash(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()

# ── Brand ─────────────────────────────────────────────────────────────────────
NAVY   = "#1B2A4A"
ORANGE = "#F47920"
BG     = "#F5F7FA"
MUTED  = "#6C757D"
BORDER = "#DDE3EC"

CATEGORY_THEME = {
    "identity":   {"hdr": "#1B2A4A", "sec": "#EFF6FF", "badge": "#BFDBFE", "badge_text": "#1E3A8A", "icon": "🪪"},
    "academic":   {"hdr": "#14532D", "sec": "#F0FDF4", "badge": "#BBF7D0", "badge_text": "#14532D", "icon": "🎓"},
    "employment": {"hdr": "#92400E", "sec": "#FFFBEB", "badge": "#FDE68A", "badge_text": "#78350F", "icon": "💼"},
    "financial":  {"hdr": "#1E40AF", "sec": "#EFF6FF", "badge": "#BFDBFE", "badge_text": "#1E3A8A", "icon": "💰"},
    "legal":      {"hdr": "#6B21A8", "sec": "#FAF5FF", "badge": "#E9D5FF", "badge_text": "#6B21A8", "icon": "⚖️"},
    "medical":    {"hdr": "#991B1B", "sec": "#FFF1F2", "badge": "#FECDD3", "badge_text": "#9F1239", "icon": "🏥"},
    "other":      {"hdr": "#374151", "sec": "#F9FAFB", "badge": "#E5E7EB", "badge_text": "#374151", "icon": "📄"},
}

# ── Global CSS ────────────────────────────────────────────────────────────────
st.markdown(f"""<style>
html,body,[data-testid="stAppViewContainer"]{{background:{BG}!important;}}
.block-container{{padding:0!important;max-width:100%!important;}}

/* ── Header ── */
.mobius-header{{background:{NAVY};padding:14px 40px;display:flex;align-items:center;justify-content:space-between;}}
.mobius-header .pname{{color:#fff;font-size:20px;font-weight:600;letter-spacing:.3px;margin-left:16px;}}
.mobius-header .tag{{color:#A8B8D0;font-size:13px;margin-left:10px;}}
.hleft{{display:flex;align-items:center;}}
.user-pill{{background:rgba(255,255,255,0.12);color:#fff;padding:4px 14px;border-radius:20px;font-size:12px;}}

/* ── Login ── */
.login-wrap{{display:flex;align-items:center;justify-content:center;min-height:80vh;}}
.login-card{{background:#fff;border:1px solid {BORDER};border-radius:14px;padding:48px 44px;width:400px;box-shadow:0 4px 24px rgba(0,0,0,0.08);}}
.login-logo{{text-align:center;margin-bottom:24px;}}
.login-title{{font-size:22px;font-weight:700;color:{NAVY};text-align:center;margin-bottom:4px;}}
.login-sub{{font-size:13px;color:{MUTED};text-align:center;margin-bottom:28px;}}
.login-err{{background:#FEE2E2;color:#991B1B;border-radius:8px;padding:10px 14px;font-size:13px;margin-bottom:16px;}}

/* ── Main layout ── */
.main-wrap{{padding:28px 40px 40px 40px;}}
.sc{{background:#fff;border:1px solid {BORDER};border-radius:10px;padding:24px 28px;margin-bottom:20px;}}
.st{{font-size:15px;font-weight:600;color:{NAVY};margin-bottom:14px;padding-bottom:10px;border-bottom:2px solid {ORANGE};display:inline-block;}}

/* ── Document cards ── */
.doc-banner{{border-radius:10px 10px 0 0;padding:16px 20px 12px 20px;}}
.doc-title{{font-size:20px;font-weight:700;color:#fff;}}
.doc-meta{{font-size:12px;color:rgba(255,255,255,0.75);margin-top:4px;}}
.doc-body{{background:#fff;border:1px solid {BORDER};border-top:none;border-radius:0 0 10px 10px;padding:0 0 8px 0;margin-bottom:24px;}}
.section-hdr{{font-size:11px;font-weight:700;letter-spacing:1.2px;padding:10px 20px 6px 20px;margin-top:4px;}}
.attr-row{{display:flex;align-items:flex-start;padding:7px 20px;border-bottom:1px solid {BORDER};gap:0;}}
.attr-row:last-child{{border-bottom:none;}}
.attr-name{{font-size:12px;font-weight:600;color:{MUTED};width:220px;flex-shrink:0;padding-top:2px;line-height:1.4;}}
.attr-value{{font-size:13px;font-weight:500;color:#111827;flex:1;line-height:1.5;word-break:break-word;}}
.conf-dot{{width:8px;height:8px;border-radius:50%;flex-shrink:0;margin-top:5px;margin-left:12px;}}
.conf-high{{background:#22C55E;}}
.conf-medium{{background:#F59E0B;}}
.conf-low{{background:#EF4444;}}
.flag-chip{{font-size:11px;padding:2px 10px;border-radius:20px;font-weight:500;margin:2px;}}
.flag-on{{background:#DCFCE7;color:#166534;}}
.flag-off{{background:#F1F5F9;color:#9CA3AF;}}
.flag-warn{{background:#FEE2E2;color:#991B1B;}}
.raw-block{{background:#0D1117;color:#E6EDF3;font-family:'Courier New',monospace;font-size:11px;border-radius:8px;padding:14px 16px;max-height:180px;overflow-y:auto;line-height:1.7;white-space:pre-wrap;margin:0 20px 12px 20px;}}
.badge{{display:inline-block;padding:3px 12px;border-radius:20px;font-size:12px;font-weight:600;margin-left:10px;vertical-align:middle;}}

/* ── Footer ── */
.mobius-footer{{text-align:center;padding:20px 40px;border-top:1px solid {BORDER};font-size:12px;color:{MUTED};background:#fff;margin-top:40px;}}
.mobius-footer a{{color:{ORANGE};text-decoration:none;}}

/* ── Streamlit overrides ── */
#MainMenu,footer,header{{visibility:hidden;}}
[data-testid="collapsedControl"]{{display:none;}}
div[data-testid="stButton"]>button[kind="primary"]{{background-color:{ORANGE}!important;border:none!important;color:#fff!important;font-weight:600!important;border-radius:6px!important;}}
div[data-testid="stButton"]>button[kind="primary"]:hover{{background-color:#D9640A!important;}}
div[data-testid="stProgressBar"]>div>div{{background-color:{ORANGE}!important;}}
div[data-testid="stTextInput"] input{{border-radius:6px!important;border-color:{BORDER}!important;}}
div[data-testid="stTextInput"] input:focus{{border-color:{ORANGE}!important;box-shadow:0 0 0 2px rgba(244,121,32,0.15)!important;}}
</style>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# AUTH
# ══════════════════════════════════════════════════════════════════════════════

def _load_users() -> dict[str, dict]:
    """
    Load users from secrets.toml.
    Format in secrets.toml:
        [users.alice]
        password = "hashed_or_plain"
        display_name = "Alice Smith"
        role = "admin"          # optional
    """
    try:
        raw = st.secrets.get("users", {})
        return dict(raw)
    except Exception:
        return {}

def _verify(username: str, password: str, users: dict) -> bool:
    u = users.get(username)
    if not u:
        return False
    stored = u.get("password", "")
    # Accept SHA-256 hash OR plain text (plain for easy setup)
    return stored == password or stored == _hash(password)

def _show_login():
    st.markdown("""<div class="login-wrap">""", unsafe_allow_html=True)
    _, col, _ = st.columns([1, 1.2, 1])
    with col:
        st.markdown(f"""
        <div class="login-card">
          <div class="login-logo">
            <img src="https://www.mobiusservices.com/img/mobius_logo.svg"
                 height="40" alt="Mobius"
                 style="filter: brightness(0) saturate(100%) invert(18%) sepia(47%) saturate(700%) hue-rotate(190deg);">
          </div>
          <div class="login-title">Document Annotator</div>
          <div class="login-sub">Sign in to continue</div>
        </div>
        """, unsafe_allow_html=True)

        if st.session_state.get("login_error"):
            st.error("Incorrect username or password. Please try again.")

        username = st.text_input("Username", placeholder="Enter your username",
                                  key="login_username")
        password = st.text_input("Password", type="password",
                                  placeholder="Enter your password",
                                  key="login_password")
        st.markdown("<br>", unsafe_allow_html=True)
        login_btn = st.button("Sign In", type="primary", use_container_width=True)

        if login_btn:
            users = _load_users()
            if _verify(username.strip(), password, users):
                st.session_state["authenticated"] = True
                st.session_state["current_user"]  = username.strip()
                display = users[username.strip()].get("display_name", username.strip())
                st.session_state["display_name"]  = display
                st.session_state["login_error"]   = False
                st.rerun()
            else:
                st.session_state["login_error"] = True
                st.rerun()

        st.markdown(f"""<div style="text-align:center;margin-top:20px;font-size:11px;color:{MUTED};">
        Contact your administrator for access.</div>""", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)


# ── Session init ──────────────────────────────────────────────────────────────
for k, v in [
    ("authenticated", False),
    ("current_user",  ""),
    ("display_name",  ""),
    ("login_error",   False),
    ("results",       []),
    ("running",       False),
]:
    if k not in st.session_state:
        st.session_state[k] = v

# ── Gate: show login if not authenticated ─────────────────────────────────────
if not st.session_state["authenticated"]:
    _show_login()
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# MAIN APP  (only reaches here after successful login)
# ══════════════════════════════════════════════════════════════════════════════

OPENAI_KEY = _secret("OPENAI_API_KEY")

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(f"""<div class="mobius-header">
  <div class="hleft">
    <img src="https://www.mobiusservices.com/img/mobius_logo.svg" height="36"
         alt="Mobius" style="filter:brightness(0) invert(1);">
    <span class="pname">Document Intelligence Annotator</span>
    <span class="tag">Any Document → Structured Data Extraction</span>
  </div>
  <div style="display:flex;align-items:center;gap:12px;">
    <span class="user-pill">👤 {st.session_state['display_name']}</span>
  </div>
</div>""", unsafe_allow_html=True)

# ── Logout in sidebar ─────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"**Signed in as**  \n{st.session_state['display_name']}")
    st.markdown("---")
    if st.button("🚪  Sign Out", use_container_width=True):
        for k in ["authenticated", "current_user", "display_name",
                  "results", "running", "login_error"]:
            st.session_state[k] = False if isinstance(st.session_state[k], bool) else \
                                   []    if isinstance(st.session_state[k], list)  else ""
        st.rerun()

st.markdown('<div class="main-wrap">', unsafe_allow_html=True)

# ── Upload section ────────────────────────────────────────────────────────────
st.markdown('<div class="sc">', unsafe_allow_html=True)
st.markdown('<span class="st">Upload Documents</span>', unsafe_allow_html=True)
col_up, col_tip = st.columns([3, 1])
with col_up:
    uploaded_files = st.file_uploader(
        "Upload documents",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )
    st.caption(
        "Supports any document — Aadhaar · PAN · Passport · Driving Licence · Voter ID · "
        "Marksheets · Salary Slips · Appointment Letters · Experience Letters · "
        "Bank Statements · Company IDs · and more"
    )
with col_tip:
    st.markdown(f"""<div style="background:#E8EDF4;border-radius:8px;padding:14px;font-size:12px;color:{NAVY};">
    <b>💡 Tips</b><br>
    • Upload multiple docs at once<br>
    • JPG / PNG / WEBP supported<br>
    • Works on scans &amp; photos<br>
    • AI auto-detects document type
    </div>""", unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ── Action buttons ─────────────────────────────────────────────────────────────
b1, b2, b3 = st.columns([2, 2, 1])
with b1:
    run_clicked = st.button(
        "▶  Extract & Annotate",
        type="primary",
        disabled=(not uploaded_files or st.session_state.running),
        use_container_width=True,
    )
with b2:
    if st.session_state.results:
        out_dir = Path(tempfile.gettempdir()) / "doc_outputs"
        out_dir.mkdir(exist_ok=True)
        out_path = out_dir / "document_annotations.xlsx"
        ExcelBuilder().build(st.session_state.results, out_path)
        with open(out_path, "rb") as f:
            st.download_button(
                "⬇  Download Excel",
                data=f.read(),
                file_name="document_annotations.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
with b3:
    if st.button("↺  Reset", use_container_width=True):
        st.session_state.results = []
        st.session_state.running = False
        st.rerun()

prog = st.progress(0)
stat = st.empty()

# ── Run extraction ─────────────────────────────────────────────────────────────
if run_clicked and uploaded_files:
    if not OPENAI_KEY:
        st.error("OpenAI API key not configured. Add OPENAI_API_KEY to .streamlit/secrets.toml")
        st.stop()

    st.session_state.running = True
    st.session_state.results = []
    extractor = DocumentExtractor(api_key=OPENAI_KEY)
    total = len(uploaded_files)

    for i, uf in enumerate(uploaded_files, start=1):
        stat.caption(f"Analysing {uf.name}  ({i} of {total})…")
        prog.progress(int((i - 1) / total * 100))
        file_bytes = uf.read()
        ext = Path(uf.name).suffix.lower()
        try:
            b64, media_type = load_image_b64(file_bytes, ext)
            result = extractor.extract(b64, media_type=media_type)
        except Exception as e:
            result = {
                "_parse_error": True,
                "_error_message": str(e),
                "document_type": "Unknown",
                "document_category": "other",
                "issuer": "",
                "sections": [],
                "flags": {},
                "document_date": None,
                "validity": {},
                "raw_text": "",
            }
        result["_filename"]   = uf.name
        result["_file_bytes"] = file_bytes
        st.session_state.results.append(result)

    prog.progress(100)
    stat.success(f"✅  {total} document(s) extracted successfully.")
    st.session_state.running = False
    st.rerun()

# ── Display results ────────────────────────────────────────────────────────────
def _flag_chip(label, val, warn_if_true=False):
    cls = ("flag-warn" if (val and warn_if_true) else
           "flag-on"   if val else "flag-off")
    return f'<span class="flag-chip {cls}">{label}: {"Yes" if val else "No"}</span>'

def _conf_dot(conf):
    cls = {"high": "conf-high", "medium": "conf-medium", "low": "conf-low"}.get(
        (conf or "high").lower(), "conf-high"
    )
    return f'<span class="conf-dot {cls}" title="{(conf or "high").title()} confidence"></span>'

if st.session_state.results:
    st.markdown("---")
    for r in st.session_state.results:
        cat     = r.get("document_category", "other")
        theme   = CATEGORY_THEME.get(cat, CATEGORY_THEME["other"])
        icon    = theme["icon"]
        hdr_col = theme["hdr"]
        sec_bg  = theme["sec"]
        bdg_bg  = theme["badge"]
        bdg_txt = theme["badge_text"]

        doc_type   = r.get("document_type", "Unknown Document")
        issuer     = r.get("issuer", "")
        doc_date   = r.get("document_date", "")
        filename   = r.get("_filename", "")
        sections   = r.get("sections", [])
        flags      = r.get("flags", {})
        file_bytes = r.get("_file_bytes")
        validity   = r.get("validity", {})

        meta_parts = []
        if issuer:                        meta_parts.append(issuer)
        if doc_date:                      meta_parts.append(f"Date: {doc_date}")
        if validity.get("valid_until"):   meta_parts.append(f"Valid until: {validity['valid_until']}")
        meta_str = "   ·   ".join(meta_parts)

        st.markdown(
            f'<div class="doc-banner" style="background:{hdr_col};">'
            f'<div class="doc-title">{icon}&nbsp; {doc_type}'
            f'<span class="badge" style="background:{bdg_bg};color:{bdg_txt};">{cat.title()}</span>'
            f'</div>'
            f'<div class="doc-meta">{meta_str or filename}</div>'
            f'</div>'
            f'<div class="doc-body">',
            unsafe_allow_html=True,
        )

        col_img, col_sections = st.columns([1, 2])

        with col_img:
            if file_bytes:
                try:
                    img = Image.open(io.BytesIO(file_bytes))
                    st.image(img, use_container_width=True)
                except Exception:
                    st.warning("Could not preview image.")
            flags_html = (
                _flag_chip("📷 Photo",     flags.get("has_photo"))     +
                _flag_chip("✍️ Signature", flags.get("has_signature")) +
                _flag_chip("🔖 Stamp",     flags.get("has_stamp"))     +
                _flag_chip("📲 QR",        flags.get("has_qr_code"))   +
                _flag_chip("🔒 Masked",    flags.get("is_masked"), warn_if_true=True)
            )
            st.markdown(f'<div style="margin-top:8px;">{flags_html}</div>',
                        unsafe_allow_html=True)
            if flags.get("masked_fields"):
                st.caption(f"Masked: {', '.join(flags['masked_fields'])}")

        with col_sections:
            if r.get("_parse_error") and not sections:
                st.error(f"Extraction failed: {r.get('_error_message', 'Unknown error')}")
            elif not sections:
                st.warning("No data extracted from this document.")
            else:
                for section in sections:
                    title = section.get("section_title", "Details")
                    attrs = section.get("attributes", [])
                    if not attrs:
                        continue
                    st.markdown(
                        f'<div class="section-hdr" style="background:{sec_bg};color:{hdr_col};">'
                        f'{title.upper()}</div>',
                        unsafe_allow_html=True,
                    )
                    rows_html = ""
                    for attr in attrs:
                        name  = attr.get("name", "")
                        value = attr.get("value", "")
                        conf  = attr.get("confidence", "high")
                        rows_html += (
                            f'<div class="attr-row">'
                            f'<div class="attr-name">{name}</div>'
                            f'<div class="attr-value">{value}</div>'
                            f'{_conf_dot(conf)}'
                            f'</div>'
                        )
                    st.markdown(rows_html, unsafe_allow_html=True)

        raw = r.get("raw_text", "")
        if raw:
            with st.expander("📄 Raw Extracted Text", expanded=False):
                st.markdown(f'<div class="raw-block">{raw}</div>', unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)

elif not st.session_state.running:
    stat.info("Upload any documents above and click **Extract & Annotate** to begin.")

# ── How it works ──────────────────────────────────────────────────────────────
with st.expander("ℹ️  How it works", expanded=False):
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown("**① Upload**\nUpload any document image — IDs, marksheets, letters, salary slips, bank statements and more.")
    c2.markdown("**② Auto-Detection**\nGPT-4o identifies the exact document type and selects the right extraction schema automatically.")
    c3.markdown("**③ Structured Extraction**\nAll fields are extracted and grouped into logical sections with clear attribute names and values.")
    c4.markdown("**④ Export**\nDownload a formatted Excel — Summary, Consolidated (one row per doc), and per-document detail sheets.")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown(f"""<div class="mobius-footer">
  Powered by&nbsp;<a href="https://www.mobiusservices.com" target="_blank">Mobius Knowledge Services</a>
  &nbsp;·&nbsp; Document Intelligence Annotator &nbsp;·&nbsp;
  © 2025 Mobius Knowledge Services. All rights reserved.
</div>""", unsafe_allow_html=True)
