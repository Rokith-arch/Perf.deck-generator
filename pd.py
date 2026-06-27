import streamlit as st
import subprocess
import sys
import copy
from pathlib import Path
from pptx import Presentation

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Performance Deck Generator",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── CSS (static only — no dynamic HTML injection during run) ──────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Space+Grotesk:wght@400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.stApp {
    background: linear-gradient(135deg, #020b18 0%, #041425 50%, #051e30 100%);
    min-height: 100vh;
}

#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 800px; }

.hero {
    text-align: center;
    padding: 2.5rem 1rem 1.2rem;
}
.hero-badge {
    display: inline-block;
    background: linear-gradient(90deg, #0d4f3c, #0a7c5e);
    color: #2fffc2;
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 0.2em;
    text-transform: uppercase;
    padding: 0.3rem 1rem;
    border-radius: 20px;
    margin-bottom: 1.2rem;
    border: 1px solid #0a7c5e55;
}
.hero-title {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 2.5rem;
    font-weight: 700;
    color: #ffffff;
    line-height: 1.15;
    margin: 0 0 0.6rem;
    letter-spacing: -0.02em;
}
.hero-title span {
    background: linear-gradient(90deg, #2fffc2 0%, #00b4d8 60%, #0077b6 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}
.hero-sub {
    color: #7a9bb5;
    font-size: 0.92rem;
    line-height: 1.6;
}
.grad-divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, #0a7c5e55, #00b4d855, transparent);
    margin: 1.6rem 0;
}
.platform-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 0.85rem;
    margin: 1rem 0;
}
.platform-card {
    background: linear-gradient(135deg, #071c2e 0%, #0a2540 100%);
    border: 1px solid #0e3a5a;
    border-radius: 12px;
    padding: 0.9rem 1.1rem;
    display: flex;
    align-items: center;
    gap: 0.7rem;
}
.platform-icon { font-size: 1.4rem; width: 2rem; text-align: center; flex-shrink: 0; }
.platform-name { font-weight: 600; color: #cde7f5; font-size: 0.87rem; }
.platform-file { color: #4a7c99; font-size: 0.7rem; margin-top: 0.1rem; font-family: monospace; }

.order-row {
    display: flex;
    gap: 0.4rem;
    align-items: center;
    justify-content: center;
    flex-wrap: wrap;
    margin: 0.8rem 0 1.4rem;
}
.order-pill {
    background: #071c2e;
    border: 1px solid #0e3a5a;
    border-radius: 20px;
    padding: 0.28rem 0.8rem;
    color: #7abfdf;
    font-size: 0.73rem;
    font-weight: 500;
    display: flex;
    align-items: center;
    gap: 0.35rem;
}
.order-pill .num {
    background: linear-gradient(135deg, #0a7c5e, #00b4d8);
    color: white;
    border-radius: 50%;
    width: 1.15rem;
    height: 1.15rem;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-size: 0.63rem;
    font-weight: 700;
}
.arr { color: #1e4d6a; }

/* Generate button */
.stButton > button {
    width: 100%;
    background: linear-gradient(135deg, #0a7c5e 0%, #00897b 50%, #0077b6 100%);
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.85rem 2rem !important;
    font-size: 1rem !important;
    font-weight: 700 !important;
    font-family: 'Space Grotesk', sans-serif !important;
    letter-spacing: 0.04em !important;
    box-shadow: 0 4px 24px #0a7c5e33;
    text-transform: uppercase;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #0d9e78 0%, #00a693 50%, #0096d6 100%) !important;
    box-shadow: 0 8px 32px #0a7c5e55 !important;
    transform: translateY(-2px);
}

/* Download button */
.stDownloadButton > button {
    width: 100%;
    background: #041e18 !important;
    color: #2fffc2 !important;
    border: 1.5px solid #0a7c5e !important;
    border-radius: 10px !important;
    padding: 0.85rem 2rem !important;
    font-size: 1rem !important;
    font-weight: 700 !important;
    font-family: 'Space Grotesk', sans-serif !important;
    letter-spacing: 0.04em !important;
    text-transform: uppercase;
    box-shadow: 0 4px 16px #0a7c5e22;
}
.stDownloadButton > button:hover {
    border-color: #2fffc2 !important;
    box-shadow: 0 6px 24px #0a7c5e44 !important;
    transform: translateY(-1px);
}

/* Success box */
.success-box {
    background: linear-gradient(135deg, #041e18 0%, #051a28 100%);
    border: 1px solid #0a7c5e;
    border-radius: 12px;
    padding: 1.5rem;
    text-align: center;
    margin: 1rem 0;
}
.success-box .s-emoji { font-size: 2.2rem; }
.success-box .s-title { color: #2fffc2; font-size: 1.05rem; font-weight: 700; margin: 0.4rem 0 0.2rem; }
.success-box .s-sub { color: #4a9e7a; font-size: 0.8rem; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent

PLATFORMS = [
    {"name": "SoMe",      "script": "generate_some_report(o).py", "output_pptx": "SoMe_Report.pptx",                    "icon": "📱", "args": None},
    {"name": "GCC Pulse", "script": "gcc_pulse_automation.py",     "output_pptx": "GCC-pulse data_GCC_Pulse_Report.pptx", "icon": "📡", "args": ["GCC-pulse data.xlsx"]},
    {"name": "SFMC",      "script": "sfmc_reportwt(1).py",         "output_pptx": "SFMC_Report.pptx",                    "icon": "✉️", "args": None},
    {"name": "REE",       "script": "generate_ree_reportnw.py",    "output_pptx": "REE-data_REE_Report.pptx",            "icon": "⚡", "args": None},
]

MERGED_OUTPUT = "Performance_Deck.pptx"

# ── Helpers ───────────────────────────────────────────────────────────────────
def run_script(script_name: str, extra_args: list = None):
    script_path = BASE_DIR / script_name
    if not script_path.exists():
        return False, f"Script not found: {script_name}"
    cmd = [sys.executable, str(script_path)] + (extra_args or [])
    try:
        result = subprocess.run(
            cmd,
            cwd=str(BASE_DIR),
            capture_output=True,
            text=True,
            timeout=300,
        )
        if result.returncode != 0:
            err = (result.stderr or result.stdout).strip()
            return False, err[-800:]
        return True, result.stdout.strip()
    except subprocess.TimeoutExpired:
        return False, "Script timed out after 5 minutes."
    except Exception as e:
        return False, str(e)


def merge_presentations(pptx_paths, output_path):
    first = Presentation(str(pptx_paths[0]))
    merged = Presentation()
    merged.slide_width  = first.slide_width
    merged.slide_height = first.slide_height

    for pptx_path in pptx_paths:
        src = Presentation(str(pptx_path))
        for slide in src.slides:
            blank_layout = merged.slide_layouts[6]
            new_slide = merged.slides.add_slide(blank_layout)
            sp_tree     = slide.shapes._spTree
            new_sp_tree = new_slide.shapes._spTree
            for elem in list(new_sp_tree):
                new_sp_tree.remove(elem)
            for elem in sp_tree:
                new_sp_tree.append(copy.deepcopy(elem))
            for rel in slide.part.rels.values():
                if "image" in rel.reltype:
                    try:
                        new_slide.part.relate_to(rel.target_part, rel.reltype)
                    except Exception:
                        pass
    merged.save(str(output_path))

# ── Static UI ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <div class="hero-badge">📊 Automated Reporting Suite</div>
    <div class="hero-title">Performance Deck<br><span>Generator</span></div>
    <div class="hero-sub">Runs all four platform scripts, merges the decks,<br>and delivers one polished PowerPoint — in one click.</div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="grad-divider"></div>', unsafe_allow_html=True)

st.markdown("""
<div class="platform-grid">
    <div class="platform-card"><div class="platform-icon">📱</div><div><div class="platform-name">SoMe</div><div class="platform-file">generate_some_report(o).py</div></div></div>
    <div class="platform-card"><div class="platform-icon">📡</div><div><div class="platform-name">GCC Pulse</div><div class="platform-file">gcc_pulse_automation.py</div></div></div>
    <div class="platform-card"><div class="platform-icon">✉️</div><div><div class="platform-name">SFMC</div><div class="platform-file">sfmc_reportwt(1).py</div></div></div>
    <div class="platform-card"><div class="platform-icon">⚡</div><div><div class="platform-name">REE</div><div class="platform-file">generate_ree_reportnw.py</div></div></div>
</div>
<p style="text-align:center;color:#4a7c99;font-size:0.7rem;font-weight:600;letter-spacing:0.12em;text-transform:uppercase;margin:1rem 0 0.3rem;">Merge Order</p>
<div class="order-row">
    <div class="order-pill"><span class="num">1</span>SoMe</div>
    <span class="arr">›</span>
    <div class="order-pill"><span class="num">2</span>GCC Pulse</div>
    <span class="arr">›</span>
    <div class="order-pill"><span class="num">3</span>SFMC</div>
    <span class="arr">›</span>
    <div class="order-pill"><span class="num">4</span>REE</div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="grad-divider"></div>', unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
for key, default in [("generated", False), ("merged_bytes", None), ("run_errors", [])]:
    if key not in st.session_state:
        st.session_state[key] = default

# ── Generate ──────────────────────────────────────────────────────────────────
if st.button("🚀  Generate Performance Deck", key="gen_btn"):
    st.session_state.generated  = False
    st.session_state.merged_bytes = None
    st.session_state.run_errors  = []

    pptx_paths = []
    all_errors = []

    # Status columns (native Streamlit — no dynamic HTML)
    st.markdown("#### Running scripts…")
    cols = st.columns(4)
    status_slots = {p["name"]: cols[i].empty() for i, p in enumerate(PLATFORMS)}

    def set_status(name, state):
        icon  = {"pending": "⬜", "running": "⏳", "done": "✅", "error": "❌"}[state]
        color = {"pending": "#4a7c99", "running": "#00e5ff", "done": "#2fffc2", "error": "#ff6b6b"}[state]
        p = next(x for x in PLATFORMS if x["name"] == name)
        status_slots[name].markdown(
            f"<div style='text-align:center;background:#071c2e;border:1px solid #0e3a5a;"
            f"border-radius:10px;padding:0.7rem 0.3rem;'>"
            f"<div style='font-size:0.65rem;font-weight:700;letter-spacing:0.1em;"
            f"text-transform:uppercase;color:{color}'>{name}</div>"
            f"<div style='font-size:1.4rem;margin-top:0.3rem'>{icon}</div></div>",
            unsafe_allow_html=True,
        )

    for p in PLATFORMS:
        set_status(p["name"], "pending")

    log = st.empty()
    log_lines = ["ℹ️  Starting Performance Deck generation…"]

    def update_log():
        log.code("\n".join(log_lines), language=None)

    update_log()

    for platform in PLATFORMS:
        name   = platform["name"]
        script = platform["script"]
        out    = BASE_DIR / platform["output_pptx"]

        set_status(name, "running")
        log_lines.append(f"▶  [{name}] Running {script} …")
        update_log()

        ok, msg = run_script(script, platform.get("args"))

        if ok and out.exists():
            pptx_paths.append(out)
            set_status(name, "done")
            log_lines.append(f"✅ [{name}] Done → {platform['output_pptx']}")
        else:
            set_status(name, "error")
            reason = msg if not ok else f"Output not found: {platform['output_pptx']}"
            log_lines.append(f"❌ [{name}] FAILED")
            for line in reason.splitlines()[-4:]:
                log_lines.append(f"   {line}")
            all_errors.append((name, reason))

        update_log()

    # Merge
    if pptx_paths:
        log_lines.append(f"ℹ️  Merging {len(pptx_paths)} deck(s)…")
        update_log()
        try:
            out_path = BASE_DIR / MERGED_OUTPUT
            merge_presentations(pptx_paths, out_path)
            with open(out_path, "rb") as f:
                st.session_state.merged_bytes = f.read()
            log_lines.append(f"✅ Merged → {MERGED_OUTPUT}  ({len(st.session_state.merged_bytes)//1024} KB)")
            st.session_state.generated = True
        except Exception as e:
            log_lines.append(f"❌ Merge failed: {e}")
            all_errors.append(("Merge", str(e)))
    else:
        log_lines.append("❌ No decks to merge — all scripts failed.")

    update_log()
    st.session_state.run_errors = all_errors

    if all_errors:
        with st.expander("⚠️ Errors — click to expand", expanded=True):
            for name, err in all_errors:
                st.error(f"**{name}:** {err[:400]}")

# ── Download ──────────────────────────────────────────────────────────────────
if st.session_state.generated and st.session_state.merged_bytes:
    st.markdown("""
    <div class="success-box">
        <div class="s-emoji">🎉</div>
        <div class="s-title">Performance Deck Ready</div>
        <div class="s-sub">Merged in order: SoMe → GCC Pulse → SFMC → REE</div>
    </div>
    """, unsafe_allow_html=True)

    st.download_button(
        label="⬇️  Download Performance_Deck.pptx",
        data=st.session_state.merged_bytes,
        file_name="Performance_Deck.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        key="dl_btn",
    )
