import streamlit as st
import subprocess
import sys
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
    {"name": "SoMe",      "script": "generate_some_report(o).py", "output_pptx": "SoMe_Report.pptx",                    "icon": "📱", "args": ["SoMe_Data(2).xlsx"]},
    {"name": "GCC Pulse", "script": "gcc_pulse_automation.py",     "output_pptx": "GCC-pulse data_GCC_Pulse_Report.pptx", "icon": "📡", "args": ["GCC-pulse data.xlsx"]},
    {"name": "SFMC",      "script": "sfmc_reportwt(1).py",         "output_pptx": "SFMC_Report.pptx",                    "icon": "✉️", "args": ["SFMC-data.xlsx"]},
    {"name": "REE",       "script": "generate_ree_reportnw.py",    "output_pptx": "REE-data_REE_Report.pptx",            "icon": "⚡", "args": ["REE-data.xlsx"]},
]

MERGED_OUTPUT = "Performance_Deck.pptx"

# ── Helpers ───────────────────────────────────────────────────────────────────
def run_script_in_tmpdir(script_name: str, excel_arg: str, tmp_dir: Path) -> tuple:
    """
    Copy script + Excel into tmp_dir, patch all file-path constants so
    everything reads & writes inside tmp_dir (writable on Streamlit Cloud).
    """
    import shutil as _sh2, os as _os2

    src_script = BASE_DIR / script_name
    if not src_script.exists():
        return False, f"Script not found: {script_name}"

    # 1. Copy Excel into tmp_dir first
    if excel_arg:
        xl_src = BASE_DIR / excel_arg
        if not xl_src.exists():
            return False, f"Excel not found: {excel_arg}"
        _sh2.copy2(str(xl_src), str(tmp_dir / excel_arg))

    # 2. Read & patch script source so all path roots point to tmp_dir
    code = src_script.read_text(encoding="utf-8")
    forced = str(tmp_dir)

    # Inject forced-dir variable right after the first import block
    patch_header = "import os as _os_pd, sys as _sys_pd\n_FORCED_DIR = " + repr(forced) + "\n"
    # Replace every common pattern scripts use to locate themselves
    replacements = [
        ('os.path.dirname(os.path.abspath(__file__))', '_FORCED_DIR'),
        ('os.path.dirname(os.path.abspath(inp))',      '_FORCED_DIR'),
        ('os.path.dirname(os.path.realpath(__file__))','_FORCED_DIR'),
    ]
    for old_pat, new_pat in replacements:
        code = code.replace(old_pat, new_pat)

    code = patch_header + code

    patched_name = "_pd_" + script_name.replace("(","_").replace(")","_").replace(" ","_")
    patched_path = tmp_dir / patched_name
    patched_path.write_text(code, encoding="utf-8")

    # 3. Run patched script from tmp_dir; pass excel arg as relative name
    cmd = [sys.executable, str(patched_path)]
    if excel_arg:
        cmd.append(excel_arg)

    env = {**_os2.environ, "PYTHONPATH": str(BASE_DIR)}

    try:
        result = subprocess.run(
            cmd,
            cwd=str(tmp_dir),
            capture_output=True,
            text=True,
            timeout=300,
            env=env,
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
    """
    Robust merge using zip-level manipulation.
    Copies every slide and all its dependencies (images, charts, media)
    directly from the source zip into the output zip.
    """
    import zipfile
    import shutil
    import re
    from lxml import etree

    # Start with a copy of the first file as our base
    shutil.copy2(str(pptx_paths[0]), str(output_path))

    for pptx_path in pptx_paths[1:]:
        _append_pptx(str(output_path), str(pptx_path))


def _append_pptx(base_path, src_path):
    """Append all slides from src_path into base_path in-place."""
    import zipfile, shutil, re, os, tempfile
    from lxml import etree

    tmp = base_path + ".tmp"
    shutil.copy2(base_path, tmp)

    with zipfile.ZipFile(tmp, "r") as base_zip, \
         zipfile.ZipFile(src_path, "r") as src_zip, \
         zipfile.ZipFile(base_path, "w", zipfile.ZIP_DEFLATED) as out_zip:

        # ── Read base presentation.xml to find current slide count ──
        prs_xml = etree.fromstring(base_zip.read("ppt/presentation.xml"))
        nsmap = {
            "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        }
        sldIdLst = prs_xml.find(".//p:sldIdLst", nsmap)
        if sldIdLst is None:
            sldIdLst = etree.SubElement(prs_xml, "{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst")

        existing_ids = [int(el.get("id")) for el in sldIdLst]
        next_id = max(existing_ids, default=255) + 1

        # ── Read base presentation.xml.rels ──
        prs_rels_path = "ppt/_rels/presentation.xml.rels"
        prs_rels_xml  = etree.fromstring(base_zip.read(prs_rels_path))
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        existing_rids = [el.get("Id") for el in prs_rels_xml]
        next_rid_num  = max(
            (int(re.sub(r"\D", "", r)) for r in existing_rids if re.sub(r"\D", "", r)),
            default=100
        ) + 1

        # ── Collect all files already in base ──
        base_files = set(base_zip.namelist())

        # ── Copy all non-slide files from base into out ──
        skip_rewrite = {"ppt/presentation.xml", prs_rels_path}
        for item in base_zip.namelist():
            if item not in skip_rewrite:
                out_zip.writestr(item, base_zip.read(item))

        # ── Find slides in src ──
        src_prs_xml  = etree.fromstring(src_zip.read("ppt/presentation.xml"))
        src_sldIdLst = src_prs_xml.find(".//p:sldIdLst", nsmap)
        if src_sldIdLst is None:
            os.remove(tmp)
            return

        src_prs_rels_xml = etree.fromstring(src_zip.read("ppt/_rels/presentation.xml.rels"))
        # Map rId -> slide path in src
        src_rid_to_target = {}
        for el in src_prs_rels_xml:
            rid    = el.get("Id")
            target = el.get("Target")
            rtype  = el.get("Type", "")
            if "slide" in rtype and "slideLayout" not in rtype and "slideMaster" not in rtype:
                src_rid_to_target[rid] = target  # e.g. "slides/slide1.xml"

        # ── For each slide in src, copy it + its deps into out ──
        all_src_files = set(src_zip.namelist())

        for sldId_el in src_sldIdLst:
            rid = sldId_el.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if rid not in src_rid_to_target:
                continue

            slide_rel_target = src_rid_to_target[rid]   # e.g. "slides/slide1.xml"
            slide_path_in_zip = f"ppt/{slide_rel_target}"  # e.g. "ppt/slides/slide1.xml"

            # Pick a new slide filename that doesn't clash
            slide_filename = os.path.basename(slide_path_in_zip)
            new_slide_name = slide_filename
            counter = 900
            while f"ppt/slides/{new_slide_name}" in base_files or f"ppt/slides/{new_slide_name}" in {f"ppt/slides/{x}" for x in [new_slide_name]}:
                counter += 1
                new_slide_name = f"slide{counter}.xml"
                if f"ppt/slides/{new_slide_name}" not in base_files:
                    break

            new_slide_zip_path = f"ppt/slides/{new_slide_name}"
            base_files.add(new_slide_zip_path)

            # Copy slide XML
            slide_data = src_zip.read(slide_path_in_zip)
            out_zip.writestr(new_slide_zip_path, slide_data)

            # Copy slide rels if present
            slide_rels_src = f"ppt/slides/_rels/{slide_filename}.rels"
            new_slide_rels_path = f"ppt/slides/_rels/{new_slide_name}.rels"
            if slide_rels_src in all_src_files:
                rels_data = src_zip.read(slide_rels_src)
                # Copy all media/image dependencies referenced in the rels
                rels_xml = etree.fromstring(rels_data)
                for rel_el in rels_xml:
                    dep_target = rel_el.get("Target", "")
                    dep_type   = rel_el.get("Type", "")
                    # Resolve relative path
                    if dep_target.startswith("../"):
                        dep_zip_path = "ppt/" + dep_target[3:]
                    elif dep_target.startswith("/"):
                        dep_zip_path = dep_target.lstrip("/")
                    else:
                        dep_zip_path = f"ppt/slides/{dep_target}"

                    if dep_zip_path in all_src_files and dep_zip_path not in base_files:
                        try:
                            out_zip.writestr(dep_zip_path, src_zip.read(dep_zip_path))
                            base_files.add(dep_zip_path)
                        except Exception:
                            pass

                out_zip.writestr(new_slide_rels_path, rels_data)

            # ── Register slide in presentation.xml ──
            new_sldId = etree.SubElement(sldIdLst, "{http://schemas.openxmlformats.org/presentationml/2006/main}sldId")
            new_sldId.set("id", str(next_id))
            new_rid = f"rId{next_rid_num}"
            new_sldId.set("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", new_rid)
            next_id      += 1
            next_rid_num += 1

            # ── Register slide rel in presentation.xml.rels ──
            new_rel = etree.SubElement(prs_rels_xml, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
            new_rel.set("Id", new_rid)
            new_rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
            new_rel.set("Target", f"slides/{new_slide_name}")

        # ── Write updated presentation.xml and rels ──
        out_zip.writestr("ppt/presentation.xml",
                         etree.tostring(prs_xml, xml_declaration=True,
                                        encoding="UTF-8", standalone=True))
        out_zip.writestr(prs_rels_path,
                         etree.tostring(prs_rels_xml, xml_declaration=True,
                                        encoding="UTF-8", standalone=True))

    os.remove(tmp)

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

    import tempfile as _tf, shutil as _sh

    # All outputs go to /tmp — Streamlit Cloud repo is read-only
    tmp_dir = Path(_tf.mkdtemp())

    for platform in PLATFORMS:
        name   = platform["name"]
        script = platform["script"]
        excel  = (platform.get("args") or [None])[0]  # Excel filename or None

        set_status(name, "running")
        log_lines.append(f"▶  [{name}] Running {script} …")
        update_log()

        ok, msg = run_script_in_tmpdir(script, excel, tmp_dir)
        out = tmp_dir / platform["output_pptx"]

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
            out_path = tmp_dir / MERGED_OUTPUT
            merge_presentations(pptx_paths, out_path)
            # Read into memory immediately — never rely on disk for download
            st.session_state.merged_bytes = out_path.read_bytes()
            kb = len(st.session_state.merged_bytes) // 1024
            log_lines.append(f"✅ Merged → {MERGED_OUTPUT}  ({kb} KB)")
            st.session_state.generated = True
        except Exception as e:
            log_lines.append(f"❌ Merge failed: {e}")
            all_errors.append(("Merge", str(e)))
    else:
        log_lines.append("❌ No decks to merge — all scripts failed.")

    update_log()
    st.session_state.run_errors = all_errors

    # Cleanup tmp
    try:
        _sh.rmtree(tmp_dir)
    except Exception:
        pass

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
        data=bytes(st.session_state.merged_bytes),
        file_name="Performance_Deck.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        key="dl_btn",
        use_container_width=True,
    )
