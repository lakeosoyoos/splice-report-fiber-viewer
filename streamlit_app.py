"""
Splice Report + Topo Viewer — Streamlit App
============================================
Left panel: bidirectional splice report (Pass 1 + Pass 2 via splicereportmatchexfo)
Right panel: 2D OTDR trace + 3D topo viewer (unchanged from combined_app.html)

Launch:  streamlit run streamlit_app.py
"""

import os
import sys
import io
import json
import tempfile
import zipfile

import streamlit as st
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from splicereportmatchexfo import (
    load_all, discover_splices, analyze_all, scan_b_events,
    REBURN_THRESHOLD, RIBBON_SIZE,
)

# ── Page config ────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Splice Report + Fiber Viewer",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Trucordia CSS ──────────────────────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&display=swap');

#MainMenu { visibility: hidden; }
footer    { visibility: hidden; }
header    { visibility: hidden; }
.stDeployButton { display: none; }

html, body, [class*="css"], .stMarkdown, p, label, span, div {
    font-family: 'Nunito', 'Segoe UI', Arial, sans-serif !important;
}
.main .block-container {
    padding-top: 0rem !important;
    padding-bottom: 1rem;
    max-width: 100% !important;
}

.tc-topbar {
    background: #1C2526; color: #cccccc; font-size: 12px; font-weight: 600;
    font-family: 'Nunito', sans-serif; padding: 7px 36px;
    display: flex; justify-content: flex-end; gap: 28px; letter-spacing: 0.4px;
}
.tc-topbar span { color: #aaaaaa; cursor: default; }
.tc-navbar {
    background: #ffffff; border-bottom: 2px solid #eeeeee;
    padding: 14px 36px; display: flex; align-items: center; gap: 16px;
}
.tc-logo-icon {
    font-size: 28px; font-weight: 900; color: #E8461E; line-height: 1;
    transform: skewX(-8deg); display: inline-block;
}
.tc-logo-name {
    font-size: 20px; font-weight: 900; color: #1a1a1a !important;
    font-family: 'Nunito', sans-serif; letter-spacing: -0.3px;
    text-decoration: none !important;
}
.tc-navbar-spacer { flex: 1; }
.tc-contact-btn {
    border: 2px solid #E8461E; color: #1a1a1a; font-family: 'Nunito', sans-serif;
    font-weight: 800; font-size: 13px; padding: 7px 16px; cursor: default;
}

[data-testid="stSidebar"] {
    background-color: #1C2526 !important;
    border-right: 3px solid #E8461E !important;
    width: 620px !important; min-width: 620px !important; max-width: 620px !important;
}
[data-testid="stSidebar"] > div:first-child { padding-top: 0.5rem; width: 620px !important; }
[data-testid="stSidebarCollapseButton"], [data-testid="collapsedControl"],
button[data-testid="stBaseButton-headerNoPadding"] { display: none !important; }
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
    text-align: center !important;
}
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3,
[data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span,
[data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
    font-family: 'Nunito', sans-serif !important; color: #e8e8e8 !important;
}
[data-testid="stSidebar"] .stTextInput input, [data-testid="stSidebar"] .stNumberInput input {
    background: #243030 !important; border-color: #E8461E !important;
    color: #e8e8e8 !important; font-family: 'Nunito', sans-serif !important;
}
[data-testid="stSidebar"] hr { border-color: #2e3d3d !important; }
[data-testid="stSidebar"] .stCaption, [data-testid="stSidebar"] small { color: #aaa !important; }

/* Fix file uploader — hide instruction text and Add button, keep Upload button only */
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] > div {
    display: none !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] {
    justify-content: center !important;
    padding: 8px 0 !important;
}
/* Hide entire dropzone instructions + any Add button once file is loaded */
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"]:has([data-testid="stFileUploaderDeleteBtn"])
[data-testid="stFileUploaderDropzoneInstructions"] {
    display: none !important;
}
/* Hide the small Add button that appears after upload */
[data-testid="stSidebar"] [data-testid="stFileUploader"] button:not([data-testid="stFileUploaderDeleteBtn"]) {
    display: none !important;
}

/* Checkbox label text — transparent background, light text */
[data-testid="stSidebar"] .stCheckbox label {
    background-color: transparent !important; background: transparent !important;
}
[data-testid="stSidebar"] .stCheckbox label:hover {
    background-color: transparent !important; background: transparent !important;
}
[data-testid="stSidebar"] .stCheckbox label > div,
[data-testid="stSidebar"] .stCheckbox label p {
    color: #e8e8e8 !important; font-family: 'Nunito', sans-serif !important;
    font-weight: 700 !important; background-color: transparent !important;
}
[data-testid="stSidebar"] .stCheckbox [data-baseweb="checkbox"] [aria-checked="true"] > div:first-child {
    background-color: #E8461E !important; border-color: #E8461E !important;
}
[data-testid="stSidebar"] .stCheckbox [data-baseweb="checkbox"] > div:first-child {
    border-color: #E8461E !important;
}

.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
    background-color: #E8461E !important; border-color: #E8461E !important;
    color: white !important; font-family: 'Nunito', sans-serif !important;
    font-weight: 800 !important; border-radius: 3px !important;
}
.stButton > button, .stDownloadButton > button {
    border-color: #E8461E !important; color: #E8461E !important;
    font-family: 'Nunito', sans-serif !important; font-weight: 800 !important;
    border-radius: 3px !important;
}
.stProgress > div > div > div > div { background-color: #E8461E !important; }

.stRadio [role="radiogroup"] label {
    font-family: 'Nunito', sans-serif !important; font-weight: 700 !important;
    background-color: transparent !important; border-color: transparent !important;
}
:root { --primary-color: #E8461E !important; }
[data-baseweb="radio"] [role="radio"][aria-checked="true"] div {
    background-color: #E8461E !important; border-color: #E8461E !important;
}
[data-baseweb="radio"] [role="radio"] div { border-color: #E8461E !important; }
a { color: #E8461E !important; }

/* ── Cards ── */
.tc-card {
    background: #ffffff; border: 1px solid #e5e5e5;
    border-top: 4px solid #E8461E; border-radius: 4px;
    padding: 24px 26px; margin-bottom: 0;
    box-shadow: 0 2px 10px rgba(0,0,0,0.05);
}
.tc-card-title {
    font-family: 'Nunito', sans-serif; font-size: 16px;
    font-weight: 900; color: #1a1a1a; margin-bottom: 12px; letter-spacing: -0.1px;
}
.tc-list { list-style: none; padding: 0; margin: 0; }
.tc-list li {
    font-family: 'Nunito', sans-serif; font-size: 14px; font-weight: 600;
    color: #333; padding: 4px 0; display: flex; gap: 10px;
    align-items: flex-start; line-height: 1.5;
}
.tc-list li::before { content: "▸"; color: #E8461E; font-weight: 900; font-size: 14px; margin-top: 2px; }
.tc-legend { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 8px; }
.tc-pill {
    display: inline-flex; align-items: center; gap: 7px; padding: 5px 12px;
    border-radius: 3px; font-family: 'Nunito', sans-serif; font-size: 12px;
    font-weight: 700; background: #f5f5f5; border: 1px solid #ddd; color: #333;
}
.tc-swatch { width: 12px; height: 12px; border-radius: 2px; flex-shrink: 0; }
</style>
""", unsafe_allow_html=True)

# ── Navbar ─────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="tc-topbar"><span>OTDR QC Tools</span><span>Help</span></div>
<div class="tc-navbar">
    <div class="tc-logo-icon">↗</div>
    <a href="/" target="_self" class="tc-logo-name">Splice Report + Fiber Viewer</a>
    <div class="tc-navbar-spacer"></div>
    <div class="tc-contact-btn">OTDR QC &nbsp;▸</div>
</div>
""", unsafe_allow_html=True)


# ── Password ───────────────────────────────────────────────────────────────────

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True
    try:
        correct = st.secrets["passwords"]["app_password"]
    except (KeyError, FileNotFoundError):
        return True
    pwd = st.text_input("Enter password", type="password", key="pwd_input")
    if pwd:
        if pwd == correct:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password")
    return False

if not check_password():
    st.stop()


# ── Session state ──────────────────────────────────────────────────────────────

for key in ["viewer_html", "done"]:
    if key not in st.session_state:
        st.session_state[key] = None
if "upload_key" not in st.session_state:
    st.session_state.upload_key = 0


# ── Sidebar ────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## Upload SOR Files")

    input_method = st.radio(
        "Input method", ["Upload ZIP", "Browse files"],
        index=0, horizontal=True,
    )

    uploaded_a = uploaded_b = zip_a = zip_b = None

    if input_method == "Upload ZIP":
        zip_a = st.file_uploader("A-direction ZIP", type=["zip"],
                                 accept_multiple_files=False,
                                 key=f"zip_a_{st.session_state.upload_key}")
        if zip_a:
            st.caption(f"A: {zip_a.name} ({zip_a.size/1024:.0f} KB)")
        zip_b = st.file_uploader("B-direction ZIP (optional)", type=["zip"],
                                 accept_multiple_files=False,
                                 key=f"zip_b_{st.session_state.upload_key}")
        if zip_b:
            st.caption(f"B: {zip_b.name} ({zip_b.size/1024:.0f} KB)")
    else:
        uploaded_a = st.file_uploader("A-direction SOR files", type=["sor"],
                                      accept_multiple_files=True,
                                      key=f"upload_a_{st.session_state.upload_key}")
        uploaded_b = st.file_uploader("B-direction SOR files (optional)", type=["sor"],
                                      accept_multiple_files=True,
                                      key=f"upload_b_{st.session_state.upload_key}")

    if st.button("Clear All", use_container_width=True):
        old_key = st.session_state.upload_key
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.session_state.upload_key = old_key + 1
        st.rerun()

    st.divider()
    st.markdown("## Settings")
    col_sa, col_sb = st.columns(2)
    with col_sa:
        site_a_input = st.text_input("Site A (launch end)", value="", placeholder="e.g. TUL")
    with col_sb:
        site_b_input = st.text_input("Site B (far end)", value="", placeholder="e.g. BAR")
    span_override = st.number_input("Span override (km, 0 = auto-detect)",
                                    value=0.0, format="%.2f", step=0.5,
                                    help="Enter the known cable span in km. Use 0 to let the app estimate it from the SOR files.")
    threshold   = st.number_input("Bidirectional threshold (dB, 0.15=auto)",
                                  value=REBURN_THRESHOLD, format="%.3f", step=0.01)
    ribbon_size = RIBBON_SIZE

    st.markdown("**Include in Report**")
    col_chk1, col_chk2 = st.columns(2)
    with col_chk1:
        inc_reburn = st.checkbox("A+B Reburn", value=True, key="inc_reburn")
        inc_break  = st.checkbox("Break",      value=True, key="inc_break")
        inc_broke  = st.checkbox("Broke",      value=True, key="inc_broke")
    with col_chk2:
        inc_bfill  = st.checkbox("B-fill",     value=True, key="inc_bfill")
        inc_a_only = st.checkbox("A-only",     value=True, key="inc_a_only")
        inc_b_only = st.checkbox("B-only",     value=True, key="inc_b_only")

    has_a = bool(uploaded_a) or bool(zip_a)
    run_button = st.button("Generate Viewer", type="primary",
                           use_container_width=True, disabled=not has_a)


# ── Helpers ────────────────────────────────────────────────────────────────────

def stage_files(uploaded, prefix="sor_"):
    tmpdir = tempfile.mkdtemp(prefix=prefix)
    for uf in uploaded:
        with open(os.path.join(tmpdir, uf.name), 'wb') as f:
            f.write(uf.getbuffer())
    return tmpdir


def stage_zip(uploaded_zip, prefix="sor_zip_"):
    tmpdir = tempfile.mkdtemp(prefix=prefix)
    with zipfile.ZipFile(io.BytesIO(uploaded_zip.getbuffer()), 'r') as zf:
        for name in zf.namelist():
            if name.lower().endswith('.sor') and not name.startswith('__MACOSX'):
                basename = os.path.basename(name)
                if basename:
                    with zf.open(name) as src, open(os.path.join(tmpdir, basename), 'wb') as dst:
                        dst.write(src.read())
    return tmpdir


def compute_fiber_topo(trace, dist, dx_km, noise_km, step_km=0.5, window_km=1.0):
    if trace is None or len(trace) < 10:
        return []
    half_win = int((window_km / 2) / dx_km)
    step_samples = max(1, int(step_km / dx_km))
    i_noise = int((noise_km - dist[0]) / dx_km) if dx_km > 0 else len(trace)
    i_noise = min(i_noise, len(trace))
    i_start = int(0.5 / dx_km) + half_win
    profile = []
    i = i_start
    while i < i_noise - half_win:
        lo, hi = i - half_win, i + half_win
        if hi >= len(trace) or lo < 0:
            i += step_samples
            continue
        try:
            c = np.polyfit(dist[lo:hi], trace[lo:hi], 1)
            profile.append([round(float(dist[i]), 2), round(abs(float(c[0])), 5)])
        except Exception:
            pass
        i += step_samples
    return profile


def fiber_dist_array(r):
    """Return (trace, dist_array, dx_km, noise_km) using IOR-calibrated dx for native SOR trace.

    r['trace'] is the native SOR DataPts trace (NOT EXFO RawSamples).
    Its correct dx uses acq_range and the file IOR, derived from event calibration:
      dx_km = 2 * acq_range * event_dist_km / (event_time_of_travel * full_points)
    This formula cancels IOR so the trace distances stay self-consistent with
    the event dist_km values already computed from the same IOR in the SOR reader.
    """
    trace = r.get('trace')
    if trace is None:
        return None, None, None, None
    full_points = r.get('full_points', len(trace))
    if full_points <= 0:
        return None, None, None, None

    acq_range = r.get('acq_range', 0)
    events = r.get('events', [])
    dx_km = None

    # Preferred: calibrate dx from a mid-span non-end event (most accurate)
    if acq_range > 0:
        cal = [e for e in events
               if e.get('time_of_travel', 0) > 0 and e.get('dist_km', 0) > 1.0
               and not e.get('is_end')]
        if not cal:
            cal = [e for e in events
                   if e.get('time_of_travel', 0) > 0 and e.get('dist_km', 0) > 0
                   and not e.get('is_end')]
        if cal:
            e = cal[len(cal) // 2]
            dx_km = (2.0 * acq_range * e['dist_km']
                     / (e['time_of_travel'] * full_points))

    # Fallback: exfo_sampling_period (less accurate — different trace sampling)
    if not dx_km or dx_km <= 0:
        _C = 2.998e8
        _IOR = 1.4682
        sp = r.get('exfo_sampling_period', 0)
        if sp <= 0:
            return None, None, None, None
        dx_km = sp * _C / (2 * _IOR) / 1000.0

    start_km = r.get('start_index', 0) * dx_km
    n = len(trace)
    dist = np.linspace(start_km, start_km + n * dx_km, n)
    noise_km = start_km + n * dx_km
    for e in events:
        if e.get('is_end'):
            noise_km = min(float(e['dist_km']), noise_km)
            break
    return trace, dist, dx_km, noise_km


def build_viewer_html(combined_data):
    """Inject combined_data into the viewer HTML as an inline JS variable."""
    html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'combined_app.html')
    with open(html_path, 'r', encoding='utf-8') as f:
        html = f.read()

    data_js = json.dumps(combined_data, separators=(',', ':'))
    # Escape </script> so the JSON blob can't close the script tag early
    data_js = data_js.replace('</', '<\\/')

    # Inject data and auto-start — replace the loadDataAndStart function body
    # to use inline data first, and remove the setup screen auto-show
    inject = (
        '\n<script>\nwindow._SPLICE_DATA = ' + data_js + ';\n</script>\n'
    )
    # Insert injection before closing </head>
    html = html.replace('</head>', inject + '</head>', 1)

    # Replace the loadDataAndStart data-loading block to use inline data
    html = html.replace(
        """  try {
    // Try pywebview API first, fall back to static file
    if (window.pywebview?.api?.get_data) {
      data = await window.pywebview.api.get_data();
    }
    if (!data) {
      data = await (await fetch('combined_data.json?v=' + Date.now())).json();
    }
  } catch (e) {
    try { data = await (await fetch('combined_data.json?v=' + Date.now())).json(); }
    catch (e2) { alert('Could not load data: ' + e2.message); return; }
  }""",
        """  data = window._SPLICE_DATA;"""
    )

    # Auto-start: hide setup screen and call loadDataAndStart on page load
    html = html.replace(
        """if (!window.pywebview) {
  // Running via http.server — try loading existing data directly
  setTimeout(() => {
    fetch('combined_data.json?v=' + Date.now()).then(r => r.json()).then(d => {
      if (d) { document.getElementById('setup-screen').style.display = 'none'; data = d; loadDataAndStart(); }
    }).catch(() => {});
  }, 500);
}""",
        """// Auto-start with inline data
document.getElementById('setup-screen').style.display = 'none';
window.addEventListener('load', () => { loadDataAndStart(); });"""
    )

    # Fix body height for iframe embedding
    html = html.replace(
        'body { background: #fff; color: #333; font-family: \'Segoe UI\', system-ui, sans-serif; overflow: hidden; display: flex; }',
        'body { background: #fff; color: #333; font-family: \'Segoe UI\', system-ui, sans-serif; overflow: hidden; display: flex; width: 100vw; height: 100vh; }'
    )

    return html


# ── Run ────────────────────────────────────────────────────────────────────────

if run_button and has_a:
  try:
    if zip_a:
        bar = st.progress(0.0, text="Extracting ZIPs...")
        dir_a = stage_zip(zip_a, "splice_a_")
        bar.progress(0.1, text="Extracting B-direction ZIP...")
        dir_b = stage_zip(zip_b, "splice_b_") if zip_b else None
    else:
        bar = st.progress(0.0, text="Staging files...")
        dir_a = stage_files(uploaded_a, "splice_a_")
        dir_b = stage_files(uploaded_b, "splice_b_") if uploaded_b else None

    bar.progress(0.15, text="Loading SOR files...")
    fibers_a, fibers_b = load_all(dir_a, dir_b)
    n_fibers = max(fibers_a.keys()) if fibers_a else 0

    bar.progress(0.25, text="Discovering splice positions...")
    splices_raw = discover_splices(fibers_a)
    splices = [{**sp, 'splice_num': i + 1, 'is_bend': False}
               for i, sp in enumerate(splices_raw)]

    # ── Span detection ────────────────────────────────────────────────────────
    # The splice positions from discover_splices are derived from OTDR event tables
    # and are always accurate. The last splice km is a guaranteed lower bound on span.
    # Start there and only upgrade — never downgrade.

    if splices:
        last_splice_km = max(sp['position_km'] for sp in splices)
        span_km = round(last_splice_km + 1.0, 2)   # at minimum, span = last splice + 1 km
        span_debug = f"span={span_km} km (last splice {last_splice_km:.3f}+1)"
    else:
        span_km = 0
        span_debug = "span=0 (no splices found)"

    if span_override and span_override > 1.0:
        span_km = round(float(span_override), 2)
        span_debug = f"span={span_km} km (manual override)"

    # Upgrade only: check 1E end events from both directions
    all_1e_a = [e['dist_km'] for r in fibers_a.values()
                for e in r['events'] if e.get('is_end') and 1.0 < e['dist_km'] < 300]
    all_1e_b = [e['dist_km'] for r in (fibers_b or {}).values()
                for e in r.get('events', []) if e.get('is_end') and 1.0 < e['dist_km'] < 300]
    for v in ([max(all_1e_a)] if all_1e_a else []) + ([max(all_1e_b)] if all_1e_b else []):
        if v > span_km:
            span_km = round(v, 2)
            span_debug += f" → upgraded by 1E event to {span_km}"

    bar.progress(0.40, text=f"Pass 1: {n_fibers} fibers x {len(splices)} splices | {span_debug}")
    results = analyze_all(fibers_a, fibers_b, splices, threshold)

    bar.progress(0.60, text="Pass 2: scanning B-direction for missed events...")
    b_results = scan_b_events(fibers_a, fibers_b, splices, threshold, results, span_km)
    all_results = {**results, **b_results}

    # ── Apply event-type filters from sidebar ──────────────────────────────────
    def _included(r):
        if r.get('is_break')  and not st.session_state.get('inc_break',  True): return False
        if r.get('is_broke')  and not st.session_state.get('inc_broke',  True): return False
        if r.get('is_bfill')  and not st.session_state.get('inc_bfill',  True): return False
        if r.get('is_a_only') and not st.session_state.get('inc_a_only', True): return False
        if r.get('is_b_only') and not st.session_state.get('inc_b_only', True): return False
        is_reburn = (r.get('event_source') == 'bidir'
                     and not r.get('is_break') and not r.get('is_broke')
                     and not r.get('is_bfill') and not r.get('is_a_only')
                     and not r.get('is_b_only'))
        if is_reburn and not st.session_state.get('inc_reburn', True): return False
        return True
    all_results = {k: v for k, v in all_results.items() if _included(v)}

    # Build breaks dict
    breaks = {}
    for (fnum, si), r in all_results.items():
        if r.get('is_break') and fnum not in breaks:
            loss_val = r.get('bidir_loss') or r.get('a_loss') or 0
            breaks[fnum] = {'km': round(r['bidir_dist'], 3), 'type': 'BREAK',
                            'step_loss': round(abs(loss_val), 2) if loss_val else None,
                            'offset_ft': 0}
        elif r.get('is_broke') and fnum not in breaks:
            breaks[fnum] = {'km': round(r['bidir_dist'], 3), 'type': 'BROKE',
                            'step_loss': None, 'offset_ft': 0}

    bar.progress(0.70, text="Building splice table...")
    n_ribbons = (n_fibers + RIBBON_SIZE - 1) // RIBBON_SIZE
    n_tubes = (n_ribbons + 1) // 2
    letters = 'ABCDEFGHIJKLMNOPQRSTUVWX'[:n_tubes]
    ribbon_data = []

    for ri in range(n_ribbons):
        f_start = ri * RIBBON_SIZE + 1
        f_end = min(f_start + RIBBON_SIZE - 1, n_fibers)
        li = ri // 2
        tube = f"{letters[li]}{(ri % 2) + 1}" if li < len(letters) else f"?{ri}"
        splice_cells = []
        for si, sp in enumerate(splices):
            entries = []
            for fnum in range(f_start, f_end + 1):
                key = (fnum, si)
                if key not in all_results:
                    continue
                r = all_results[key]
                label = r.get('label', f"{fnum}")
                parts = label.split(' ', 1)
                label_short = parts[1] if len(parts) > 1 else label
                if r.get('is_break'):
                    etype = 'break'
                elif r.get('is_broke'):
                    etype = 'broke'
                elif r.get('is_bfill'):
                    etype = 'bfill'
                elif r.get('is_a_only'):
                    # est_bidir >= threshold → pink like a normal reburn (matches Excel)
                    etype = 'flagged' if r.get('est_bidir_flagged') else 'a_only'
                elif r.get('is_b_only'):
                    # est_bidir >= threshold → pink like a normal reburn (matches Excel)
                    etype = 'flagged' if r.get('est_bidir_flagged') else 'b_only'
                else:
                    etype = 'flagged'
                loss = r.get('bidir_loss') or r.get('a_loss') or r.get('b_loss') or 0
                entries.append({'fnum': fnum, 'text': f"F{fnum} {label_short}",
                                'type': etype, 'loss': round(abs(loss), 3) if loss else 0})
            splice_cells.append({'splice_idx': si, 'entries': entries})
        ribbon_data.append({'ribbon': ri + 1, 'tube': tube,
                            'fibers': f'{f_start}-{f_end}',
                            'f_start': f_start, 'f_end': f_end, 'cells': splice_cells})

    bar.progress(0.80, text="Computing fiber profiles and traces...")
    all_slopes = []
    fiber_profiles = {}
    fiber_traces = {}
    target_pts = 1500

    for fnum in sorted(fibers_a.keys()):
        ra = fibers_a[fnum]
        trace, dist, dx_km, noise_km = fiber_dist_array(ra)
        if trace is None:
            continue
        prof = compute_fiber_topo(trace, dist, dx_km, noise_km)
        if prof:
            fiber_profiles[str(fnum)] = prof
            all_slopes.extend(s for _, s in prof)
        brk = breaks.get(fnum)
        if brk:
            end_idx = max(0, int((brk['km'] - dist[0]) / dx_km))
        else:
            end_idx = int((noise_km - dist[0]) / dx_km)
        end_idx = min(end_idx, len(trace))
        step = max(1, end_idx // target_pts)
        trace_pts = []
        for i in range(0, end_idx - step, step):
            km = round(float(np.mean(dist[i:i + step])), 3)
            db = round(float(np.mean(trace[i:i + step])), 3)
            trace_pts.append([km, db])
        if brk and end_idx > 0 and end_idx < len(trace):
            live_db = round(float(trace[min(end_idx, len(trace) - 1)]), 3)
            noise_db = round(float(np.min(trace[max(0, end_idx - 10):min(len(trace), end_idx + 200)])), 3)
            trace_pts.append([round(brk['km'], 3), live_db])
            trace_pts.append([round(brk['km'] + 0.01, 3), noise_db])
            trace_pts.append([round(span_km, 3), noise_db])
        if trace_pts:
            fiber_traces[str(fnum)] = trace_pts

    baseline = float(np.median(all_slopes)) if all_slopes else 0.19

    # Build B-direction traces in A-frame (for A+B combined view)
    # Each B trace is converted: km_a = span_km - km_b, then sorted ascending.
    # The result shows the B-end view mirrored onto the A-frame axis so both
    # traces can be overlaid — A slopes down left-to-right, B slopes down
    # right-to-left, together covering the full span even across breaks.
    fiber_traces_b = {}
    if fibers_b:
        for fnum in sorted(fibers_b.keys()):
            rb = fibers_b[fnum]
            trace_b, dist_b, dx_km_b, noise_km_b = fiber_dist_array(rb)
            if trace_b is None:
                continue

            # Clip B trace to its SIGNAL region only — stop at the first
            # significant reflective event (break from B's side) so we never
            # include B's noise floor in the trace.  For a broken fiber this
            # is the cut reflection; for a healthy fiber it's the far connector.
            b_signal_end_km = noise_km_b   # default: end event or full trace
            for e in rb.get('events', []):
                if e.get('is_reflective') and e.get('dist_km', 0) > 1.0:
                    b_signal_end_km = float(e['dist_km'])
                    break   # first reflective past launch = break or far end
            # Take the min so we never go past the true end
            b_signal_end_km = min(b_signal_end_km, noise_km_b)

            end_idx_b = min(int((b_signal_end_km - dist_b[0]) / dx_km_b), len(trace_b))
            end_idx_b = max(end_idx_b, 10)   # at least some points
            step_b = max(1, end_idx_b // target_pts)
            pts_b = []
            for i in range(0, end_idx_b - step_b, step_b):
                km_b = float(np.mean(dist_b[i:i + step_b]))
                db   = float(np.mean(trace_b[i:i + step_b]))
                pts_b.append([round(span_km - km_b, 3), round(db, 3)])
            pts_b.sort(key=lambda p: p[0])   # sort ascending by A-frame km
            if pts_b:
                fiber_traces_b[str(fnum)] = pts_b

    # Site names: prefer sidebar inputs; fall back to auto-detect from folder name
    _sa = site_a_input.strip().upper()
    _sb = site_b_input.strip().upper()
    if _sa and _sb:
        site_a, site_b = _sa, _sb
    else:
        folder_name = os.path.basename(os.path.normpath(dir_a)).upper()
        alpha = ''.join(c for c in folder_name if c.isalpha())
        if len(alpha) == 6:
            site_a, site_b = alpha[:3], alpha[3:]
        elif len(alpha) in (7, 8):
            mid = len(alpha) // 2
            site_a, site_b = alpha[:mid], alpha[mid:]
        else:
            site_a = _sa if _sa else 'A'
            site_b = _sb if _sb else 'B'

    # ── Per-fiber event lists for dark-zone extrapolation ────────────────────
    # For each fiber: all non-trivial events from A direction (at their A-frame km)
    # plus all non-trivial events from B direction (converted to A-frame).
    # The JS composite builder uses these to draw the Rayleigh slope + splice steps
    # across any gap between A's trace end and B's trace start.
    bar.progress(0.88, text="Building fiber event lists...")
    fiber_events = {}
    for fnum in sorted(fibers_a.keys()):
        ra = fibers_a[fnum]
        rb = fibers_b.get(fnum) if fibers_b else None
        evts = []

        # A-direction events (skip launch at 0 km and end events)
        for e in ra.get('events', []):
            if e.get('is_end') or e.get('dist_km', 0) < 0.5:
                continue
            loss = abs(e.get('splice_loss', 0))
            if loss > 0.005:
                evts.append([round(float(e['dist_km']), 3), round(loss, 4)])

        # B-direction events converted to A-frame
        if rb:
            b_end_evt = next((e for e in rb.get('events', []) if e.get('is_end')), None)
            b_span = float(b_end_evt['dist_km']) if b_end_evt else None
            if b_span is None:
                # fallback: last reflective event
                b_refs = [e for e in rb.get('events', [])
                          if e.get('is_reflective') and not e.get('is_end') and e.get('dist_km', 0) > 1.0]
                b_span = float(b_refs[-1]['dist_km']) if b_refs else None
            if b_span:
                for e in rb.get('events', []):
                    if e.get('is_end') or e.get('dist_km', 0) < 0.5:
                        continue
                    km_a = round(span_km - e['dist_km'], 3)
                    if km_a < 0.5 or km_a > span_km:
                        continue
                    loss = abs(e.get('splice_loss', 0))
                    if loss > 0.005:
                        evts.append([km_a, round(loss, 4)])

        # Sort by km; deduplicate: within 0.3 km keep highest loss
        evts.sort(key=lambda x: x[0])
        deduped = []
        for ev in evts:
            if deduped and abs(ev[0] - deduped[-1][0]) < 0.3:
                if ev[1] > deduped[-1][1]:
                    deduped[-1] = ev
            else:
                deduped.append(ev)
        if deduped:
            fiber_events[str(fnum)] = deduped

    combined_data = {
        'meta': {
            'site_a': site_a, 'site_b': site_b,
            'span_km': span_km, 'n_fibers': n_fibers,
            'n_ribbons': n_ribbons, 'ribbon_size': RIBBON_SIZE,
            'baseline': round(baseline, 5), 'threshold': threshold,
        },
        'splices': [{'km': sp['position_km'], 'is_bend': sp['is_bend'],
                     'splice_num': sp['splice_num'], 'count': sp['count']}
                    for sp in splices],
        'ribbons': ribbon_data,
        'breaks': {str(f): b for f, b in breaks.items()},
        'fiber_profiles': fiber_profiles,
        'fiber_traces': fiber_traces,
        'fiber_traces_b': fiber_traces_b,
        'fiber_events': fiber_events,
        'raw_events_a': {
            str(fnum): [
                {
                    'km':         round(float(e['dist_km']), 4),
                    'splice_loss': round(float(e.get('splice_loss', 0)), 4),
                    'reflection':  round(float(e.get('reflection', 0)), 4),
                    'slope':       round(float(e.get('slope', 0)), 4),
                    'type':        e.get('type', ''),
                    'is_reflective': bool(e.get('is_reflective', False)),
                    'is_end':      bool(e.get('is_end', False)),
                }
                for e in fibers_a[fnum].get('events', [])
                if e.get('dist_km', 0) > 0.1
            ]
            for fnum in sorted(fibers_a.keys())
        },
    }

    bar.progress(0.95, text="Building viewer...")
    st.session_state.viewer_html = build_viewer_html(combined_data)
    st.session_state.done = True
    bar.progress(1.0, text="Viewer ready — loading interface...")
    # Do NOT call bar.empty() here — keep the bar visible.
    # The rerun triggered by session_state change will show the viewer.
    # The bar is replaced naturally when the display section renders.
  except Exception as _e:
    import traceback
    st.error(f"Error generating viewer: {_e}")
    st.code(traceback.format_exc())


# ── Display ────────────────────────────────────────────────────────────────────

if st.session_state.get("done") and st.session_state.viewer_html:
    import streamlit.components.v1 as components

    # Hide sidebar and remove its margin when in viewer mode
    st.markdown("""
    <style>
    [data-testid="stSidebar"]          { display: none !important; }
    [data-testid="stSidebarCollapsed"] { display: none !important; }
    .main .block-container {
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        padding-top: 0.25rem !important;
        max-width: 100% !important;
    }
    section[data-testid="stMain"] > div:first-child { margin-left: 0 !important; }
    </style>
    """, unsafe_allow_html=True)

    # Show a loading bar while the iframe paints
    load_bar = st.progress(1.0, text="Viewer loading — please wait...")

    # Return button left-aligned, then full-width orange line
    col_btn, _ = st.columns([1, 5])
    with col_btn:
        if st.button("← Return to Main Page", type="primary", use_container_width=True):
            st.session_state.done = None
            st.session_state.viewer_html = None
            st.rerun()
    st.markdown("""
    <div style="background:#E8461E; height:3px; width:100%; margin:4px 0 6px 0;"></div>
    """, unsafe_allow_html=True)

    components.html(st.session_state.viewer_html, height=860, scrolling=False)
    load_bar.empty()   # clear bar once iframe content has been sent to browser

else:
    st.markdown("""
    <div style="background:linear-gradient(105deg,#E8461E 55%,#c23610 100%);
                padding:42px 36px 38px 36px; margin-bottom:28px;">
        <h1 style="font-family:'Nunito',sans-serif;font-size:32px;font-weight:900;
                   color:#fff;margin:0 0 10px 0;line-height:1.15;">
            Splice Report<br>+ Fiber Viewer
        </h1>
        <p style="font-family:'Nunito',sans-serif;font-size:15px;
                  color:rgba(255,255,255,0.88);margin:0;font-weight:600;">
            Upload A and B direction SOR files to generate the interactive viewer
        </p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style="display:flex; gap:18px; align-items:stretch; margin-bottom:18px;">
        <div class="tc-card" style="flex:1;">
            <div class="tc-card-title">Pass 1 — Splice Position Analysis</div>
            <ul class="tc-list">
                <li>Discovers splice closure positions where 20+ fibers share an event</li>
                <li>Finds A+B bidirectional events and flags if loss &ge; threshold</li>
                <li>Detects broke fibers (mid-span trace termination)</li>
                <li>Fills B-direction data past breaks where A is blind</li>
                <li>Flags A-only events with estimated bidir = A / 2</li>
            </ul>
        </div>
        <div class="tc-card" style="flex:1;">
            <div class="tc-card-title">Pass 2 — B-Direction Event Scan</div>
            <ul class="tc-list">
                <li>Scans every B-direction event above threshold not caught in Pass 1</li>
                <li>Converts B-frame positions to A-frame coordinates</li>
                <li>Matches to nearest splice position within 1.5 km</li>
                <li>If A event also found: computes true bidirectional average</li>
                <li>If no A event: flags as B-only with estimated bidir = B / 2</li>
                <li>Catches events regardless of which direction saw it first</li>
            </ul>
        </div>
    </div>
    <div style="display:flex; gap:18px; align-items:stretch; margin-bottom:18px;">
        <div class="tc-card" style="flex:1;">
            <div class="tc-card-title">How To Use</div>
            <ul class="tc-list">
                <li>Upload A-direction SOR files (required) and B-direction (optional) as a ZIP or individual files</li>
                <li>Adjust the reburn threshold if needed — default is 0.150 dB</li>
                <li>Use the <strong>Include in Report</strong> checkboxes to filter which event types appear in the viewer</li>
                <li>Click <strong>Generate Viewer</strong> to launch the interactive splice table and 2D OTDR trace</li>
                <li>Use <strong>Pop Out Right Screen</strong> to open the trace on a second monitor, then <strong>Collapse Right Screen</strong> to go full-width on the splice table, or <strong>Expand Right Screen</strong> to restore the split view</li>
            </ul>
        </div>
        <div class="tc-card" style="flex:1;">
            <div class="tc-card-title">Viewer Color Key</div>
            <div style="display:flex;flex-direction:column;gap:7px;margin-top:4px;">
                <div style="display:flex;align-items:center;">
                    <div style="background:#FFC7CE;color:#1a1a1a;font-family:'Nunito',sans-serif;font-size:11px;font-weight:700;padding:2px 7px;border:1px solid rgba(0,0,0,0.14);white-space:nowrap;width:148px;flex-shrink:0;">325 .172</div>
                    <div style="font-family:'Nunito',sans-serif;font-size:12px;font-weight:600;color:#444;padding-left:12px;">A+B bidirectional reburn</div>
                </div>
                <div style="display:flex;align-items:center;">
                    <div style="background:#FF4444;color:#ffffff;font-family:'Nunito',sans-serif;font-size:11px;font-weight:700;padding:2px 7px;border:1px solid rgba(0,0,0,0.14);white-space:nowrap;width:148px;flex-shrink:0;">107 BREAK .210</div>
                    <div style="font-family:'Nunito',sans-serif;font-size:12px;font-weight:600;color:#444;padding-left:12px;">1F reflective break</div>
                </div>
                <div style="display:flex;align-items:center;">
                    <div style="background:#FF8800;color:#ffffff;font-family:'Nunito',sans-serif;font-size:11px;font-weight:700;padding:2px 7px;border:1px solid rgba(0,0,0,0.14);white-space:nowrap;width:148px;flex-shrink:0;">107 broke</div>
                    <div style="font-family:'Nunito',sans-serif;font-size:12px;font-weight:600;color:#444;padding-left:12px;">trace terminates mid-span</div>
                </div>
                <div style="display:flex;align-items:center;">
                    <div style="background:#BDD7EE;color:#1F4E79;font-family:'Nunito',sans-serif;font-size:11px;font-weight:700;padding:2px 7px;border:1px solid rgba(0,0,0,0.14);white-space:nowrap;width:148px;flex-shrink:0;">214 .188 (B-fill)</div>
                    <div style="font-family:'Nunito',sans-serif;font-size:12px;font-weight:600;color:#444;padding-left:12px;">B-direction past a break</div>
                </div>
                <div style="display:flex;align-items:center;">
                    <div style="background:#FFF2CC;color:#7F6000;font-family:'Nunito',sans-serif;font-size:11px;font-weight:700;padding:2px 7px;border:1px solid rgba(0,0,0,0.14);white-space:nowrap;width:148px;flex-shrink:0;">83 .151(A) ~.075bd</div>
                    <div style="font-family:'Nunito',sans-serif;font-size:12px;font-weight:600;color:#444;padding-left:12px;">A-only, est. bidir below threshold</div>
                </div>
                <div style="display:flex;align-items:center;">
                    <div style="background:#FFD700;color:#4B3000;font-family:'Nunito',sans-serif;font-size:11px;font-weight:700;padding:2px 7px;border:1px solid rgba(0,0,0,0.14);white-space:nowrap;width:148px;flex-shrink:0;">122 .285(A) &#9888;.143bd</div>
                    <div style="font-family:'Nunito',sans-serif;font-size:12px;font-weight:600;color:#444;padding-left:12px;">A-only, est. bidir above threshold</div>
                </div>
                <div style="display:flex;align-items:center;">
                    <div style="background:#E8D5F5;color:#4B0082;font-family:'Nunito',sans-serif;font-size:11px;font-weight:700;padding:2px 7px;border:1px solid rgba(0,0,0,0.14);white-space:nowrap;width:148px;flex-shrink:0;">430 .161(B) ~.081bd</div>
                    <div style="font-family:'Nunito',sans-serif;font-size:12px;font-weight:600;color:#444;padding-left:12px;">B-only, est. bidir below threshold</div>
                </div>
                <div style="display:flex;align-items:center;">
                    <div style="background:#C084FC;color:#1A0033;font-family:'Nunito',sans-serif;font-size:11px;font-weight:700;padding:2px 7px;border:1px solid rgba(0,0,0,0.14);white-space:nowrap;width:148px;flex-shrink:0;">325 .340(B) &#9888;.170bd</div>
                    <div style="font-family:'Nunito',sans-serif;font-size:12px;font-weight:600;color:#444;padding-left:12px;">B-only, est. bidir above threshold</div>
                </div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style="display:flex;gap:18px;align-items:stretch;margin-bottom:18px;">
        <div style="flex:1;background:#fff;border:1px solid #e5e5e5;border-top:4px solid #E8461E;
                    border-radius:4px;padding:22px 24px;box-shadow:0 2px 10px rgba(0,0,0,0.05);">
            <div style="font-family:'Nunito',sans-serif;font-size:15px;font-weight:900;
                        color:#1a1a1a;margin-bottom:12px;">Left Panel - Splice Report</div>
            <ul style="list-style:none;padding:0;margin:0;">
                <li style="font-family:'Nunito',sans-serif;font-size:13px;font-weight:600;
                            color:#333;padding:4px 0;display:flex;gap:8px;">
                    <span style="color:#E8461E;font-weight:900;">▸</span>
                    Pass 1 + Pass 2 bidirectional splice analysis
                </li>
                <li style="font-family:'Nunito',sans-serif;font-size:13px;font-weight:600;
                            color:#333;padding:4px 0;display:flex;gap:8px;">
                    <span style="color:#E8461E;font-weight:900;">▸</span>
                    Color-coded cells: pink (A+B), red (break), orange (broke), blue (B-fill), yellow/gold (A-only), lavender/purple (B-only)
                </li>
                <li style="font-family:'Nunito',sans-serif;font-size:13px;font-weight:600;
                            color:#333;padding:4px 0;display:flex;gap:8px;">
                    <span style="color:#E8461E;font-weight:900;">▸</span>
                    Click any row or column to pan the right panels to that splice
                </li>
            </ul>
        </div>
        <div style="flex:1;background:#fff;border:1px solid #e5e5e5;border-top:4px solid #E8461E;
                    border-radius:4px;padding:22px 24px;box-shadow:0 2px 10px rgba(0,0,0,0.05);">
            <div style="font-family:'Nunito',sans-serif;font-size:15px;font-weight:900;
                        color:#1a1a1a;margin-bottom:12px;">Right Panel - Bidirectional OTDR Trace &amp; Event Table</div>
            <ul style="list-style:none;padding:0;margin:0;">
                <li style="font-family:'Nunito',sans-serif;font-size:13px;font-weight:600;
                            color:#333;padding:4px 0;display:flex;gap:8px;">
                    <span style="color:#E8461E;font-weight:900;">▸</span>
                    EXFO-style composite trace — A-direction raw signal stitched to B-direction (mirrored), with dark-zone fill across any break gap; A+B average shown where both are available
                </li>
                <li style="font-family:'Nunito',sans-serif;font-size:13px;font-weight:600;
                            color:#333;padding:4px 0;display:flex;gap:8px;">
                    <span style="color:#E8461E;font-weight:900;">▸</span>
                    View a single fiber or the full ribbon (up to 12 fibers overlaid) — click a ribbon cell or type a fiber number to switch; scroll to zoom, drag to pan
                </li>
                <li style="font-family:'Nunito',sans-serif;font-size:13px;font-weight:600;
                            color:#333;padding:4px 0;display:flex;gap:8px;">
                    <span style="color:#E8461E;font-weight:900;">▸</span>
                    Break and broke markers shown in red/orange with km label; trace automatically clips at the noise floor so only clean signal is displayed
                </li>
                <li style="font-family:'Nunito',sans-serif;font-size:13px;font-weight:600;
                            color:#333;padding:4px 0;display:flex;gap:8px;">
                    <span style="color:#E8461E;font-weight:900;">▸</span>
                    Event table below the trace lists every OTDR event with distance (km &amp; ft), event type, splice loss, reflection level, and attenuation slope
                </li>
                <li style="font-family:'Nunito',sans-serif;font-size:13px;font-weight:600;
                            color:#333;padding:4px 0;display:flex;gap:8px;">
                    <span style="color:#E8461E;font-weight:900;">▸</span>
                    Pop Out Right Screen opens a live interactive copy on a second monitor; Collapse / Expand Right Screen controls the split between the splice table and the trace panel
                </li>
            </ul>
        </div>
    </div>
    """, unsafe_allow_html=True)
