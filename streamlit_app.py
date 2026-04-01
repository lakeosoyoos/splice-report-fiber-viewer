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

/* Fix file uploader overlap at wide sidebar widths */
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
    flex-direction: column !important;
    align-items: stretch !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] {
    flex-direction: column !important;
    align-items: center !important;
    gap: 8px !important;
}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] > div {
    flex-direction: column !important;
    align-items: center !important;
    text-align: center !important;
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
    trace = r.get('trace')
    if trace is None:
        return None, None, None, None
    acq_range = r.get('acq_range', 0)
    full_points = r.get('full_points', len(trace))
    if acq_range <= 0 or full_points <= 0:
        return None, None, None, None
    dx_km = acq_range / full_points
    start_km = r.get('start_index', 0) * dx_km
    n = len(trace)
    dist = np.linspace(start_km, start_km + n * dx_km, n)
    noise_km = start_km + n * dx_km
    for e in r.get('events', []):
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

    # Inject data and auto-start — replace the loadDataAndStart function body
    # to use inline data first, and remove the setup screen auto-show
    inject = f"""
<script>
window._SPLICE_DATA = {data_js};
</script>
"""
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

    all_ends = sorted([e['dist_km'] for r in fibers_a.values()
                       for e in r['events'] if e.get('is_end')])
    if all_ends:
        top_q = all_ends[int(len(all_ends) * 0.75):]
        span_km = round(float(np.median(top_q)), 2)
    else:
        span_km = 0

    bar.progress(0.40, text=f"Pass 1: analyzing {n_fibers} fibers x {len(splices)} splices...")
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
                    etype = 'a_only_high' if r.get('est_bidir_flagged') else 'a_only'
                elif r.get('is_b_only'):
                    etype = 'b_only_high' if r.get('est_bidir_flagged') else 'b_only'
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

    # Auto-detect site names
    folder_name = os.path.basename(os.path.normpath(dir_a)).upper()
    alpha = ''.join(c for c in folder_name if c.isalpha())
    if len(alpha) == 6:
        site_a, site_b = alpha[:3], alpha[3:]
    elif len(alpha) in (7, 8):
        mid = len(alpha) // 2
        site_a, site_b = alpha[:mid], alpha[mid:]
    else:
        site_a, site_b = 'A', 'B'

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
    }

    bar.progress(0.95, text="Building viewer...")
    st.session_state.viewer_html = build_viewer_html(combined_data)
    st.session_state.done = True
    bar.progress(1.0, text="Done!")
    bar.empty()


# ── Display ────────────────────────────────────────────────────────────────────

if st.session_state.get("done") and st.session_state.viewer_html:
    import streamlit.components.v1 as components
    components.html(st.session_state.viewer_html, height=860, scrolling=False)

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
                        color:#1a1a1a;margin-bottom:12px;">Right Panels - Trace + Topo</div>
            <ul style="list-style:none;padding:0;margin:0;">
                <li style="font-family:'Nunito',sans-serif;font-size:13px;font-weight:600;
                            color:#333;padding:4px 0;display:flex;gap:8px;">
                    <span style="color:#E8461E;font-weight:900;">▸</span>
                    Top: 2D OTDR trace view - scroll to zoom, drag to pan, hover for event details
                </li>
                <li style="font-family:'Nunito',sans-serif;font-size:13px;font-weight:600;
                            color:#333;padding:4px 0;display:flex;gap:8px;">
                    <span style="color:#E8461E;font-weight:900;">▸</span>
                    Bottom: 3D topo view - attenuation profile surface with OrbitControls
                </li>
                <li style="font-family:'Nunito',sans-serif;font-size:13px;font-weight:600;
                            color:#333;padding:4px 0;display:flex;gap:8px;">
                    <span style="color:#E8461E;font-weight:900;">▸</span>
                    Enter a fiber number or click Full Ribbon to switch views
                </li>
            </ul>
        </div>
    </div>
    """, unsafe_allow_html=True)
