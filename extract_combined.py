#!/usr/bin/env python3
"""
extract_combined.py  (splice report edition)
============================================
Extract bidirectional splice report data + per-fiber attenuation profiles
for the combined splice-table + topo viewer.

Uses splicereportmatchexfo (Pass 1 + Pass 2) for the left panel.
Right-panel trace and topo data extracted from SOR RawSamples.

Usage:
    python3 extract_combined.py --dir-a <path> --dir-b <path> --site-a TUL --site-b BAR
"""

import argparse
import json
import os
import sys
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from splicereportmatchexfo import (
    load_all, discover_splices, analyze_all, scan_b_events,
    REBURN_THRESHOLD, RIBBON_SIZE,
)


def compute_fiber_topo(trace, dist, dx_km, noise_km, step_km=0.5, window_km=1.0):
    """Compute local attenuation slope profile for one fiber."""
    if trace is None or len(trace) < 10:
        return []
    half_win = int((window_km / 2) / dx_km)
    step_samples = max(1, int(step_km / dx_km))
    i_noise = int(min(noise_km, float(dist[-1])) / dist[0] * len(dist)) if dist[0] > 0 else int(noise_km / dx_km)
    # simpler: index from dist array
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
    """Return (trace, dist_array, dx_km, noise_km) from a parse_sor_full dict.

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
    # noise_km: use end event dist
    noise_km = start_km + n * dx_km
    for e in events:
        if e.get('is_end'):
            noise_km = min(float(e['dist_km']), noise_km)
            break
    return trace, dist, dx_km, noise_km


def main():
    ap = argparse.ArgumentParser(description='Extract combined splice + topo data')
    ap.add_argument('--dir-a', required=True)
    ap.add_argument('--dir-b', required=True)
    ap.add_argument('--site-a', default='A')
    ap.add_argument('--site-b', default='B')
    ap.add_argument('--threshold', type=float, default=REBURN_THRESHOLD)
    ap.add_argument('--ribbon-size', type=int, default=RIBBON_SIZE)
    ap.add_argument('--output', default='combined_data.json')
    args = ap.parse_args()

    print("Loading SOR files...")
    fibers_a, fibers_b = load_all(args.dir_a, args.dir_b)
    n_fibers = max(fibers_a.keys()) if fibers_a else 0

    print("Discovering splice positions...")
    splices_raw = discover_splices(fibers_a)
    # Add splice_num and is_bend (splicereportmatchexfo has no bend detection)
    splices = []
    for i, sp in enumerate(splices_raw):
        splices.append({**sp, 'splice_num': i + 1, 'is_bend': False})

    # Span
    all_ends = sorted([e['dist_km'] for r in fibers_a.values()
                       for e in r['events'] if e.get('is_end')])
    if all_ends:
        top_q = all_ends[int(len(all_ends) * 0.75):]
        span_km = round(float(np.median(top_q)), 2)
    else:
        span_km = 0
    print(f"  {len(fibers_a)} fibers, {span_km:.1f} km span, {len(splices)} splice positions")

    print("Running Pass 1 (splice position analysis)...")
    results = analyze_all(fibers_a, fibers_b, splices, args.threshold)

    print("Running Pass 2 (B-direction event scan)...")
    b_results = scan_b_events(fibers_a, fibers_b, splices, args.threshold, results, span_km)
    all_results = {**results, **b_results}

    # ── Build breaks dict from results ────────────────────────────────────────
    breaks = {}
    for (fnum, si), r in all_results.items():
        if r.get('is_break') and fnum not in breaks:
            loss_val = r.get('bidir_loss') or r.get('a_loss') or 0
            breaks[fnum] = {
                'km': round(r['bidir_dist'], 3),
                'type': 'BREAK',
                'step_loss': round(abs(loss_val), 2) if loss_val else None,
                'offset_ft': 0,
            }
        elif r.get('is_broke') and fnum not in breaks:
            breaks[fnum] = {
                'km': round(r['bidir_dist'], 3),
                'type': 'BROKE',
                'step_loss': None,
                'offset_ft': 0,
            }

    # ── Build ribbon data ─────────────────────────────────────────────────────
    print("Building splice table...")
    n_ribbons = (n_fibers + args.ribbon_size - 1) // args.ribbon_size
    n_tubes = (n_ribbons + 1) // 2
    letters = 'ABCDEFGHIJKLMNOPQRSTUVWX'[:n_tubes]
    ribbon_data = []

    for ri in range(n_ribbons):
        f_start = ri * args.ribbon_size + 1
        f_end = min(f_start + args.ribbon_size - 1, n_fibers)
        li = ri // 2
        side = (ri % 2) + 1
        tube = f"{letters[li]}{side}" if li < len(letters) else f"?{ri}"

        splice_cells = []
        for si, sp in enumerate(splices):
            entries = []
            for fnum in range(f_start, f_end + 1):
                key = (fnum, si)
                if key not in all_results:
                    continue
                r = all_results[key]
                label = r.get('label', f"{fnum}")
                # Strip leading fiber number from label for display
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
                entries.append({
                    'fnum': fnum,
                    'text': f"F{fnum} {label_short}",
                    'type': etype,
                    'loss': round(abs(loss), 3) if loss else 0,
                })
            splice_cells.append({'splice_idx': si, 'entries': entries})

        ribbon_data.append({
            'ribbon': ri + 1,
            'tube': tube,
            'fibers': f'{f_start}-{f_end}',
            'f_start': f_start,
            'f_end': f_end,
            'cells': splice_cells,
        })

    # ── Per-fiber attenuation profiles + raw dB traces ────────────────────────
    print("Computing per-fiber profiles + raw traces...")
    all_slopes = []
    fiber_profiles = {}
    fiber_traces = {}
    target_pts = 1500

    for fnum in sorted(fibers_a.keys()):
        ra = fibers_a[fnum]
        trace, dist, dx_km, noise_km = fiber_dist_array(ra)
        if trace is None:
            continue

        # Slope profile (for 3D topo view)
        prof = compute_fiber_topo(trace, dist, dx_km, noise_km)
        if prof:
            fiber_profiles[str(fnum)] = prof
            all_slopes.extend(s for _, s in prof)

        # Raw dB trace (for 2D trace view)
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
        # Add break cliff
        if brk and end_idx > 0 and end_idx < len(trace):
            live_db = round(float(trace[min(end_idx, len(trace) - 1)]), 3)
            noise_db = round(float(np.min(trace[max(0, end_idx - 10):min(len(trace), end_idx + 200)])), 3)
            trace_pts.append([round(brk['km'], 3), live_db])
            trace_pts.append([round(brk['km'] + 0.01, 3), noise_db])
            trace_pts.append([round(span_km, 3), noise_db])
        if trace_pts:
            fiber_traces[str(fnum)] = trace_pts

        if fnum % 100 == 0:
            print(f"    ...fiber {fnum}/{n_fibers}")

    baseline = float(np.median(all_slopes)) if all_slopes else 0.19

    # ── Output ────────────────────────────────────────────────────────────────
    output = {
        'meta': {
            'site_a': args.site_a,
            'site_b': args.site_b,
            'span_km': span_km,
            'n_fibers': n_fibers,
            'n_ribbons': n_ribbons,
            'ribbon_size': args.ribbon_size,
            'baseline': round(baseline, 5),
            'threshold': args.threshold,
        },
        'splices': [
            {'km': sp['position_km'], 'is_bend': sp['is_bend'],
             'splice_num': sp['splice_num'], 'count': sp['count']}
            for sp in splices
        ],
        'ribbons': ribbon_data,
        'breaks': {str(f): b for f, b in breaks.items()},
        'fiber_profiles': fiber_profiles,
        'fiber_traces': fiber_traces,
    }

    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), args.output)
    with open(out_path, 'w') as f:
        json.dump(output, f, separators=(',', ':'))
    size_mb = os.path.getsize(out_path) / 1e6
    print(f"  Saved: {out_path} ({size_mb:.1f} MB)")


if __name__ == '__main__':
    main()
