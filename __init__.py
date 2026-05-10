"""
ct_module/__init__.py
Cycle Time Report — Flask Blueprint
Integrated into ABSC Main as module 05.
All routes are prefixed with /ct/ so they never collide with existing ABSC routes.
"""

import io, os, sys, traceback, pickle, hashlib
from datetime import datetime
from flask import Blueprint, request, jsonify, make_response, session, redirect, url_for, render_template_string

# Make ct_module itself importable as a package path for its sibling modules.
# IMPORTANT: use append, NOT insert(0, ...) — inserting at position 0 causes
# ct_module/report_engine.py to shadow the root-level report_engine.py,
# breaking the main app.py import on Render.
_CT_DIR = os.path.dirname(os.path.abspath(__file__))
if _CT_DIR not in sys.path:
    sys.path.append(_CT_DIR)

ct = Blueprint('ct', __name__, url_prefix='/ct')

# ── File-based store (shared across all gunicorn workers via /tmp) ────────────
# This replaces the old in-memory _ct_store which broke under multi-worker
# Render deployments because each worker had its own isolated copy.
_STORE_DIR = '/tmp/ct_store'
os.makedirs(_STORE_DIR, exist_ok=True)

_EMPTY_STORE = {
    'df_ct': None, 'df_anomaly': None,
    'df_wb_total': None, 'df_wb_rejected': None,
    'failure_list': [], 'exceedances': 0,
    'applicable_hours': 0, 'net_weight': 0.0,
    'start_dt': None, 'end_dt': None, 'reason': '',
    'main_xlsx': None, 'summary_xlsx': None, 'ppt_bytes': None,
}

def _store_path(sid: str) -> str:
    """Return a safe /tmp path for this session's store."""
    safe = hashlib.sha256(sid.encode()).hexdigest()[:32]
    return os.path.join(_STORE_DIR, f'ct_{safe}.pkl')

def _load_store() -> dict:
    """Load this session's store from disk, or return a fresh empty store."""
    sid = session.get('username', '__anonymous__')
    path = _store_path(sid)
    if os.path.exists(path):
        try:
            with open(path, 'rb') as fh:
                return pickle.load(fh)
        except Exception:
            pass  # corrupted file — fall through to fresh store
    import copy
    return copy.deepcopy(_EMPTY_STORE)

def _save_store(store: dict) -> None:
    """Persist this session's store to disk so all workers can read it."""
    sid = session.get('username', '__anonymous__')
    path = _store_path(sid)
    with open(path, 'wb') as fh:
        pickle.dump(store, fh)

def _clear_store() -> None:
    """Delete this session's on-disk store (called on logout / new process)."""
    sid = session.get('username', '__anonymous__')
    path = _store_path(sid)
    try:
        os.remove(path)
    except FileNotFoundError:
        pass

def _get_ct_engine():
    """Lazy import of Cycle Time report_engine (ct_module version)."""
    import importlib.util, sys
    mod_name = 'ct_report_engine'
    if mod_name in sys.modules:
        return sys.modules[mod_name]
    spec = importlib.util.spec_from_file_location(
        mod_name, os.path.join(_CT_DIR, 'report_engine.py'))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[mod_name] = mod
    spec.loader.exec_module(mod)
    return mod

def _get_ct_app():
    """Lazy import of Cycle Time app module (for build_full_report etc.)."""
    import importlib.util, sys
    mod_name = 'ct_app_module'
    if mod_name in sys.modules:
        return sys.modules[mod_name]
    spec = importlib.util.spec_from_file_location(
        mod_name, os.path.join(_CT_DIR, 'app.py'))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[mod_name] = mod
    spec.loader.exec_module(mod)
    return mod

def _make_date_suffix(start_dt, end_dt):
    if start_dt is None:
        return ''
    s_day = start_dt.date()
    e_day = end_dt.date() if end_dt else None
    if e_day is None or s_day == e_day:
        return start_dt.strftime('%Y%m%d')
    return f"{start_dt.strftime('%Y%m%d')}&{end_dt.strftime('%Y%m%d')}"

# ── Auth guard ────────────────────────────────────────────────────────────────
def _require_login():
    if 'username' not in session:
        return redirect(url_for('login_page'))
    return None

# ── Page ──────────────────────────────────────────────────────────────────────
@ct.route('/')
def ct_page():
    guard = _require_login()
    if guard: return guard
    return render_template_string(CT_PAGE_HTML)

# ── Process ───────────────────────────────────────────────────────────────────
@ct.route('/process', methods=['POST', 'OPTIONS'])
def ct_process():
    if request.method == 'OPTIONS':
        return '', 200
    guard = _require_login()
    if guard: return guard
    try:
        eng = _get_ct_engine()
        ct_files    = request.files.getlist('ct_files')
        online_file = request.files.get('online_file')
        from_s      = request.form.get('from_dt', '')
        to_s        = request.form.get('to_dt', '')
        reason      = request.form.get('reason', '')
        thresh_queue    = float(request.form.get('thresh_queue',    45))
        thresh_lag      = float(request.form.get('thresh_lag',      45))
        thresh_duration = float(request.form.get('thresh_duration', 120))

        if not ct_files or ct_files[0].filename == '':
            return jsonify({'error': 'No Cycle Time files uploaded'}), 400
        if not online_file or online_file.filename == '':
            return jsonify({'error': 'No Online Data file uploaded'}), 400

        try:
            start_dt = datetime.fromisoformat(from_s) if from_s else None
            end_dt   = datetime.fromisoformat(to_s)   if to_s   else None
        except ValueError:
            start_dt = end_dt = None

        ct_ios = [io.BytesIO(f.read()) for f in ct_files]
        on_io  = io.BytesIO(online_file.read())

        df_ct, df_an, fl, exc, ah = eng.prepare_ct_data(
            ct_ios, start_dt, end_dt,
            queue_minutes=thresh_queue,
            lag_minutes=thresh_lag,
            duration_threshold=thresh_duration,
        )
        df_wb, df_rj, nw = eng.prepare_online_data(on_io, start_dt, end_dt)

        store = _load_store()
        store.update({
            'df_ct': df_ct, 'df_anomaly': df_an,
            'df_wb_total': df_wb, 'df_wb_rejected': df_rj,
            'failure_list': fl, 'exceedances': exc,
            'applicable_hours': ah, 'net_weight': nw,
            'start_dt': start_dt, 'end_dt': end_dt, 'reason': reason,
            'main_xlsx': None, 'summary_xlsx': None, 'ppt_bytes': None,
        })
        _save_store(store)
        return jsonify({'ok': True, 'ct_records': len(df_ct), 'wb_records': len(df_wb),
                        'exceedances': exc, 'applicable_hours': ah, 'net_weight': nw})
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

# ── Build main ────────────────────────────────────────────────────────────────
@ct.route('/build/main', methods=['POST', 'OPTIONS'])
def ct_build_main():
    if request.method == 'OPTIONS':
        return '', 200
    guard = _require_login()
    if guard: return guard
    try:
        store = _load_store()
        if store['df_ct'] is None:
            return jsonify({'error': 'Run /ct/process first'}), 400
        eng  = _get_ct_engine()
        xlsx = eng.build_main_report(
            store['df_ct'], store['df_anomaly'],
            store['df_wb_total'], store['df_wb_rejected'],
            store['failure_list'], store['start_dt'], store['end_dt'])
        store['main_xlsx'] = xlsx
        _save_store(store)
        return jsonify({'ok': True, 'size_kb': round(len(xlsx)/1024)})
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

# ── Build summary ─────────────────────────────────────────────────────────────
@ct.route('/build/summary', methods=['POST', 'OPTIONS'])
def ct_build_summary():
    if request.method == 'OPTIONS':
        return '', 200
    guard = _require_login()
    if guard: return guard
    try:
        store = _load_store()
        if store['df_wb_total'] is None:
            return jsonify({'error': 'Run /ct/process first'}), 400
        eng       = _get_ct_engine()
        demo_path = os.path.join(_CT_DIR, 'Demo.xlsx')
        xlsx      = eng.build_summary_report(
            store['df_wb_total'], demo_path,
            start_dt=store.get('start_dt'),
            end_dt=store.get('end_dt'),
        )
        store['summary_xlsx'] = xlsx
        _save_store(store)
        return jsonify({'ok': True, 'size_kb': round(len(xlsx)/1024)})
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

# ── Build PPT ─────────────────────────────────────────────────────────────────
@ct.route('/build/ppt', methods=['POST', 'OPTIONS'])
def ct_build_ppt():
    if request.method == 'OPTIONS':
        return '', 200
    guard = _require_login()
    if guard: return guard
    try:
        store = _load_store()
        if store['df_ct'] is None:
            return jsonify({'error': 'Run /ct/process first'}), 400
        if store['summary_xlsx'] is None:
            return jsonify({'error': 'Build summary first'}), 400
        eng = _get_ct_engine()
        ppt = eng.build_ppt_report(
            store['df_ct'], store['df_wb_total'],
            reason=store['reason'], exceedances=store['exceedances'],
            applicable_hours=store['applicable_hours'],
            summary_xlsx_bytes=store['summary_xlsx'],
            ppt_template_path=os.path.join(_CT_DIR, 'demo.pptx'))
        store['ppt_bytes'] = ppt
        _save_store(store)
        return jsonify({'ok': True, 'size_kb': round(len(ppt)/1024)})
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

# ── Download ──────────────────────────────────────────────────────────────────
@ct.route('/download/<kind>', methods=['GET'])
def ct_download(kind):
    guard = _require_login()
    if guard: return guard
    store = _load_store()
    _sd = store.get('start_dt')
    _ed = store.get('end_dt')
    _suffix = _make_date_suffix(_sd, _ed)
    _ct_name  = f"Cycle Time - {_suffix}" if _suffix else "Cycle Time"
    _hrq_name = f"APSG-Hourly Truck Quantity {_suffix}" if _suffix else "APSG-Hourly Truck Quantity"

    MAP = {
        'main':    ('main_xlsx',    f'{_ct_name}.xlsx',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
        'summary': ('summary_xlsx', f'{_hrq_name}.xlsx',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
        'ppt':     ('ppt_bytes',    f'{_ct_name}.pptx',
                    'application/vnd.openxmlformats-officedocument.presentationml.presentation'),
    }
    if kind not in MAP:
        return jsonify({'error': 'Unknown type'}), 404
    key, fname, mime = MAP[kind]
    data = store.get(key)
    if not data:
        return jsonify({'error': f'{kind} not generated yet'}), 404
    resp = make_response(data)
    resp.headers['Content-Type'] = mime
    resp.headers['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp


# ══════════════════════════════════════════════════════════════════════════════
#  CT PAGE HTML — follows ABSC Main theme exactly
# ══════════════════════════════════════════════════════════════════════════════
CT_PAGE_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Cycle Time Report — APSG</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&family=Poppins:wght@600;700;800&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --indigo:#6366F1;--indigo-l:#818CF8;--cyan:#22D3EE;
  --green:#10B981;--amber:#F59E0B;--red:#F87171;
  --text:#E8EEF8;--muted:#64748B;--border:rgba(99,102,241,0.15);
}
html,body{
  background-image:url('/static/bg.jpg') !important;
  background-size:cover !important;background-position:center !important;
  background-attachment:fixed !important;background-color:#08101E !important;
  font-family:'Inter',system-ui,sans-serif;min-height:100vh;color:var(--text);
  font-size:15px;font-weight:500;line-height:1.55;
}
body::before{content:'';position:fixed;inset:0;z-index:0;
  background:rgba(3,7,18,0.45);pointer-events:none;}
body>*{position:relative;z-index:1;}

/* ── Top bar — identical to ABSC Main ── */
.top-bar{position:sticky;top:0;z-index:200;height:40px;display:flex;align-items:center;
  padding:0 1.2rem;gap:.75rem;
  background:rgba(6,10,28,0.55) !important;backdrop-filter:blur(18px);
  border-bottom:1px solid rgba(255,255,255,0.08);}
.top-mini-brand{font-size:.75rem;font-weight:800;color:#6366F1;letter-spacing:.04em;white-space:nowrap;}
.top-sep{width:1px;height:18px;background:rgba(99,102,241,.18);}
.top-page-label{font-size:.78rem;font-weight:600;color:#FFFFFF;white-space:nowrap;}
.top-spacer{flex:1;}
.karthi-tag{font-size:.65rem;color:#6366F1;font-weight:700;white-space:nowrap;letter-spacing:.01em;}
.back-btn{background:rgba(99,102,241,0.18);border:1px solid rgba(99,102,241,0.35);
  border-radius:6px;padding:.22rem .7rem;font-size:.7rem;font-weight:600;
  color:#C5D5FF;text-decoration:none;white-space:nowrap;transition:all .2s;}
.back-btn:hover{background:rgba(99,102,241,0.32);}

/* ── Page body ── */
.page{padding:1rem 1.2rem 4rem;max-width:1100px;margin:0 auto;}

/* ── Cards ── */
.card{background:rgba(8,14,38,0.70);border:1px solid rgba(255,255,255,0.10);
  border-radius:14px;padding:1.2rem 1.4rem;margin-bottom:1rem;
  backdrop-filter:blur(16px);box-shadow:0 4px 32px rgba(0,0,0,0.35);}
.card-title{font-size:.9rem;font-weight:700;color:var(--indigo-l);margin-bottom:.9rem;
  display:flex;align-items:center;gap:.5rem;letter-spacing:.02em;}
.step-badge{display:inline-flex;align-items:center;justify-content:center;
  width:24px;height:24px;border-radius:50%;background:var(--indigo);
  color:#fff;font-size:.72rem;font-weight:800;flex-shrink:0;}

/* ── Side-by-side upload grid ── */
.upload-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;}
@media(max-width:600px){.upload-grid{grid-template-columns:1fr;}}
.upload-col-label{font-size:.68rem;font-weight:700;color:var(--muted);
  letter-spacing:.08em;text-transform:uppercase;margin-bottom:.4rem;}

/* ── Date dropdowns ── */
.date-sel{width:100%;padding:.72rem .9rem;
  background:rgba(5,9,28,.85);border:1.5px solid rgba(99,102,241,.45);
  border-radius:10px;color:#EEF3FF;font-size:15px;font-weight:600;
  font-family:'Inter',sans-serif;cursor:pointer;appearance:none;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%236366F1' stroke-width='2' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right .9rem center;padding-right:2.2rem;}
.date-sel:focus{outline:none;border-color:var(--indigo);box-shadow:0 0 0 3px rgba(99,102,241,.2);}
.date-sel option{background:#0A0F2E;color:#EEF3FF;font-size:15px;}
.date-sel:disabled{opacity:.45;cursor:not-allowed;}

/* ── Upload zones ── */
.upload-zone{border:2px dashed rgba(99,102,241,0.55);border-radius:11px;
  padding:1.4rem;text-align:center;cursor:pointer;
  background:rgba(6,12,34,0.55);transition:all .2s;position:relative;}
.upload-zone:hover,.upload-zone.dragover{border-color:var(--indigo);background:rgba(99,102,241,.08);}
.upload-zone.ok{border-color:var(--green);background:rgba(16,185,129,.06);border-style:solid;}
.upload-zone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%;}
.file-list{display:flex;flex-wrap:wrap;gap:.35rem;margin-top:.6rem;}
.file-tag{background:rgba(99,102,241,.12);border:1px solid rgba(99,102,241,.25);
  color:#818CF8;padding:.2rem .6rem;border-radius:20px;font-size:.72rem;
  display:flex;align-items:center;gap:.3rem;}
.file-tag button{background:none;border:none;color:#F87171;cursor:pointer;font-size:.85rem;line-height:1;}

/* ── Anomaly config ── */
.anom-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;}
@media(max-width:640px){.anom-grid{grid-template-columns:1fr;}}
.anom-field{background:rgba(8,20,50,.6);border:1.5px solid rgba(99,102,241,.25);
  border-radius:10px;padding:12px 14px;}
.anom-field label{display:block;font-size:12px;font-weight:700;color:var(--indigo-l);
  margin-bottom:4px;letter-spacing:.3px;}
.anom-sub{font-size:11px;color:rgba(190,210,255,.55);margin-bottom:8px;line-height:1.4;}
.anom-input-wrap{display:flex;align-items:center;gap:8px;}
.anom-input{flex:1;border:1.5px solid rgba(99,102,241,.40);border-radius:7px;
  padding:10px 10px;font-size:15px;font-weight:700;color:#EEF3FF;
  background:rgba(5,9,28,.80);text-align:center;}
.anom-input:focus{outline:none;border-color:var(--indigo);box-shadow:0 0 0 3px rgba(99,102,241,.18);}
.anom-unit{font-size:13px;font-weight:700;color:var(--muted);white-space:nowrap;}
.anom-note{font-size:11px;color:#805ad5;margin-top:6px;font-weight:600;}

/* ── Reason dropdown ── */
select.reason-sel{width:100%;padding:.72rem .9rem;
  background:rgba(5,9,28,.80);border:1.5px solid rgba(99,102,241,.40);
  border-radius:10px;color:#EEF3FF;font-size:15px;font-weight:500;font-family:'Inter',sans-serif;}
select.reason-sel:focus{outline:none;border-color:var(--indigo);}

/* ── Date inputs ── */
.date-row{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
@media(max-width:500px){.date-row{grid-template-columns:1fr;}}
.date-field label{display:block;font-size:.68rem;font-weight:700;color:var(--muted);
  letter-spacing:.08em;text-transform:uppercase;margin-bottom:.4rem;}
input[type=datetime-local]{width:100%;padding:.65rem .9rem;
  background:rgba(5,9,28,.80);border:1.5px solid rgba(99,102,241,.40);
  border-radius:10px;color:#EEF3FF;font-size:14px;font-family:'Inter',sans-serif;}
input[type=datetime-local]:focus{outline:none;border-color:var(--indigo);}

/* ── Generate button ── */
.btn-gen{width:100%;padding:15px;border:none;border-radius:10px;
  background:linear-gradient(135deg,#2b6cb0,#1a365d);color:#fff;
  font-size:15px;font-weight:700;cursor:pointer;display:flex;
  align-items:center;justify-content:center;gap:10px;transition:all .2s;}
.btn-gen:hover:not(:disabled){transform:translateY(-2px);box-shadow:0 8px 24px rgba(43,108,176,.4);}
.btn-gen:disabled{opacity:.55;cursor:not-allowed;transform:none;}

/* ── Progress cards (step indicators) ── */
.steps-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:.75rem;margin:.75rem 0;}
@media(max-width:640px){.steps-grid{grid-template-columns:1fr 1fr;}}
.step-item{border-radius:10px;padding:.85rem .9rem;text-align:center;border:1px solid rgba(255,255,255,.08);
  background:rgba(8,14,38,.6);transition:all .3s;}
.step-item.active{border-color:var(--amber);background:rgba(245,158,11,.06);}
.step-item.done{border-color:var(--green);background:rgba(16,185,129,.07);}
.step-item.error{border-color:var(--red);background:rgba(248,113,113,.06);}
.step-icon{font-size:1.4rem;margin-bottom:.35rem;}
.step-name{font-size:.78rem;font-weight:700;color:#EEF3FF;margin-bottom:.2rem;}
.step-stat{font-size:.72rem;color:var(--muted);}
.step-item.done .step-stat{color:var(--green);}
.step-item.error .step-stat{color:var(--red);}

/* ── Progress bar ── */
.prog-wrap{margin:.6rem 0;}
.prog-track{height:5px;background:rgba(99,102,241,.12);border-radius:3px;overflow:hidden;}
.prog-fill{height:100%;border-radius:3px;background:linear-gradient(90deg,#4338CA,#6366F1);
  width:0%;transition:width .4s;}
.prog-label{font-size:.75rem;font-weight:600;color:var(--muted);margin-top:.3rem;text-align:right;}

/* ── Log terminal ── */
.log-box{background:rgba(2,5,15,.85);border:1px solid rgba(99,102,241,.15);
  border-radius:8px;padding:.75rem 1rem;font-family:'Courier New',monospace;
  font-size:.78rem;line-height:1.8;max-height:180px;overflow-y:auto;margin:.5rem 0;}
.log-ok{color:#4ADE80;font-weight:600;} .log-err{color:#F87171;font-weight:600;} .log-norm{color:#94A3B8;}

/* ── Download section ── */
.dl-section{display:none;}
.success-banner{background:rgba(16,185,129,.09);border:1px solid rgba(16,185,129,.22);
  border-radius:10px;padding:.75rem 1rem;margin-bottom:1rem;
  color:#34D399;font-size:.82rem;font-weight:600;}
.btn-download-all{width:100%;padding:15px;border:none;border-radius:10px;
  background:linear-gradient(135deg,#276749,#48bb78);color:#fff;
  font-size:15px;font-weight:700;cursor:pointer;display:flex;
  align-items:center;justify-content:center;gap:10px;transition:all .2s;margin-bottom:.75rem;}
.btn-download-all:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(72,187,120,.4);}
.stats-row{display:flex;gap:.65rem;flex-wrap:wrap;margin-top:.75rem;}
.stat{flex:1;min-width:80px;background:rgba(99,102,241,.08);border:1px solid rgba(99,102,241,.15);
  border-radius:9px;padding:.75rem;text-align:center;}
.stat-val{font-size:1.5rem;font-weight:800;color:#818CF8;}
.stat-lbl{font-size:.62rem;color:var(--muted);margin-top:.15rem;}

/* ── Loading section ── */
.loading-section{display:none;}

/* ── Shimmer ── */
@keyframes shimmer{0%{background-position:-200% 0}100%{background-position:200% 0}}
.shimmer-bar{height:4px;background:linear-gradient(90deg,
  rgba(99,102,241,.1) 0%,rgba(99,102,241,.5) 50%,rgba(99,102,241,.1) 100%);
  background-size:200% 100%;animation:shimmer 1.5s infinite;border-radius:2px;margin:.5rem 0;}

/* ── Spinner ── */
@keyframes spin{to{transform:rotate(360deg)}}
.spin{display:inline-block;animation:spin 1s linear infinite;}

/* Fixed dev footer */
.apsg-footer{position:fixed;bottom:0;left:0;right:0;z-index:9999;
  text-align:center;padding:.28rem 1rem;
  background:rgba(4,7,20,.70);backdrop-filter:blur(8px);
  border-top:1px solid rgba(255,255,255,.07);
  font-size:11px;font-weight:600;color:rgba(200,220,255,.70);
  letter-spacing:.06em;pointer-events:none;user-select:none;}
</style>
</head>
<body>

<!-- Top bar — matches ABSC Main exactly -->
<div class="top-bar">
  <span class="top-mini-brand">APSG</span>
  <div class="top-sep"></div>
  <span class="top-page-label">⏱ Cycle Time Report</span>
  <div class="top-spacer"></div>
  <span class="karthi-tag">✦ Karthi</span>
  <a href="/" class="back-btn">← Dashboard</a>
</div>

<div class="page">

  <!-- STEP 1: Upload Files — CT left, Online right -->
  <div class="card">
    <div class="card-title"><span class="step-badge">1</span> Upload Files</div>
    <div class="upload-grid">

      <!-- LEFT: Cycle Time Files (primary — drives date dropdowns) -->
      <div>
        <div class="upload-col-label">📊 Cycle Time Files <span style="font-size:.62rem;color:var(--muted);font-weight:400">(one or more)</span></div>
        <div class="upload-zone" id="ctZone">
          <input type="file" id="ctFiles" multiple accept=".xlsx,.xls,.csv" onchange="addCTFiles(this.files)">
          <div style="font-size:1.5rem;margin-bottom:.3rem">📊</div>
          <div style="font-size:.82rem;font-weight:600">Drop CT files or click</div>
          <div style="font-size:.7rem;color:var(--muted);margin-top:.2rem">.xlsx · .xls · .csv — multiple OK</div>
        </div>
        <div class="file-list" id="ctFileList"></div>
      </div>

      <!-- RIGHT: Online Data File -->
      <div>
        <div class="upload-col-label">📋 Online Data File</div>
        <div class="upload-zone" id="onZone">
          <input type="file" id="onFile" accept=".xlsx,.xls,.csv" onchange="setOnFile(this.files[0])">
          <div style="font-size:1.5rem;margin-bottom:.3rem">📋</div>
          <div style="font-size:.82rem;font-weight:600">Drop Online file or click</div>
          <div style="font-size:.7rem;color:var(--muted);margin-top:.2rem">.xlsx · .xls · .csv</div>
        </div>
        <div id="onFileName" style="font-size:.73rem;color:var(--green);margin-top:.4rem;font-weight:600;min-height:1rem;"></div>
      </div>

    </div>
  </div>

  <!-- STEP 2: Date Range -->
  <div class="card">
    <div class="card-title"><span class="step-badge">2</span> Date &amp; Time Range
      <span id="date-auto-badge" style="display:none;margin-left:8px;font-size:.65rem;font-weight:700;background:rgba(16,185,129,.15);color:#34D399;border:1px solid rgba(16,185,129,.25);border-radius:5px;padding:.15rem .55rem;">✦ Dates loaded from file</span>
    </div>

    <!-- Info bar: times are always fixed -->
    <div style="display:flex;align-items:center;gap:.6rem;margin-bottom:.85rem;padding:.6rem .9rem;background:rgba(99,102,241,.07);border:1px solid rgba(99,102,241,.2);border-radius:9px;">
      <span style="font-size:1rem;">⏰</span>
      <span style="font-size:.82rem;font-weight:600;color:#C5D5FF;">
        Time is always &nbsp;<span style="color:var(--indigo-l);font-weight:700;">From → 07:00 AM &nbsp;·&nbsp; To → 05:00 AM</span>
        &nbsp;<span style="color:var(--muted);font-weight:400;font-size:.73rem;">— upload a file to select available dates below</span>
      </span>
    </div>

    <div class="date-row" id="date-inputs">
      <div class="date-field">
        <label style="font-size:.75rem;font-weight:700;color:var(--muted);letter-spacing:.06em;text-transform:uppercase;display:block;margin-bottom:.45rem;">From Date</label>
        <select id="from-dt-sel" class="date-sel" onchange="_syncHiddenDates()">
          <option value="">— Upload a file first —</option>
        </select>
        <input type="hidden" id="from-dt" value="">
      </div>
      <div class="date-field">
        <label style="font-size:.75rem;font-weight:700;color:var(--muted);letter-spacing:.06em;text-transform:uppercase;display:block;margin-bottom:.45rem;">To Date</label>
        <select id="to-dt-sel" class="date-sel" onchange="_syncHiddenDates()">
          <option value="">— Upload a file first —</option>
        </select>
        <input type="hidden" id="to-dt" value="">
      </div>
    </div>

    <div style="margin-top:.85rem;">
      <div style="font-size:.75rem;font-weight:700;color:var(--muted);letter-spacing:.08em;text-transform:uppercase;margin-bottom:.45rem;">Queue Condition</div>
      <select id="reason" class="reason-sel">
        <option value="There is a queue condition due to hourly loads are higher than expected.">Queue condition — hourly loads higher than expected</option>
        <option value="There is a queue condition due to heavy rain.">Queue condition — heavy rain</option>
        <option value="There is no queue condition.">No queue condition</option>
        <option value="There is a queue condition due to weighbridge breakdown.">Queue condition — weighbridge breakdown</option>
      </select>
    </div>
  </div>

  <!-- STEP 2b: Anomaly Thresholds -->
  <div class="card">
    <div class="card-title"><span class="step-badge" style="background:#805ad5">⚙</span> Anomaly Detection Thresholds <span style="font-size:10px;font-weight:400;color:var(--muted);margin-left:6px">rows exceeding any threshold → Anomaly sheet</span></div>
    <div class="anom-grid">
      <div class="anom-field">
        <label>⏱ Date Time Arrival → WB In Time</label>
        <div class="anom-sub">Pre-weighbridge queue time</div>
        <div class="anom-input-wrap">
          <input type="number" id="thresh-queue" class="anom-input" value="2" min="1" max="9999" step="1">
          <span class="anom-unit">min</span>
        </div>
        <div class="anom-note">Default: 2 minutes</div>
      </div>
      <div class="anom-field">
        <label>⏱ WB Out Time → Date Time Exit</label>
        <div class="anom-sub">Post-weighbridge lag time</div>
        <div class="anom-input-wrap">
          <input type="number" id="thresh-lag" class="anom-input" value="2" min="1" max="9999" step="1">
          <span class="anom-unit">min</span>
        </div>
        <div class="anom-note">Default: 2 minutes</div>
      </div>
      <div class="anom-field">
        <label>⏱ Date Time Arrival → Date Time Exit</label>
        <div class="anom-sub">Total trip duration</div>
        <div class="anom-input-wrap">
          <input type="number" id="thresh-duration" class="anom-input" value="70" min="1" max="9999" step="1">
          <span class="anom-unit">min</span>
        </div>
        <div class="anom-note">Default: 70 minutes</div>
      </div>
    </div>
  </div>

  <!-- STEP 3: Generate -->
  <div class="card">
    <div class="card-title"><span class="step-badge">3</span> Generate Report</div>
    <button class="btn-gen" id="btn-gen" onclick="runGenerate()">🚀 &nbsp; Generate Report</button>
  </div>

  <!-- Loading / progress -->
  <div class="loading-section" id="loading-section">
    <div class="card">
      <div class="shimmer-bar"></div>
      <div class="steps-grid" id="steps-grid">
        <div class="step-item" id="scard-0"><div class="step-icon" id="sicon-0">📁</div><div class="step-name">Processing Files</div><div class="step-stat" id="sstat-0">Waiting…</div></div>
        <div class="step-item" id="scard-1"><div class="step-icon" id="sicon-1">📊</div><div class="step-name">Cycle Time Report</div><div class="step-stat" id="sstat-1">Waiting…</div></div>
        <div class="step-item" id="scard-2"><div class="step-icon" id="sicon-2">📋</div><div class="step-name">Hourly Track</div><div class="step-stat" id="sstat-2">Waiting…</div></div>
        <div class="step-item" id="scard-3"><div class="step-icon" id="sicon-3">📑</div><div class="step-name">PowerPoint</div><div class="step-stat" id="sstat-3">Waiting…</div></div>
      </div>
      <div class="prog-wrap">
        <div class="prog-track"><div class="prog-fill" id="prog-fill"></div></div>
        <div class="prog-label" id="prog-label"></div>
      </div>
      <div class="log-box" id="log"></div>
    </div>
  </div>

  <!-- Download section -->
  <div class="dl-section" id="dl-section">
    <div class="card">
      <div class="card-title"><span class="step-badge">4</span> Download Your Reports</div>
      <div class="success-banner">✅ &nbsp; All 3 reports generated successfully and ready for download!</div>
      <button class="btn-download-all" onclick="downloadAll()">⬇ &nbsp; Download Report</button>
      <div class="stats-row" id="stats"></div>
    </div>
  </div>

</div><!-- /page -->

<div class="apsg-footer">✦ Internal Reporting Platform — APSG Staging Ground &nbsp;·&nbsp; Developed by Karthik</div>

<script>
// ── SheetJS CDN (lazy-loaded on first upload) ─────────────────────────────────
let _XLSX = null;
async function _loadXLSX() {
  if (_XLSX) return _XLSX;
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    s.onload = () => { _XLSX = window.XLSX; resolve(_XLSX); };
    s.onerror = () => reject(new Error('SheetJS failed to load'));
    document.head.appendChild(s);
  });
}

// ── File state ────────────────────────────────────────────────────────────────
const _files = { ct: [], on: null };

// ── Extract unique calendar dates from uploaded file ─────────────────────────
// Priority: WB In Time column → Date Time Arrival column → all cells scan.
// Returns sorted array of 'YYYY-MM-DD' strings.
async function _extractDatesFromFile(file) {
  try {
    const XLSX = await _loadXLSX();
    const buf  = await file.arrayBuffer();
    // Read with cellDates:true AND raw:true so we get both Date objects and raw values
    const wb   = XLSX.read(buf, { type: 'array', cellDates: true, raw: true });
    const ws   = wb.Sheets[wb.SheetNames[0]];

    // Get rows with raw values (Date objects, numbers, strings)
    const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
    // Also get formatted string rows for string-based date parsing
    const strRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false });

    if (rawRows.length < 2) return [];

    const fmt = d => {
      const y = d.getFullYear(), m = d.getMonth()+1, day = d.getDate();
      return `${y}-${String(m).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
    };

    // Find the target column index from the header row
    const header = rawRows[0].map(h => String(h || '').trim().toLowerCase());
    let colIdx = -1;

    // Try WB In Time first
    colIdx = header.findIndex(h =>
      h === 'wb in time' || h === 'wbintime' || h === 'wb_in_time' ||
      (h.includes('wb') && (h.includes('in time') || h === 'wb in'))
    );
    // Try Date Time Arrival as first fallback
    if (colIdx < 0) {
      colIdx = header.findIndex(h =>
        h.includes('arrival') || h === 'date time arrival' || h === 'datetime arrival'
      );
    }
    // Try any column with "date" or "time" in the name
    if (colIdx < 0) {
      colIdx = header.findIndex(h => h.includes('date') || h.includes('time'));
    }

    const dateSet = new Set();

    // ── Pass 1: scan the target column using raw values ──────────────────────
    if (colIdx >= 0) {
      for (let r = 1; r < rawRows.length; r++) {
        const raw = rawRows[r][colIdx];
        const str = strRows[r] ? strRows[r][colIdx] : null;
        const d = _parseCell(XLSX, raw, str);
        if (d) dateSet.add(fmt(d));
      }
    }

    // ── Pass 2: if still empty, scan ALL columns ──────────────────────────────
    if (dateSet.size === 0) {
      for (let r = 1; r < Math.min(rawRows.length, 3000); r++) {
        for (let c = 0; c < (rawRows[r] || []).length; c++) {
          const raw = rawRows[r][c];
          const str = strRows[r] ? strRows[r][c] : null;
          const d = _parseCell(XLSX, raw, str);
          if (d) dateSet.add(fmt(d));
        }
      }
    }

    return [...dateSet].sort();
  } catch(e) {
    console.error('Date extraction failed:', e);
    return [];
  }
}

// ── Parse a single cell value into a JS Date (or null) ───────────────────────
function _parseCell(XLSX, raw, str) {
  // 1. Already a JS Date object
  if (raw instanceof Date && !isNaN(raw) && raw.getFullYear() > 2000) return raw;

  // 2. Excel serial number (date: 40000–60000, datetime has decimal)
  if (typeof raw === 'number' && raw > 40000 && raw < 60000) {
    try {
      const info = XLSX.SSF.parse_date_code(Math.floor(raw));
      if (info && info.y > 2000) return new Date(info.y, info.m - 1, info.d);
    } catch(e) {}
  }

  // 3. String parsing — try the formatted value first, then the raw string
  const candidates = [str, raw].filter(v => typeof v === 'string' && v.trim());
  for (const s of candidates) {
    const t = s.trim();
    // ISO: 2026-04-21 or 2026-04-21T07:00:00
    const iso = t.match(/^(\d{4})[-\/](\d{2})[-\/](\d{2})/);
    if (iso) {
      const d = new Date(+iso[1], +iso[2]-1, +iso[3]);
      if (!isNaN(d) && d.getFullYear() > 2000) return d;
    }
    // Named month: 21-Apr-2026 or 21 Apr 2026
    const named = t.match(/^(\d{1,2})[\s\-](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[\s\-](\d{4})/i);
    if (named) {
      const mo = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
      const d = new Date(+named[3], mo[named[2].toLowerCase()], +named[1]);
      if (!isNaN(d) && d.getFullYear() > 2000) return d;
    }
    // DMY: 21/04/2026 or 21-04-2026
    const dmy = t.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (dmy) {
      const d = new Date(+dmy[3], +dmy[2]-1, +dmy[1]);
      if (!isNaN(d) && d.getFullYear() > 2000) return d;
    }
    // MDY fallback: 04/21/2026
    const mdy = t.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (mdy) {
      const d = new Date(t);
      if (!isNaN(d) && d.getFullYear() > 2000) return d;
    }
  }
  return null;
}

// ── Populate From/To dropdowns from date array ────────────────────────────────
function _populateDateDropdowns(dates, prevFrom, prevTo) {
  const fromSel = document.getElementById('from-dt-sel');
  const toSel   = document.getElementById('to-dt-sel');
  if (!dates.length) return;

  function _label(ymd) {
    // Use noon to avoid any timezone-shift issues flipping the date
    const d = new Date(ymd + 'T12:00:00');
    return d.toLocaleDateString('en-GB', { day:'2-digit', month:'short', year:'numeric', weekday:'short' });
  }

  const opts = dates.map(d => `<option value="${d}">${_label(d)}</option>`).join('');
  fromSel.innerHTML = opts;
  toSel.innerHTML   = opts;

  // Restore previous selection if it still exists, else default earliest/latest
  fromSel.value = (prevFrom && dates.includes(prevFrom)) ? prevFrom : dates[0];
  toSel.value   = (prevTo   && dates.includes(prevTo))   ? prevTo   : dates[dates.length - 1];
  _syncHiddenDates();

  document.getElementById('date-auto-badge').style.display = 'inline-block';
}

// ── Sync hidden datetime-local values (always 07:00 From / 05:00 To) ─────────
function _syncHiddenDates() {
  const fromDate = document.getElementById('from-dt-sel').value;
  const toDate   = document.getElementById('to-dt-sel').value;
  document.getElementById('from-dt').value = fromDate ? fromDate + 'T07:00' : '';
  document.getElementById('to-dt').value   = toDate   ? toDate   + 'T05:00' : '';
}

// ── File upload handlers ──────────────────────────────────────────────────────

// Scan ALL current CT files, merge all unique dates, populate dropdowns.
// Preserves current From/To selections if the dates still exist after a change.
async function _refreshCTDates() {
  if (_files.ct.length === 0) {
    // Reset dropdowns
    const fromSel = document.getElementById('from-dt-sel');
    const toSel   = document.getElementById('to-dt-sel');
    fromSel.innerHTML = '<option value="">— Upload a file first —</option>';
    toSel.innerHTML   = '<option value="">— Upload a file first —</option>';
    document.getElementById('from-dt').value = '';
    document.getElementById('to-dt').value   = '';
    document.getElementById('date-auto-badge').style.display = 'none';
    return;
  }

  // Remember current selections before rebuilding
  const prevFrom = document.getElementById('from-dt-sel').value;
  const prevTo   = document.getElementById('to-dt-sel').value;

  // Extract dates from EVERY CT file and merge into one sorted unique set
  const allDatesSet = new Set();
  for (const file of _files.ct) {
    const dates = await _extractDatesFromFile(file);
    dates.forEach(d => allDatesSet.add(d));
  }
  const allDates = [...allDatesSet].sort();

  if (allDates.length === 0) return;

  _populateDateDropdowns(allDates, prevFrom, prevTo);
}

async function addCTFiles(files) {
  Array.from(files).forEach(f => {
    if (!_files.ct.find(x => x.name === f.name && x.size === f.size)) _files.ct.push(f);
  });
  renderCTList();
  document.getElementById('ctZone').classList.toggle('ok', _files.ct.length > 0);
  await _refreshCTDates();
}

function removeCT(i) {
  _files.ct.splice(i, 1);
  renderCTList();
  document.getElementById('ctZone').classList.toggle('ok', _files.ct.length > 0);
  _refreshCTDates();  // Re-scan remaining files
}

function renderCTList() {
  document.getElementById('ctFileList').innerHTML = _files.ct.map((f,i) =>
    `<div class="file-tag">📄 ${esc(f.name)} <button onclick="removeCT(${i})">×</button></div>`
  ).join('');
}

async function setOnFile(f) {
  _files.on = f;
  document.getElementById('onFileName').textContent = f ? '✓ ' + f.name : '';
  document.getElementById('onZone').classList.toggle('ok', !!f);
  // Use Online file dates only if no CT files have been uploaded yet
  if (f && _files.ct.length === 0 && !document.getElementById('from-dt').value) {
    const dates = await _extractDatesFromFile(f);
    if (dates.length) _populateDateDropdowns(dates, '', '');
  }
}

// Drag-drop
const ctZ = document.getElementById('ctZone');
ctZ.addEventListener('dragover',  e => { e.preventDefault(); ctZ.classList.add('dragover'); });
ctZ.addEventListener('dragleave', () => ctZ.classList.remove('dragover'));
ctZ.addEventListener('drop', e => { e.preventDefault(); ctZ.classList.remove('dragover'); addCTFiles(e.dataTransfer.files); });

const onZ = document.getElementById('onZone');
onZ.addEventListener('dragover',  e => { e.preventDefault(); onZ.classList.add('dragover'); });
onZ.addEventListener('dragleave', () => onZ.classList.remove('dragover'));
onZ.addEventListener('drop', e => { e.preventDefault(); onZ.classList.remove('dragover'); setOnFile(e.dataTransfer.files[0]); });

// ── Step / progress helpers ───────────────────────────────────────────────────
function setStep(i, state, stat) {
  document.getElementById('scard-'+i).className = 'step-item' + (state ? ' '+state : '');
  const s = document.getElementById('sstat-'+i);
  if (s) s.textContent = stat || '';
}
function setProgress(pct, label) {
  document.getElementById('prog-fill').style.width = pct + '%';
  document.getElementById('prog-label').textContent = label || '';
}
function addLog(msg, cls) {
  const lb = document.getElementById('log');
  const d  = document.createElement('div');
  d.className = 'log-' + (cls || 'norm');
  d.textContent = '[' + new Date().toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit',second:'2-digit'}) + '] ' + msg;
  lb.appendChild(d);
  lb.scrollTop = lb.scrollHeight;
}

// ── Download all ──────────────────────────────────────────────────────────────
function downloadAll() {
  [['main','Full_Report.xlsx'],['summary','Hourly_Report.xlsx'],['ppt','Report.pptx']].forEach(([k,n],i) => {
    setTimeout(() => {
      const a = document.createElement('a');
      a.href = '/ct/download/' + k; a.download = n;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
    }, i * 700);
  });
}

// ── Main generate ─────────────────────────────────────────────────────────────
async function runGenerate() {
  if (!_files.ct.length) { alert('⚠️ Please upload at least one Cycle Time file.'); return; }
  if (!_files.on)        { alert('⚠️ Please upload the Online Data file.'); return; }

  const fromDt = document.getElementById('from-dt').value;
  const toDt   = document.getElementById('to-dt').value;
  if (!fromDt || !toDt) {
    alert('⚠️ No dates available yet — please upload a Cycle Time or Online file first so dates can be loaded.');
    return;
  }

  document.getElementById('btn-gen').disabled = true;
  document.getElementById('dl-section').style.display = 'none';
  document.getElementById('loading-section').style.display = 'block';
  document.getElementById('log').innerHTML = '';
  for (let i=0; i<4; i++) setStep(i,'','Waiting…');
  setProgress(0,'');

  const reason  = document.getElementById('reason').value;
  const threshQ = parseFloat(document.getElementById('thresh-queue').value)    || 2;
  const threshL = parseFloat(document.getElementById('thresh-lag').value)      || 2;
  const threshD = parseFloat(document.getElementById('thresh-duration').value) || 70;

  const fd = new FormData();
  _files.ct.forEach(f => fd.append('ct_files', f));
  fd.append('online_file', _files.on);
  fd.append('from_dt', fromDt);
  fd.append('to_dt',   toDt);
  fd.append('reason',  reason);
  fd.append('thresh_queue',    threshQ);
  fd.append('thresh_lag',      threshL);
  fd.append('thresh_duration', threshD);

  let j1;
  try {
    // ── Step 0: Process ─────────────────────────────────────────────────
    setStep(0, 'active', 'Processing…');
    setProgress(5, 'Reading and validating uploaded files…');
    const r1 = await fetch('/ct/process', {method:'POST', body:fd});
    try { j1 = await r1.json(); } catch(e) { j1 = {error:'HTTP '+r1.status}; }
    if (!r1.ok || j1.error) { setStep(0,'error','Failed'); addLog('❌ '+j1.error,'err'); done(); return; }
    setStep(0,'done',`CT:${j1.ct_records} WB:${j1.wb_records}`);
    addLog(`✅ Files processed — CT: ${j1.ct_records}, WB: ${j1.wb_records} records`,'ok');
    setProgress(20,'Building Cycle Time Report…');

    // ── Step 1: Main Excel ───────────────────────────────────────────────
    setStep(1,'active','Building…'); setProgress(25,'Building Cycle Time Report…');
    const r2 = await fetch('/ct/build/main',{method:'POST'});
    let j2; try{j2=await r2.json();}catch(e){j2={error:'Unexpected response'};}
    if(!r2.ok||j2.error){setStep(1,'error','Failed');addLog('❌ '+j2.error,'err');done();return;}
    setStep(1,'done',j2.size_kb+' KB'); addLog('✅ Cycle Time Report ready','ok');
    setProgress(65,'Building Hourly Track…');

    // ── Step 2: Summary ──────────────────────────────────────────────────
    setStep(2,'active','Building…'); setProgress(68,'Building Hourly Track…');
    const r3 = await fetch('/ct/build/summary',{method:'POST'});
    let j3; try{j3=await r3.json();}catch(e){j3={error:'Unexpected response'};}
    if(!r3.ok||j3.error){setStep(2,'error','Failed');addLog('❌ '+j3.error,'err');done();return;}
    setStep(2,'done',j3.size_kb+' KB'); addLog('✅ Hourly Track ready','ok');
    setProgress(80,'Generating PowerPoint…');

    // ── Step 3: PowerPoint ───────────────────────────────────────────────
    setStep(3,'active','Building…'); setProgress(82,'Generating PowerPoint…');
    const r4 = await fetch('/ct/build/ppt',{method:'POST'});
    let j4; try{j4=await r4.json();}catch(e){j4={error:'Unexpected response'};}
    if(!r4.ok||j4.error){setStep(3,'error','Failed');addLog('❌ '+j4.error,'err');done();return;}
    setStep(3,'done',j4.size_kb+' KB'); addLog('✅ PowerPoint ready','ok');
    setProgress(100,'✅ All reports complete!');

    addLog('🎉 All reports ready — click Download Report below.','ok');
    const s = j1;
    document.getElementById('stats').innerHTML = `
      <div class="stat"><div class="stat-val">${s.ct_records||0}</div><div class="stat-lbl">CT Records</div></div>
      <div class="stat"><div class="stat-val">${s.wb_records||0}</div><div class="stat-lbl">WB Records</div></div>
      <div class="stat"><div class="stat-val">${s.exceedances||0}</div><div class="stat-lbl">Exceedances</div></div>
      <div class="stat"><div class="stat-val">${s.applicable_hours||0}</div><div class="stat-lbl">Applicable Hours</div></div>`;
    document.getElementById('dl-section').style.display = 'block';
    document.getElementById('dl-section').scrollIntoView({behavior:'smooth'});

  } catch(e) {
    addLog('❌ Connection error: '+e.message,'err');
  }
  done();
}

function done() { document.getElementById('btn-gen').disabled = false; }
function esc(s) { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
</script>
</body>
</html>"""
