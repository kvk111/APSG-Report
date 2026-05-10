"""
server.py -- APSG Report Generator
Self-contained: all routes are defined here directly (no separate routes.py import issues).
Run: python server.py   OR double-click START_APP.bat
Render deployment: set PORT environment variable (Render sets this automatically).
"""

import io, os, sys, traceback, threading, webbrowser, time
from datetime import datetime
from flask import Flask, request, jsonify, make_response, send_from_directory

_HERE = os.path.dirname(os.path.abspath(__file__))
_STATIC = os.path.join(_HERE, 'static')

# ── App ───────────────────────────────────────────────────────────────────────
# CRITICAL: static_url_path MUST NOT be '' — that causes 405 on POST routes
app = Flask(__name__, static_folder=_STATIC, static_url_path='/static')

# ── CORS ──────────────────────────────────────────────────────────────────────
@app.after_request
def add_cors(resp):
    resp.headers['Access-Control-Allow-Origin']  = '*'
    resp.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return resp

@app.route('/', methods=['GET'])
@app.route('/index.html', methods=['GET'])
def home():
    return send_from_directory(_STATIC, 'index.html')

# ── In-memory store ───────────────────────────────────────────────────────────
_store = {
    'df_ct': None, 'df_anomaly': None,
    'df_wb_total': None, 'df_wb_rejected': None,
    'failure_list': [], 'exceedances': 0,
    'applicable_hours': 0, 'net_weight': 0.0,
    'start_dt': None, 'end_dt': None, 'reason': '',
    'main_xlsx': None, 'summary_xlsx': None, 'ppt_bytes': None,
}

# ── Lazy-import heavy modules so startup errors are caught cleanly ─────────────
def _get_engine():
    try:
        import report_engine as eng
        return eng
    except Exception as e:
        raise RuntimeError(f"Failed to import report_engine: {e}\n{traceback.format_exc()}")


def _make_date_suffix(start_dt, end_dt):
    """
    Single date : same calendar day  → 'YYYYMMDD'
    Date range  : different days     → 'YYYYMMDD&YYYYMMDD'
    """
    if start_dt is None:
        return ''
    s_day = start_dt.date()
    e_day = end_dt.date() if end_dt else None
    if e_day is None or s_day == e_day:
        return start_dt.strftime('%Y%m%d')
    return f"{start_dt.strftime('%Y%m%d')}&{end_dt.strftime('%Y%m%d')}"


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route('/report/process', methods=['POST', 'OPTIONS'])
def process():
    if request.method == 'OPTIONS':
        return '', 200
    global _store
    try:
        eng = _get_engine()
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

        _store.update({
            'df_ct': df_ct, 'df_anomaly': df_an,
            'df_wb_total': df_wb, 'df_wb_rejected': df_rj,
            'failure_list': fl, 'exceedances': exc,
            'applicable_hours': ah, 'net_weight': nw,
            'start_dt': start_dt, 'end_dt': end_dt, 'reason': reason,
            'main_xlsx': None, 'summary_xlsx': None, 'ppt_bytes': None,
        })
        return jsonify({'ok': True, 'ct_records': len(df_ct), 'wb_records': len(df_wb),
                        'exceedances': exc, 'applicable_hours': ah, 'net_weight': nw})
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/report/build/main', methods=['POST', 'OPTIONS'])
def build_main():
    if request.method == 'OPTIONS':
        return '', 200
    global _store
    try:
        if _store['df_ct'] is None:
            return jsonify({'error': 'Run /report/process first'}), 400
        eng  = _get_engine()
        xlsx = eng.build_main_report(
            _store['df_ct'], _store['df_anomaly'],
            _store['df_wb_total'], _store['df_wb_rejected'],
            _store['failure_list'], _store['start_dt'], _store['end_dt'])
        _store['main_xlsx'] = xlsx
        return jsonify({'ok': True, 'size_kb': round(len(xlsx)/1024)})
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/report/build/summary', methods=['POST', 'OPTIONS'])
def build_summary():
    if request.method == 'OPTIONS':
        return '', 200
    global _store
    try:
        if _store['df_wb_total'] is None:
            return jsonify({'error': 'Run /report/process first'}), 400
        eng       = _get_engine()
        demo_path = os.path.join(_HERE, 'Demo.xlsx')
        xlsx      = eng.build_summary_report(
            _store['df_wb_total'], demo_path,
            start_dt=_store['start_dt'], end_dt=_store['end_dt'])
        _store['summary_xlsx'] = xlsx
        return jsonify({'ok': True, 'size_kb': round(len(xlsx)/1024)})
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/report/build/ppt', methods=['POST', 'OPTIONS'])
def build_ppt():
    if request.method == 'OPTIONS':
        return '', 200
    global _store
    try:
        if _store['df_ct'] is None:
            return jsonify({'error': 'Run /report/process first'}), 400
        if _store['summary_xlsx'] is None:
            return jsonify({'error': 'Build summary first'}), 400
        eng = _get_engine()
        ppt = eng.build_ppt_report(
            _store['df_ct'], _store['df_wb_total'],
            reason=_store['reason'], exceedances=_store['exceedances'],
            applicable_hours=_store['applicable_hours'],
            summary_xlsx_bytes=_store['summary_xlsx'],
            ppt_template_path=os.path.join(_HERE, 'demo.pptx'))
        _store['ppt_bytes'] = ppt
        return jsonify({'ok': True, 'size_kb': round(len(ppt)/1024)})
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/report/download/<kind>', methods=['GET'])
def download(kind):
    global _store
    _suffix = _make_date_suffix(_store.get('start_dt'), _store.get('end_dt'))
    # Filenames per spec:
    #   Cycle Time:  "Cycle Time - YYYYMMDD"  or  "Cycle Time - YYYYMMDD&YYYYMMDD"
    #   Hourly:      "APSG-Hourly Truck Quantity YYYYMMDD&YYYYMMDD"
    _ct_date  = f"Cycle Time - {_suffix}" if _suffix else "Cycle Time"
    _hrq_name = f"APSG-Hourly Truck Quantity {_suffix}" if _suffix else "APSG-Hourly Truck Quantity"

    MAP = {
        'main':    ('main_xlsx',
                    f'{_ct_date}.xlsx',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
        'summary': ('summary_xlsx',
                    f'{_hrq_name}.xlsx',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
        'ppt':     ('ppt_bytes',
                    f'{_ct_date}.pptx',
                    'application/vnd.openxmlformats-officedocument.presentationml.presentation'),
    }
    if kind not in MAP:
        return jsonify({'error': 'Unknown type'}), 404
    key, fname, mime = MAP[kind]
    data = _store.get(key)
    if not data:
        return jsonify({'error': f'{kind} not generated yet'}), 404
    resp = make_response(data)
    resp.headers['Content-Type'] = mime
    resp.headers['Content-Disposition'] = f'attachment; filename="{fname}"'
    return resp


# ── Startup self-test ─────────────────────────────────────────────────────────
def _self_test():
    """Run before serving to catch config problems early."""
    with app.test_client() as c:
        r = c.post('/report/process')
        if r.status_code == 405:
            print("  [FATAL] POST /report/process returned 405 — routing bug!")
            sys.exit(1)
        if 'application/json' not in r.content_type:
            print(f"  [FATAL] Expected JSON, got {r.content_type}")
            sys.exit(1)
        print("  [OK] Self-test passed — routes are working correctly")


# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    print("=" * 60)
    print("  APSG Report Generator")
    print("=" * 60)
    print("  Checking dependencies...")

    missing = []
    for pkg in ['flask', 'pandas', 'openpyxl', 'matplotlib', 'numpy', 'xlrd']:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)

    if missing:
        print(f"  [ERROR] Missing packages: {', '.join(missing)}")
        print("  Run: pip install -r requirements.txt")
        sys.exit(1)
    print("  [OK] All dependencies present")

    _self_test()

    # Render sets PORT env var; local default is 5050
    _port = int(os.environ.get('PORT', 5050))
    _is_local = _port == 5050

    if _is_local:
        def _open_browser():
            time.sleep(1.5)
            webbrowser.open(f'http://localhost:{_port}')
        threading.Thread(target=_open_browser, daemon=True).start()
        print(f"  Browser will open automatically at http://localhost:{_port}")

    print(f"  Starting server on port {_port}…")
    print("  Keep this window open. Press Ctrl+C to stop.")
    print("=" * 60)

    app.run(host='0.0.0.0', port=_port, debug=False)

