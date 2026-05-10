"""
routes.py — Flask routes for /report/* endpoints
Registered as blueprint on the main app.
"""

import io, os, traceback
from flask import Blueprint, request, jsonify, make_response, current_app
from datetime import datetime

from report_engine import (prepare_ct_data, prepare_online_data,
                            build_main_report, build_summary_report, build_ppt_report)

report_bp = Blueprint('report', __name__, url_prefix='/report')

# In-memory store
_store = {
    'df_ct':          None,
    'df_anomaly':     None,
    'df_wb_total':    None,
    'df_wb_rejected': None,
    'failure_list':   [],
    'exceedances':    0,
    'applicable_hours': 0,
    'net_weight':     0.0,
    'start_dt':       None,
    'end_dt':         None,
    'reason':         '',
    'main_xlsx':      None,
    'summary_xlsx':   None,
    'ppt_bytes':      None,
}


@report_bp.route('/process', methods=['POST'])
def process():
    """Upload + parse files + apply date filter."""
    global _store
    try:
        ct_files   = request.files.getlist('ct_files')
        online_file = request.files.get('online_file')
        from_s      = request.form.get('from_dt','')
        to_s        = request.form.get('to_dt','')
        reason      = request.form.get('reason','')

        if not ct_files:
            return jsonify({'error': 'No Cycle Time files uploaded'}), 400
        if not online_file:
            return jsonify({'error': 'No Online Data file uploaded'}), 400

        try:
            start_dt = datetime.fromisoformat(from_s) if from_s else None
            end_dt   = datetime.fromisoformat(to_s)   if to_s   else None
        except ValueError:
            start_dt = end_dt = None

        # Read file bytes into BytesIO objects (Flask file objects are one-shot)
        ct_ios = [io.BytesIO(f.read()) for f in ct_files]
        on_io  = io.BytesIO(online_file.read())

        df_ct, df_an, fl, exc, ah = prepare_ct_data(ct_ios, start_dt, end_dt)
        df_wb, df_rj, nw           = prepare_online_data(on_io, start_dt, end_dt)

        _store.update({
            'df_ct':          df_ct,
            'df_anomaly':     df_an,
            'df_wb_total':    df_wb,
            'df_wb_rejected': df_rj,
            'failure_list':   fl,
            'exceedances':    exc,
            'applicable_hours': ah,
            'net_weight':     nw,
            'start_dt':       start_dt,
            'end_dt':         end_dt,
            'reason':         reason,
            'main_xlsx':      None,
            'summary_xlsx':   None,
            'ppt_bytes':      None,
        })

        return jsonify({
            'ok':              True,
            'ct_records':      len(df_ct),
            'wb_records':      len(df_wb),
            'exceedances':     exc,
            'applicable_hours': ah,
            'net_weight':      nw,
        })

    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@report_bp.route('/build/main', methods=['POST'])
def build_main():
    """Build 7-sheet Excel report."""
    global _store
    try:
        if _store['df_ct'] is None:
            return jsonify({'error': 'Run /report/process first'}), 400

        xlsx = build_main_report(
            _store['df_ct'], _store['df_anomaly'],
            _store['df_wb_total'], _store['df_wb_rejected'],
            _store['failure_list'],
            _store['start_dt'], _store['end_dt'],
        )
        _store['main_xlsx'] = xlsx
        return jsonify({'ok': True, 'size_kb': round(len(xlsx)/1024)})

    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@report_bp.route('/build/summary', methods=['POST'])
def build_summary():
    """Build 2-sheet Summary Excel."""
    global _store
    try:
        if _store['df_wb_total'] is None:
            return jsonify({'error': 'Run /report/process first'}), 400

        demo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Demo.xlsx')
        xlsx = build_summary_report(_store['df_wb_total'], demo_path)
        _store['summary_xlsx'] = xlsx
        return jsonify({'ok': True, 'size_kb': round(len(xlsx)/1024)})

    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@report_bp.route('/build/ppt', methods=['POST'])
def build_ppt():
    """Build 2-slide PowerPoint."""
    global _store
    try:
        if _store['df_ct'] is None:
            return jsonify({'error': 'Run /report/process first'}), 400
        if _store['summary_xlsx'] is None:
            return jsonify({'error': 'Build summary first'}), 400

        _dir = os.path.dirname(os.path.abspath(__file__))
        ppt_tmpl = os.path.join(_dir, 'demo.pptx')

        ppt = build_ppt_report(
            _store['df_ct'],
            _store['df_wb_total'],
            reason=_store['reason'],
            exceedances=_store['exceedances'],
            applicable_hours=_store['applicable_hours'],
            summary_xlsx_bytes=_store['summary_xlsx'],
            ppt_template_path=ppt_tmpl,
        )
        _store['ppt_bytes'] = ppt
        return jsonify({'ok': True, 'size_kb': round(len(ppt)/1024)})

    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


@report_bp.route('/download/<kind>')
def download(kind):
    """Download generated report."""
    global _store
    MAP = {
        'main':    ('main_xlsx',    'Full_Report_7Sheets.xlsx',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
        'summary': ('summary_xlsx', 'Summary_Report.xlsx',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
        'ppt':     ('ppt_bytes',    'Cycle_Time_Report.pptx',
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
