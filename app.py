import os
import json
import sqlite3
import base64
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, g
from openpyxl import load_workbook
from dotenv import load_dotenv
from utils.ar_parser import parse_ar_novedades, parse_ar_liquidacion, run_ar_control
from utils.mx_parser import parse_mx_ops, parse_mx_prenomina, run_mx_control
from utils.exporters import export_ar_control, export_mx_novedades, export_mx_control
from utils.db import init_db, save_period_ar, save_period_mx, get_metrics

load_dotenv()

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

ANTHROPIC_KEY = os.getenv('ANTHROPIC_API_KEY', '')

def get_db():
    if 'db' not in g:
        g.db = sqlite3.connect('moova_payroll.db')
        g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop('db', None)
    if db: db.close()

with app.app_context():
    init_db(sqlite3.connect('moova_payroll.db'))

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/ar')
def ar():
    return render_template('ar.html')

@app.route('/mx')
def mx():
    return render_template('mx.html')

@app.route('/metrics')
def metrics():
    return render_template('metrics.html')

# ── API KEY para el browser ──────────────────────────────────
@app.route('/api/config', methods=['GET'])
def get_config():
    # Expone solo lo necesario para el browser
    return jsonify({'anthropic_key': ANTHROPIC_KEY})

# ── API: AR CONTROL ─────────────────────────────────────────
@app.route('/api/ar/control', methods=['POST'])
def ar_control():
    nov_file = request.files.get('novedades')
    liq_file = request.files.get('liquidacion')
    manuales = request.form.get('manuales', '[]')
    if not nov_file or not liq_file:
        return jsonify({'error': 'Faltan archivos'}), 400
    try:
        manuales_data = json.loads(manuales)
        nov_data, periodo = parse_ar_novedades(nov_file, manuales_data)
        liq_data = parse_ar_liquidacion(liq_file)
        results = run_ar_control(nov_data, liq_data)
        db = get_db()
        save_period_ar(db, periodo, results)
        return jsonify({'success': True, 'periodo': periodo, 'results': results})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── API: AR EXPORT ──────────────────────────────────────────
@app.route('/api/ar/export', methods=['POST'])
def ar_export():
    data = request.get_json()
    periodo = data.get('periodo', 'PERIODO')
    results = data.get('results', [])
    filepath = export_ar_control(results, periodo)
    return send_file(filepath, as_attachment=True,
                     download_name=f'AR-control-nomina-{periodo.replace("/","-")}.xlsx')

# ── API: AR EXPORT NOVEDADES ────────────────────────────────
@app.route('/api/ar/export-novedades', methods=['POST'])
def ar_export_novedades():
    data = request.get_json()
    novedades = data.get('novedades', {})
    periodo = data.get('periodo', 'PERIODO')
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    import tempfile
    wb = Workbook()
    ws = wb.active
    ws.title = 'Novedades HR Strategy'
    ws.merge_cells('A1:D1')
    ws['A1'] = f'MOOVA · Novedades para HR Strategy · {periodo}'
    ws['A1'].font = Font(bold=True, size=12, color='2563EB')
    ws['A1'].alignment = Alignment(horizontal='center')
    headers = ['Apellido y Nombre', 'HH.EE 50% (hs)', 'HH.EE 100% (hs)', 'Total HH.EE']
    blue_fill = PatternFill("solid", fgColor="2563EB")
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = Font(bold=True, color='FFFFFF', size=10)
        c.fill = blue_fill
        c.alignment = Alignment(horizontal='center')
    for ri, (nombre, d) in enumerate(novedades.items(), 4):
        he50 = d.get('he50', 0) or 0
        he100 = d.get('he100', 0) or 0
        ws.cell(row=ri, column=1, value=nombre)
        ws.cell(row=ri, column=2, value=he50 or '')
        ws.cell(row=ri, column=3, value=he100 or '')
        ws.cell(row=ri, column=4, value=he50 + he100)
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 14
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    wb.save(tmp.name)
    per = periodo.replace('/', '-') if periodo else 'PERIODO'
    return send_file(tmp.name, as_attachment=True,
                     download_name=f'AR-novedades-HR-Strategy-{per}.xlsx')

# ── API: MX NOVEDADES ───────────────────────────────────────
@app.route('/api/mx/novedades', methods=['POST'])
def mx_novedades():
    ops_file = request.files.get('ops')
    manuales = request.form.get('manuales', '[]')
    if not ops_file:
        return jsonify({'error': 'Falta archivo de OPS'}), 400
    try:
        manuales_data = json.loads(manuales)
        colabs, periodo, tope, correcciones_lft = parse_mx_ops(ops_file, manuales_data)
        return jsonify({'success': True, 'periodo': periodo, 'tope': tope,
                        'colabs': colabs, 'correcciones_lft': correcciones_lft})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── API: MX EXPORT NOVEDADES ────────────────────────────────
@app.route('/api/mx/export-novedades', methods=['POST'])
def mx_export_novedades():
    data = request.get_json()
    colabs = data.get('colabs', {})
    periodo = data.get('periodo', 'PERIODO')
    filepath = export_mx_novedades(colabs, periodo)
    return send_file(filepath, as_attachment=True,
                     download_name=f'MX-novedades-estudio-{periodo.replace("/","-").replace(" ","-")}.xlsx')

# ── API: MX CONTROL ─────────────────────────────────────────
@app.route('/api/mx/control', methods=['POST'])
def mx_control():
    ops_file = request.files.get('ops')
    nom_file = request.files.get('nomina')
    manuales = request.form.get('manuales', '[]')
    if not nom_file:
        return jsonify({'error': 'Falta prenómina del estudio'}), 400
    try:
        manuales_data = json.loads(manuales)
        colabs, periodo, tope, _ = parse_mx_ops(ops_file, manuales_data) if ops_file else ({}, '', 9, [])
        empleados, periodo_nom = parse_mx_prenomina(nom_file)
        results = run_mx_control(colabs, empleados)
        per = periodo_nom or periodo
        db = get_db()
        save_period_mx(db, per, results, colabs)
        return jsonify({'success': True, 'periodo': per, 'results': results})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── API: MX EXPORT CONTROL ──────────────────────────────────
@app.route('/api/mx/export-control', methods=['POST'])
def mx_export_control():
    data = request.get_json()
    results = data.get('results', [])
    periodo = data.get('periodo', 'PERIODO')
    filepath = export_mx_control(results, periodo)
    return send_file(filepath, as_attachment=True,
                     download_name=f'MX-control-nomina-{periodo.replace("/","-").replace(" ","-")}.xlsx')

# ── API: MÉTRICAS ────────────────────────────────────────────
@app.route('/api/metrics', methods=['GET'])
def api_metrics():
    try:
        db = get_db()
        data = get_metrics(db)
        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
