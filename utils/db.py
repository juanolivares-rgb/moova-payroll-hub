import json
from datetime import datetime

def init_db(db):
    db.executescript('''
        CREATE TABLE IF NOT EXISTS ar_periods (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            periodo TEXT NOT NULL,
            fecha_proceso TEXT NOT NULL,
            total_empleados INTEGER,
            con_diferencia INTEGER,
            neto_total REAL,
            diferencia_total REAL,
            total_he50 REAL,
            total_he100 REAL,
            results_json TEXT
        );

        CREATE TABLE IF NOT EXISTS mx_periods (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            periodo TEXT NOT NULL,
            fecha_proceso TEXT NOT NULL,
            total_colabs INTEGER,
            con_diferencia INTEGER,
            neto_total REAL,
            total_he_doble REAL,
            total_he_triple REAL,
            alertas_lft INTEGER,
            results_json TEXT,
            colabs_json TEXT
        );

        CREATE TABLE IF NOT EXISTS ar_empleados_hist (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            periodo TEXT,
            legajo INTEGER,
            nombre TEXT,
            he50_hs REAL,
            he100_hs REAL,
            neto REAL,
            diferencia REAL
        );

        CREATE TABLE IF NOT EXISTS mx_colabs_hist (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            periodo TEXT,
            nombre TEXT,
            he_doble REAL,
            he_triple REAL,
            he_dom REAL,
            he_fest REAL,
            total_he REAL,
            neto REAL,
            tiene_alerta INTEGER
        );
    ''')
    db.commit()


def save_period_ar(db, periodo, results):
    total = len(results)
    con_dif = len([r for r in results if r['hasDiff']])
    neto_total = sum(r['neto'] for r in results)
    dif_total = sum(r['dif'] for r in results)
    he50 = sum(r.get('he50Hs') or 0 for r in results)
    he100 = sum(r.get('he100Hs') or 0 for r in results)

    db.execute('''
        INSERT INTO ar_periods
        (periodo, fecha_proceso, total_empleados, con_diferencia, neto_total,
         diferencia_total, total_he50, total_he100, results_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (periodo, datetime.now().isoformat(), total, con_dif,
          neto_total, dif_total, he50, he100, json.dumps(results)))

    for r in results:
        db.execute('''
            INSERT INTO ar_empleados_hist
            (periodo, legajo, nombre, he50_hs, he100_hs, neto, diferencia)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (periodo, r['leg'], r['nombre'],
              r.get('he50Hs') or 0, r.get('he100Hs') or 0,
              r['neto'], r['dif']))
    db.commit()


def save_period_mx(db, periodo, results, colabs):
    total = len(results)
    con_dif = len([r for r in results if r['hasDiff']])
    neto_total = sum(r['neto'] for r in results)
    he_doble = sum(v.get('heDoble', 0) for v in colabs.values())
    he_triple = sum(v.get('heTriple', 0) for v in colabs.values())
    alertas = sum(1 for v in colabs.values() if v.get('alertas'))

    db.execute('''
        INSERT INTO mx_periods
        (periodo, fecha_proceso, total_colabs, con_diferencia, neto_total,
         total_he_doble, total_he_triple, alertas_lft, results_json, colabs_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (periodo, datetime.now().isoformat(), total, con_dif, neto_total,
          he_doble, he_triple, alertas,
          json.dumps(results), json.dumps(colabs)))

    for nombre, d in colabs.items():
        total_he = d.get('totalHE', d.get('heDoble', 0) + d.get('heTriple', 0))
        db.execute('''
            INSERT INTO mx_colabs_hist
            (periodo, nombre, he_doble, he_triple, he_dom, he_fest, total_he, neto, tiene_alerta)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (periodo, nombre,
              d.get('heDoble', 0), d.get('heTriple', 0),
              d.get('heDom', 0), d.get('heFest', 0), total_he,
              0, 1 if d.get('alertas') else 0))
    db.commit()


def get_metrics(db):
    # AR: evolución mensual
    ar_trend = db.execute('''
        SELECT periodo, total_empleados, con_diferencia, neto_total,
               diferencia_total, total_he50, total_he100,
               fecha_proceso
        FROM ar_periods ORDER BY fecha_proceso ASC
    ''').fetchall()

    # MX: evolución mensual
    mx_trend = db.execute('''
        SELECT periodo, total_colabs, con_diferencia, neto_total,
               total_he_doble, total_he_triple, alertas_lft,
               fecha_proceso
        FROM mx_periods ORDER BY fecha_proceso ASC
    ''').fetchall()

    # AR: ranking HH.EE por empleado
    ar_he_ranking = db.execute('''
        SELECT nombre,
               SUM(he50_hs) as total_he50,
               SUM(he100_hs) as total_he100,
               SUM(he50_hs + he100_hs) as total_he,
               COUNT(*) as periodos
        FROM ar_empleados_hist
        GROUP BY nombre
        ORDER BY total_he DESC
        LIMIT 10
    ''').fetchall()

    # AR: empleados con más diferencias
    ar_diff_ranking = db.execute('''
        SELECT nombre, COUNT(*) as periodos_con_diff,
               SUM(ABS(diferencia)) as dif_acum
        FROM ar_empleados_hist
        WHERE ABS(diferencia) > 1
        GROUP BY nombre
        ORDER BY periodos_con_diff DESC
        LIMIT 10
    ''').fetchall()

    # MX: ranking HH.EE
    mx_he_ranking = db.execute('''
        SELECT nombre,
               SUM(he_doble) as total_doble,
               SUM(he_triple) as total_triple,
               SUM(total_he) as total_he,
               COUNT(*) as periodos
        FROM mx_colabs_hist
        GROUP BY nombre
        ORDER BY total_he DESC
        LIMIT 10
    ''').fetchall()

    # MX: alertas recurrentes
    mx_alertas = db.execute('''
        SELECT nombre, SUM(tiene_alerta) as total_alertas, COUNT(*) as periodos
        FROM mx_colabs_hist
        WHERE tiene_alerta = 1
        GROUP BY nombre
        ORDER BY total_alertas DESC
        LIMIT 10
    ''').fetchall()

    def rows_to_list(rows):
        return [dict(r) for r in rows]

    return {
        'ar_trend': rows_to_list(ar_trend),
        'mx_trend': rows_to_list(mx_trend),
        'ar_he_ranking': rows_to_list(ar_he_ranking),
        'ar_diff_ranking': rows_to_list(ar_diff_ranking),
        'mx_he_ranking': rows_to_list(mx_he_ranking),
        'mx_alertas': rows_to_list(mx_alertas),
    }
