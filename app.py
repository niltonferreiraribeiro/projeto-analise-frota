from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
import openpyxl
import os
from datetime import datetime
from collections import defaultdict

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = 'uploads'

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

ALLOWED_EXTENSIONS = {'xlsx', 'xlsm'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def parse_time(time_obj):
    if time_obj is None:
        return 0
    if isinstance(time_obj, str):
        try:
            parts = time_obj.split(':')
            h = int(parts[0]) if len(parts) > 0 else 0
            m = int(parts[1]) if len(parts) > 1 else 0
            s = int(parts[2]) if len(parts) > 2 else 0
            return h * 3600 + m * 60 + s
        except:
            return 0
    try:
        if hasattr(time_obj, 'hour'):
            return time_obj.hour * 3600 + time_obj.minute * 60 + time_obj.second
        return 0
    except:
        return 0

def fmt_time(seconds):
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    return "{}h{:02d}min".format(h, m)

def parse_df_meta(v):
    try:
        if isinstance(v, str):
            v = v.replace(',', '.')
        return float(v)
    except:
        return 90.50

def to_str(obj):
    if obj is None:
        return ""
    if hasattr(obj, 'strftime'):
        return obj.strftime('%d/%m/%Y')
    if hasattr(obj, 'hour') and not hasattr(obj, 'year'):
        return "{:02d}:{:02d}:{:02d}".format(obj.hour, obj.minute, obj.second)
    return str(obj)

def normalize_system(name):
    if not name:
        return name
    m = {
        'eletroeletrônico': 'Eletroeletrônico',
        'eletroeletronico': 'Eletroeletrônico',
        'hidráulico': 'Hidráulico',
        'hidraulico': 'Hidráulico',
        'ar condicionado': 'Ar Condicionado',
        'combustivél': 'Combustível',
        'combustivel': 'Combustível',
        'sci': 'SCI',
    }
    return m.get(name.lower().strip(), name)

def classify_failures(failures, fault_tracker, critical_failures, sys_failures):
    """Mutually exclusive classification: RECORRENTE > CRITICA > POTENCIAL > PONTUAL"""
    critical_set = set(id(f) for f in critical_failures)
    counts = {'RECORRENTE': 0, 'CRITICA': 0, 'POTENCIAL': 0, 'PONTUAL': 0}
    for f in failures:
        is_recurrent = f['fault'] and fault_tracker.get(f['fault'], {}).get('count', 0) >= 2
        is_critical = id(f) in critical_set
        is_potential = f['system'] and len(sys_failures.get(f['system'], [])) > 1
        if is_recurrent:
            counts['RECORRENTE'] += 1
        elif is_critical:
            counts['CRITICA'] += 1
        elif is_potential:
            counts['POTENCIAL'] += 1
        else:
            counts['PONTUAL'] += 1
    return counts

def process_excel(filepath, df_meta):
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        if 'Escavadeira' not in wb.sheetnames:
            return None, "Aba 'Escavadeira' não encontrada"
        if 'DF' not in wb.sheetnames:
            return None, "Aba 'DF' não encontrada"

        ws_f = wb['Escavadeira']
        ws_df = wb['DF']

        # === FAILURES ===
        failures = []
        eq_failures = defaultdict(list)
        sys_failures = defaultdict(list)
        subsys_failures = defaultdict(list)
        daily_failures = defaultdict(list)
        fault_tracker = defaultdict(lambda: {'count': 0, 'equipments': set(), 'dates': [], 'system': '', 'subsystem': ''})

        row = 2
        empty = 0
        while empty < 5:
            eq = ws_f['B{}'.format(row)].value
            if not eq:
                empty += 1
                row += 1
                continue
            empty = 0
            if not str(eq).startswith('94'):
                row += 1
                continue

            date_val = ws_f['A{}'.format(row)].value
            desc = ws_f['C{}'.format(row)].value or ""
            system = normalize_system(ws_f['D{}'.format(row)].value or "")
            subsystem = ws_f['E{}'.format(row)].value or ""
            start_t = ws_f['F{}'.format(row)].value
            end_t = ws_f['G{}'.format(row)].value
            dur = ws_f['H{}'.format(row)].value
            fault = ws_f['I{}'.format(row)].value or ""

            dur_sec = parse_time(dur)
            date_str = to_str(date_val)

            rec = {
                'date': date_str,
                'date_obj': date_val,
                'equipment': str(eq),
                'description': desc,
                'system': system,
                'subsystem': subsystem,
                'start': to_str(start_t),
                'end': to_str(end_t),
                'duration': fmt_time(dur_sec),
                'duration_sec': dur_sec,
                'fault': fault
            }

            failures.append(rec)
            eq_failures[str(eq)].append(rec)
            if system:
                sys_failures[system].append(rec)
            if subsystem:
                subsys_failures[subsystem].append(rec)
            if date_str:
                daily_failures[date_str].append(rec)
            if fault:
                fault_tracker[fault]['count'] += 1
                fault_tracker[fault]['equipments'].add(str(eq))
                fault_tracker[fault]['dates'].append(date_str)
                fault_tracker[fault]['system'] = system
                fault_tracker[fault]['subsystem'] = subsystem

            row += 1

        # === DF DATA ===
        equipment_df = {}
        for r in range(3, 10):
            eq_id = ws_df['F{}'.format(r)].value
            df_val = ws_df['G{}'.format(r)].value
            if eq_id and df_val is not None:
                try:
                    equipment_df[str(int(float(str(eq_id))))] = float(str(df_val).replace(',', '.'))
                except:
                    pass

        df_acumulado = 0
        for r in range(10, 25):
            label = ws_df['F{}'.format(r)].value
            if label and "acumulado" in str(label).lower():
                v = ws_df['G{}'.format(r)].value
                if v is not None:
                    try:
                        df_acumulado = float(str(v).replace(',', '.'))
                    except:
                        pass
                break

        wb.close()

        # === ANALYSIS ===
        total = len(failures)
        total_time_sec = sum(f['duration_sec'] for f in failures)

        # System stats
        sys_stats = []
        for s, flist in sorted(sys_failures.items(), key=lambda x: len(x[1]), reverse=True):
            c = len(flist)
            pct = (c / total * 100) if total > 0 else 0
            risk = 'ALTO' if c >= 4 else 'ALTO' if c >= 3 else 'MEDIO' if c >= 2 else 'BAIXO'
            sys_stats.append({'system': s, 'count': c, 'pct': round(pct, 1), 'risk': risk})

        # Equipment stats
        eq_stats = []
        for e, flist in sorted(eq_failures.items(), key=lambda x: len(x[1]), reverse=True):
            c = len(flist)
            pct = (c / total * 100) if total > 0 else 0
            df_val = equipment_df.get(e, 0)
            if c >= 5:
                status = 'CRITICO'
            elif c >= 3:
                status = 'ALTO'
            elif c >= 2:
                status = 'MEDIO'
            else:
                status = 'BAIXO'
            eq_stats.append({'equipment': e, 'count': c, 'pct': round(pct, 1), 'status': status, 'df': df_val})

        # All equipment DF (including those without failures)
        all_eq_df = []
        all_eq_ids = set(list(eq_failures.keys()) + [str(k) for k in equipment_df.keys()])
        for e in sorted(all_eq_ids):
            df_val = equipment_df.get(e, 0)
            fc = len(eq_failures.get(e, []))
            if df_val == 0 and fc == 0:
                st = 'PREVENTIVA'
            elif df_val < 85:
                st = 'CRITICO'
            elif df_val < df_meta:
                st = 'ALERTA'
            else:
                st = 'OK'
            all_eq_df.append({'equipment': e, 'df': df_val, 'failures': fc, 'status': st})

        # Recurrent patterns (subsystem level)
        subsys_dates = defaultdict(lambda: {'dates': [], 'count': 0, 'equipments': set()})
        for f in failures:
            if f['subsystem']:
                key = f['subsystem']
                subsys_dates[key]['dates'].append(f['date'])
                subsys_dates[key]['count'] += 1
                subsys_dates[key]['equipments'].add(f['equipment'])

        recurrent = []
        for sub, data in sorted(subsys_dates.items(), key=lambda x: x[1]['count'], reverse=True):
            if data['count'] >= 2:
                dates_str = ', '.join(sorted(set(data['dates']))[:5])
                if data['count'] >= 3:
                    interpretation = 'Degradação PROGRESSIVA'
                else:
                    interpretation = 'Padrão de desgaste'
                recurrent.append({
                    'subsystem': sub,
                    'count': data['count'],
                    'dates': dates_str,
                    'equipments': ', '.join(sorted(data['equipments'])),
                    'interpretation': interpretation
                })

        # Critical failures (equipment with DF < 85% or 5+ failures on that equipment)
        critical_failures = []
        for f in failures:
            eq = f['equipment']
            eq_df = equipment_df.get(eq, 100)
            eq_fc = len(eq_failures.get(eq, []))
            if eq_df < 85 or eq_fc >= 5:
                critical_failures.append(f)

        # Recommendations
        recommendations = {'CRITICA': [], 'ALTA': [], 'MEDIA': []}

        # Critical: equipment with very low DF or many failures
        action_num = 1
        for es in eq_stats[:3]:
            if es['count'] >= 3 or es['df'] < 85:
                top_sys = defaultdict(int)
                for f in eq_failures[es['equipment']]:
                    if f['system']:
                        top_sys[f['system']] += 1
                main_sys = max(top_sys.items(), key=lambda x: x[1])[0] if top_sys else 'Geral'
                recommendations['CRITICA'].append({
                    'num': action_num,
                    'title': 'Eq {} - Revisar sistema de {}'.format(es['equipment'], main_sys),
                    'detail': '{} falhas registradas | DF: {:.2f}%'.format(es['count'], es['df']),
                    'impact': 'Reduzir falhas recorrentes e recuperar DF'
                })
                action_num += 1

        # High: top systems
        for ss in sys_stats[:2]:
            if ss['count'] >= 3:
                recommendations['ALTA'].append({
                    'num': action_num,
                    'title': 'Inspeção preventiva no sistema: {}'.format(ss['system']),
                    'detail': '{} falhas ({:.1f}% do total)'.format(ss['count'], ss['pct']),
                    'impact': 'Prevenir novas ocorrências'
                })
                action_num += 1

        # Medium: developing patterns
        for r in recurrent[:2]:
            recommendations['MEDIA'].append({
                'num': action_num,
                'title': 'Monitorar padrão emergente: {}'.format(r['subsystem']),
                'detail': '{}x ocorrências'.format(r['count']),
                'impact': 'Evitar degradação progressiva'
            })
            action_num += 1

        # Pareto data (for chart)
        pareto_labels = [s['system'] for s in sys_stats]
        pareto_values = [s['count'] for s in sys_stats]
        pareto_cumulative = []
        cum = 0
        for v in pareto_values:
            cum += v
            pareto_cumulative.append(round(cum / total * 100, 1) if total > 0 else 0)

        # Daily failures for chart
        daily_labels = sorted(daily_failures.keys())
        daily_values = [len(daily_failures[d]) for d in daily_labels]

        # Equipment failures for chart
        eq_chart_labels = [e['equipment'] for e in eq_stats]
        eq_chart_values = [e['count'] for e in eq_stats]

        analysis = {
            'total': total,
            'df_acumulado': df_acumulado,
            'df_meta': df_meta,
            'df_status': 'OK' if df_acumulado >= df_meta else 'ALERTA' if df_acumulado >= 50 else 'CRITICO',
            'total_time': fmt_time(total_time_sec),
            'sys_stats': sys_stats,
            'eq_stats': eq_stats,
            'all_eq_df': all_eq_df,
            'equipment_df': equipment_df,
            'recurrent': recurrent,
            'critical_failures': critical_failures,
            'recommendations': recommendations,
            'classified': classify_failures(failures, fault_tracker, critical_failures, sys_failures),
            'pareto_labels': pareto_labels,
            'pareto_values': pareto_values,
            'pareto_cumulative': pareto_cumulative,
            'daily_labels': daily_labels,
            'daily_values': daily_values,
            'eq_chart_labels': eq_chart_labels,
            'eq_chart_values': eq_chart_values,
            'generation_time': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
            'month_year': datetime.now().strftime('%B %Y')
        }
        # PONTUAL already computed in classify_failures

        return analysis, None

    except Exception as e:
        import traceback
        return None, "Erro: {}".format(traceback.format_exc())


def generate_report(a):
    p = []

    # ======================== HTML START ========================
    p.append('''<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Relatório de Confiabilidade</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',Tahoma,sans-serif;background:#f0f2f5;color:#333;line-height:1.6}
.container{max-width:1200px;margin:0 auto;padding:20px}
.header{background:linear-gradient(135deg,#1a3a5c 0%,#2a5298 100%);color:#fff;padding:40px;border-radius:12px;margin-bottom:30px;text-align:center}
.header h1{font-size:32px;margin-bottom:8px}
.header .subtitle{font-size:16px;opacity:.9}
.header .meta{font-size:14px;opacity:.7;margin-top:5px}
.section{background:#fff;padding:30px;border-radius:12px;margin-bottom:25px;box-shadow:0 2px 8px rgba(0,0,0,.08)}
.section h2{font-size:22px;color:#1a3a5c;margin-bottom:20px;padding-bottom:12px;border-bottom:3px solid #2a5298}
.section h3{font-size:16px;color:#2a5298;margin:20px 0 12px}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:20px;margin-bottom:30px}
.kpi{background:#fff;padding:25px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.08);border-left:5px solid #2a5298;text-align:center}
.kpi .value{font-size:32px;font-weight:700;color:#1a3a5c}
.kpi .label{font-size:12px;color:#888;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px}
.kpi .sub{font-size:12px;margin-top:5px}
.kpi.green{border-left-color:#27ae60}.kpi.green .value{color:#27ae60}.kpi.green .sub{color:#27ae60}
.kpi.red{border-left-color:#e74c3c}.kpi.red .value{color:#e74c3c}.kpi.red .sub{color:#e74c3c}
.kpi.orange{border-left-color:#f39c12}.kpi.orange .value{color:#f39c12}.kpi.orange .sub{color:#f39c12}
table{width:100%;border-collapse:collapse;margin:15px 0}
th{background:linear-gradient(135deg,#2a5298,#1a3a5c);color:#fff;padding:14px;text-align:left;font-weight:600;font-size:13px}
td{padding:12px 14px;border-bottom:1px solid #eee;font-size:13px}
tr:hover{background:#f8f9fa}
.badge{display:inline-block;padding:4px 12px;border-radius:20px;font-size:11px;font-weight:700}
.badge-red{background:#fce4e4;color:#c0392b}.badge-orange{background:#fef5e7;color:#e67e22}
.badge-green{background:#e8f8f0;color:#27ae60}.badge-blue{background:#eaf2f8;color:#2980b9}
.badge-gray{background:#f0f0f0;color:#666}
.critical-day{background:#fce4e4;border-left:4px solid #e74c3c;padding:15px;border-radius:6px;margin-bottom:12px}
.critical-day .title{color:#c0392b;font-weight:700;font-size:15px}
.critical-day .cause{color:#666;font-size:13px;margin-top:5px}
.action-group h3{margin-bottom:15px}
.action-item{padding:15px;border-radius:6px;margin-bottom:12px;border-left:4px solid}
.action-critical{background:#fce4e4;border-left-color:#e74c3c}
.action-alta{background:#fef5e7;border-left-color:#f39c12}
.action-media{background:#eaf2f8;border-left-color:#3498db}
.action-title{font-weight:700;color:#1a3a5c;margin-bottom:5px}
.action-detail{font-size:13px;color:#555}
.status-box{background:#f0f8f0;border-left:4px solid #27ae60;padding:20px;border-radius:6px;margin-bottom:15px}
.status-box.alert{background:#fef5e7;border-left-color:#f39c12}
.status-box.critical{background:#fce4e4;border-left-color:#e74c3c}
.chart-container{position:relative;height:350px;margin:20px 0}
.chart-row{display:grid;grid-template-columns:1fr 1fr;gap:25px;margin:20px 0}
ul{margin:10px 0 10px 25px}li{margin-bottom:8px}
.footer{text-align:center;padding:20px;color:#999;font-size:12px;margin-top:30px}
@media(max-width:768px){.chart-row{grid-template-columns:1fr}.kpi-grid{grid-template-columns:1fr 1fr}}
@media print{body{background:#fff}.section{box-shadow:none;page-break-inside:avoid}}
</style></head><body><div class="container">''')

    # ======================== HEADER ========================
    df_acum = a['df_acumulado']
    df_meta = a['df_meta']
    diff = df_acum - df_meta

    p.append('<div class="header">')
    p.append('<h1>📊 Relatório de Confiabilidade</h1>')
    p.append('<div class="subtitle">{} | Frota de Escavadeiras GMO</div>'.format(a['month_year']))
    p.append('<div class="meta">Meta de DF: {:.2f}%</div>'.format(df_meta))
    p.append('</div>')

    # ======================== RESUMO EXECUTIVO ========================
    p.append('<div class="section"><h2>⚡ Resumo Executivo</h2>')

    kpi_df_class = 'green' if df_acum >= df_meta else 'orange' if df_acum >= 50 else 'red'
    kpi_df_sub = '✅ ACIMA da meta ({:+.2f}%)'.format(diff) if diff >= 0 else '⚠️ ABAIXO da meta ({:.2f}%)'.format(diff)

    p.append('<div class="kpi-grid">')
    p.append('<div class="kpi {}"><div class="label">DF Acumulada (Frota)</div><div class="value">{:.2f}%</div><div class="sub">{}</div></div>'.format(kpi_df_class, df_acum, kpi_df_sub))
    p.append('<div class="kpi"><div class="label">Meta de DF</div><div class="value">{:.2f}%</div><div class="sub">Meta definida</div></div>'.format(df_meta))
    p.append('<div class="kpi orange"><div class="label">Total de Falhas</div><div class="value">{}</div><div class="sub">⚠️ Monitorado</div></div>'.format(a['total']))
    p.append('<div class="kpi"><div class="label">Tempo Total Parado</div><div class="value">{}</div><div class="sub">Acumulado</div></div>'.format(a['total_time']))
    p.append('</div>')

    # Status box
    if diff >= 0:
        p.append('<div class="status-box"><strong>Status Geral: BOAS NOTÍCIAS ✅</strong><br>')
        p.append('A frota está <strong>acima da meta</strong> de DF ({:.2f}% vs {:.2f}%).'.format(df_acum, df_meta))
    else:
        p.append('<div class="status-box alert"><strong>Status Geral: ATENÇÃO ⚠️</strong><br>')
        p.append('A frota está <strong>abaixo da meta</strong> de DF ({:.2f}% vs {:.2f}%).'.format(df_acum, df_meta))

    n_rec = len(a['recurrent'])
    if n_rec > 0:
        p.append(' Porém, identificamos <strong>{} padrões de falhas</strong> que precisam ação imediata.'.format(n_rec))
    p.append('</div></div>')

    # ======================== DISPONIBILIDADE FÍSICA ========================
    p.append('<div class="section"><h2>📈 Disponibilidade Física (DF)</h2>')
    p.append('<h3>Performance da Frota</h3>')

    # DF stats
    df_values = [v for v in a['equipment_df'].values() if v > 0]
    df_max = max(df_values) if df_values else 0
    df_min = min(df_values) if df_values else 0
    df_avg = sum(df_values) / len(df_values) if df_values else 0
    above_meta = sum(1 for v in df_values if v >= df_meta)
    below_meta = sum(1 for v in df_values if v < df_meta)

    p.append('<ul>')
    p.append('<li><strong>DF Média:</strong> {:.2f}% ({})'.format(df_avg, 'Acima da meta' if df_avg >= df_meta else 'Abaixo da meta'))
    p.append('<li><strong>DF Máxima:</strong> {:.2f}%</li>'.format(df_max))
    p.append('<li><strong>DF Mínima:</strong> {:.2f}%</li>'.format(df_min))
    p.append('<li><strong>Equipamentos acima da meta:</strong> {} de {} ({:.1f}%)</li>'.format(above_meta, len(df_values), (above_meta/len(df_values)*100) if df_values else 0))
    p.append('<li><strong>Equipamentos abaixo da meta:</strong> {} ({:.1f}%)</li>'.format(below_meta, (below_meta/len(df_values)*100) if df_values else 0))
    p.append('</ul>')

    # Equipment DF table
    p.append('<h3>DF por Equipamento</h3>')
    p.append('<table><tr><th>Equipamento</th><th>DF (%)</th><th>Falhas</th><th>Status</th></tr>')
    for eq in a['all_eq_df']:
        if eq['status'] == 'CRITICO':
            badge = '<span class="badge badge-red">CRÍTICO</span>'
        elif eq['status'] == 'ALERTA':
            badge = '<span class="badge badge-orange">ALERTA</span>'
        elif eq['status'] == 'PREVENTIVA':
            badge = '<span class="badge badge-blue">PREVENTIVA</span>'
        else:
            badge = '<span class="badge badge-green">OK</span>'
        p.append('<tr><td><strong>{}</strong></td><td>{:.2f}%</td><td>{}</td><td>{}</td></tr>'.format(
            eq['equipment'], eq['df'], eq['failures'], badge))
    p.append('</table></div>')

    # ======================== DIAS CRÍTICOS ========================
    critical_eq = [eq for eq in a['all_eq_df'] if eq['df'] > 0 and eq['df'] < 85]
    if critical_eq:
        p.append('<div class="section"><h2>🔴 Dias Críticos (DF &lt; 85%)</h2>')
        for eq in critical_eq:
            p.append('<div class="critical-day">')
            p.append('<div class="title">🔴 Eq {}: {:.2f}% - CRÍTICO</div>'.format(eq['equipment'], eq['df']))
            # Find top failure for this equipment
            eq_f = [f for f in a['critical_failures'] if f['equipment'] == eq['equipment']]
            if eq_f:
                p.append('<div class="cause">Causa: {} ({}) - parada {}</div>'.format(
                    eq_f[0]['fault'][:80] if eq_f[0]['fault'] else eq_f[0]['system'],
                    eq_f[0]['system'], eq_f[0]['duration']))
            p.append('</div>')
        p.append('</div>')

    # ======================== FALHAS CRÍTICAS ========================
    if a['critical_failures']:
        p.append('<div class="section"><h2>🔴 Falhas Críticas Identificadas</h2>')
        p.append('<table><tr><th>Data</th><th>Equipamento</th><th>Sistema</th><th>Descrição</th><th>Duração</th></tr>')
        for f in a['critical_failures'][:15]:
            p.append('<tr><td>{}</td><td><strong>{}</strong></td><td>{}</td><td>{}</td><td>{}</td></tr>'.format(
                f['date'], f['equipment'], f['system'], f['fault'][:60], f['duration']))
        p.append('</table></div>')

    # ======================== PADRÕES E TENDÊNCIAS ========================
    p.append('<div class="section"><h2>📊 Padrões e Tendências</h2>')

    # Falhas por Sistema
    p.append('<h3>Falhas por Sistema (Todo Mês)</h3>')
    p.append('<table><tr><th>Sistema</th><th>Ocorrências</th><th>%</th><th>Risco</th></tr>')
    for s in a['sys_stats']:
        if s['risk'] == 'ALTO':
            badge = '<span class="badge badge-red">ALTO</span>'
        elif s['risk'] == 'MEDIO':
            badge = '<span class="badge badge-orange">MÉDIO</span>'
        else:
            badge = '<span class="badge badge-green">BAIXO</span>'
        p.append('<tr><td><strong>{}</strong></td><td>{}</td><td>{:.1f}%</td><td>{}</td></tr>'.format(
            s['system'], s['count'], s['pct'], badge))
    p.append('</table>')

    # Equipamentos Críticos
    p.append('<h3>Equipamentos Críticos (Todo Mês)</h3>')
    p.append('<table><tr><th>Equipamento</th><th>Falhas</th><th>%</th><th>Status</th></tr>')
    for e in a['eq_stats']:
        if e['status'] == 'CRITICO':
            badge = '<span class="badge badge-red">CRÍTICO</span>'
        elif e['status'] == 'ALTO':
            badge = '<span class="badge badge-orange">ALTO</span>'
        elif e['status'] == 'MEDIO':
            badge = '<span class="badge badge-orange">MÉDIO</span>'
        else:
            badge = '<span class="badge badge-green">BAIXO</span>'
        p.append('<tr><td><strong>{}</strong></td><td>{}</td><td>{:.1f}%</td><td>{}</td></tr>'.format(
            e['equipment'], e['count'], e['pct'], badge))
    p.append('</table>')

    # Padrões Recorrentes
    if a['recurrent']:
        p.append('<h3>📊 Padrões Recorrentes (2+ vezes)</h3><ul>')
        for r in a['recurrent']:
            p.append('<li><strong>{} ({}x):</strong> {} &rarr; {}</li>'.format(
                r['subsystem'], r['count'], r['dates'], r['interpretation']))
        p.append('</ul>')
    p.append('</div>')

    # ======================== DASHBOARD ========================
    p.append('<div class="section"><h2>📊 Dashboard</h2>')
    p.append('<div class="chart-row">')
    p.append('<div><h3>Pareto de Falhas por Sistema</h3><div class="chart-container"><canvas id="paretoChart"></canvas></div></div>')
    p.append('<div><h3>Falhas por Equipamento</h3><div class="chart-container"><canvas id="eqChart"></canvas></div></div>')
    p.append('</div>')
    p.append('<div class="chart-row">')
    p.append('<div><h3>Falhas por Dia</h3><div class="chart-container"><canvas id="dailyChart"></canvas></div></div>')
    p.append('<div><h3>DF por Equipamento</h3><div class="chart-container"><canvas id="dfChart"></canvas></div></div>')
    p.append('</div>')
    p.append('</div>')

    # ======================== AÇÕES RECOMENDADAS ========================
    p.append('<div class="section"><h2>✅ Ações Recomendadas</h2>')

    if a['recommendations']['CRITICA']:
        p.append('<h3>🚨 CRÍTICAS (Executar HOJE)</h3>')
        for act in a['recommendations']['CRITICA']:
            p.append('<div class="action-item action-critical">')
            p.append('<div class="action-title">{}. {}</div>'.format(act['num'], act['title']))
            p.append('<div class="action-detail">{} | Impacto: {}</div>'.format(act['detail'], act['impact']))
            p.append('</div>')

    if a['recommendations']['ALTA']:
        p.append('<h3>🔴 ALTAS (Executar esta semana)</h3>')
        for act in a['recommendations']['ALTA']:
            p.append('<div class="action-item action-alta">')
            p.append('<div class="action-title">{}. {}</div>'.format(act['num'], act['title']))
            p.append('<div class="action-detail">{} | Impacto: {}</div>'.format(act['detail'], act['impact']))
            p.append('</div>')

    if a['recommendations']['MEDIA']:
        p.append('<h3>🟡 MÉDIAS (Próximas 2 semanas)</h3>')
        for act in a['recommendations']['MEDIA']:
            p.append('<div class="action-item action-media">')
            p.append('<div class="action-title">{}. {}</div>'.format(act['num'], act['title']))
            p.append('<div class="action-detail">{} | Impacto: {}</div>'.format(act['detail'], act['impact']))
            p.append('</div>')
    p.append('</div>')

    # ======================== CONCLUSÃO ========================
    p.append('<div class="section"><h2>🎯 Conclusão</h2>')

    if diff >= 0:
        p.append('<div class="status-box"><strong>Status Geral: EXCELENTE ✅</strong><br>')
        p.append('A frota está acima da meta de DF ({:.2f}% vs {:.2f}%).'.format(df_acum, df_meta))
    else:
        p.append('<div class="status-box alert"><strong>Status Geral: ATENÇÃO ⚠️</strong><br>')
        p.append('A frota está abaixo da meta de DF ({:.2f}% vs {:.2f}%).'.format(df_acum, df_meta))

    if a['recurrent']:
        p.append(' Porém, há {} padrões recorrentes que precisam de atenção:'.format(len(a['recurrent'])))
        p.append('<ol>')
        for r in a['recurrent'][:3]:
            p.append('<li><strong>{}</strong> - {} ({}x)</li>'.format(r['subsystem'], r['interpretation'], r['count']))
        p.append('</ol>')
    p.append('</div>')

    p.append('<p><strong>Recomendação Final:</strong> Executar ações críticas nos próximos 2-3 dias para:</p>')
    p.append('<ul>')
    p.append('<li>✅ Evitar quedas de DF</li>')
    p.append('<li>✅ Prevenir falhas recorrentes</li>')
    p.append('<li>✅ Manter frota confiável</li>')
    p.append('<li>✅ Atingir/manter meta de {:.2f}%+</li>'.format(df_meta))
    p.append('</ul>')
    p.append('<p><strong>Próxima análise:</strong> {} (próximo dia útil)</p>'.format(
        (datetime.now()).strftime('%d/%m/%Y')))
    p.append('</div>')

    # ======================== FOOTER ========================
    p.append('<div class="footer">')
    p.append('<p>Relatório gerado automaticamente | {} | Sistema de Análise de Falhas e Confiabilidade</p>'.format(a['generation_time']))
    p.append('</div>')

    # ======================== CHART.JS SCRIPTS ========================
    import json as _json
    p.append('<script>')

    # Pareto Chart
    p.append('const paretoCtx=document.getElementById("paretoChart").getContext("2d");')
    p.append('new Chart(paretoCtx,{type:"bar",data:{')
    p.append('labels:' + _json.dumps(a['pareto_labels'], ensure_ascii=False) + ',')
    p.append('datasets:[{label:"Falhas",data:' + _json.dumps(a['pareto_values']) + ',backgroundColor:"rgba(42,82,152,0.8)",yAxisID:"y"},')
    p.append('{label:"% Acumulado",data:' + _json.dumps(a['pareto_cumulative']) + ',type:"line",borderColor:"#e74c3c",borderWidth:2,pointRadius:4,yAxisID:"y1"}]')
    p.append('},options:{responsive:true,maintainAspectRatio:false,scales:{y:{beginAtZero:true,position:"left"},y1:{beginAtZero:true,max:100,position:"right",grid:{drawOnChartArea:false}}}}});')

    # Equipment Chart
    p.append('const eqCtx=document.getElementById("eqChart").getContext("2d");')
    p.append('new Chart(eqCtx,{type:"bar",data:{')
    p.append('labels:' + _json.dumps(a['eq_chart_labels'], ensure_ascii=False) + ',')
    p.append('datasets:[{label:"Falhas",data:' + _json.dumps(a['eq_chart_values']) + ',backgroundColor:["#e74c3c","#e74c3c","#f39c12","#f39c12","#3498db","#27ae60"]}]')
    p.append('},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}}}});')

    # Daily Chart
    p.append('const dailyCtx=document.getElementById("dailyChart").getContext("2d");')
    p.append('new Chart(dailyCtx,{type:"line",data:{')
    p.append('labels:' + _json.dumps(a['daily_labels'], ensure_ascii=False) + ',')
    p.append('datasets:[{label:"Falhas/Dia",data:' + _json.dumps(a['daily_values']) + ',borderColor:"#2a5298",backgroundColor:"rgba(42,82,152,0.1)",fill:true,tension:0.3}]')
    p.append('},options:{responsive:true,maintainAspectRatio:false}});')

    # DF Chart
    df_labels = [str(e['equipment']) for e in a['all_eq_df']]
    df_vals = [e['df'] for e in a['all_eq_df']]
    df_meta_val = a['df_meta']
    df_colors = ['#e74c3c' if v < 85 else '#f39c12' if v < df_meta_val else '#27ae60' for v in df_vals]

    p.append('const dfCtx=document.getElementById("dfChart").getContext("2d");')
    p.append('new Chart(dfCtx,{type:"bar",data:{')
    p.append('labels:' + _json.dumps(df_labels, ensure_ascii=False) + ',')
    p.append('datasets:[{label:"DF %",data:' + _json.dumps(df_vals) + ',backgroundColor:' + _json.dumps(df_colors) + '}]')
    p.append('},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{min:0,max:100}}}});')

    p.append('</script>')
    p.append('</div></body></html>')

    return '\n'.join(p)


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/upload', methods=['POST'])
def upload():
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'Nenhum arquivo enviado'}), 400
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'Nenhum arquivo selecionado'}), 400
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': 'Apenas arquivos .xlsx são permitidos'}), 400

        df_meta = parse_df_meta(request.form.get('df_meta', '90.50'))
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        analysis, error = process_excel(filepath, df_meta)
        try:
            os.remove(filepath)
        except:
            pass

        if error:
            return jsonify({'success': False, 'error': error}), 400

        html_report = generate_report(analysis)
        return jsonify({'success': True, 'report': html_report}), 200

    except Exception as e:
        return jsonify({'success': False, 'error': 'Erro no servidor: {}'.format(str(e))}), 500


if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
