from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
import openpyxl
import os
import json as _json
from datetime import datetime, timedelta
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

def to_date(obj):
    """Convert to datetime.date for comparisons"""
    if obj is None:
        return None
    if hasattr(obj, 'date'):
        return obj.date()
    if hasattr(obj, 'year'):
        return obj
    if isinstance(obj, str):
        try:
            parts = obj.split('/')
            if len(parts) == 3:
                from datetime import date
                return date(int(parts[2]), int(parts[1]), int(parts[0]))
        except:
            pass
    return None

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
        'motriz': 'Motriz',
        'estrutura': 'Estrutura',
        'giro': 'Giro',
    }
    return m.get(name.lower().strip(), name)


def process_excel(filepath, df_meta):
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        if 'Escavadeira' not in wb.sheetnames:
            return None, "Aba 'Escavadeira' não encontrada"
        if 'DF' not in wb.sheetnames:
            return None, "Aba 'DF' não encontrada"

        ws_f = wb['Escavadeira']
        ws_df = wb['DF']

        # === READ ALL FAILURES (full year data) ===
        all_failures = []
        eq_failures_all = defaultdict(list)
        sys_failures_all = defaultdict(list)
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
            date_obj = to_date(date_val)

            rec = {
                'date': date_str,
                'date_obj': date_obj,
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

            all_failures.append(rec)
            eq_failures_all[str(eq)].append(rec)
            if system:
                sys_failures_all[system].append(rec)
            if fault:
                fault_tracker[fault]['count'] += 1
                fault_tracker[fault]['equipments'].add(str(eq))
                fault_tracker[fault]['dates'].append(date_str)
                fault_tracker[fault]['system'] = system
                fault_tracker[fault]['subsystem'] = subsystem

            row += 1

        # === SEPARATE: current month vs year ===
        today = datetime.now().date()
        current_month = today.month
        current_year = today.year
        seven_days_ago = today - timedelta(days=7)
        yesterday = today - timedelta(days=1)

        failures_month = [f for f in all_failures if f['date_obj'] and f['date_obj'].month == current_month and f['date_obj'].year == current_year]
        failures_year = all_failures
        failures_7d = [f for f in all_failures if f['date_obj'] and f['date_obj'] >= seven_days_ago]
        failures_yesterday = [f for f in all_failures if f['date_obj'] and f['date_obj'] == yesterday]

        # Monthly aggregations
        eq_failures = defaultdict(list)
        sys_failures = defaultdict(list)
        subsys_failures = defaultdict(list)
        daily_failures = defaultdict(list)

        for f in failures_month:
            eq_failures[f['equipment']].append(f)
            if f['system']:
                sys_failures[f['system']].append(f)
            if f['subsystem']:
                subsys_failures[f['subsystem']].append(f)
            if f['date']:
                daily_failures[f['date']].append(f)

        total_month = len(failures_month)
        total_year = len(failures_year)
        total_time_month = sum(f['duration_sec'] for f in failures_month)
        total_time_7d = sum(f['duration_sec'] for f in failures_7d)

        # === DF DATA (columns: Equipamento, DF, Semana, Diario) ===
        equipment_df = {}
        equipment_df_semana = {}
        equipment_df_diario = {}

        for r in range(3, 20):
            eq_id = ws_df['F{}'.format(r)].value
            if not eq_id:
                continue
            eq_str = str(eq_id).strip()
            if 'acumulado' in eq_str.lower():
                break
            df_val = ws_df['G{}'.format(r)].value
            sem_val = ws_df['H{}'.format(r)].value
            dia_val = ws_df['I{}'.format(r)].value
            try:
                eq_key = str(int(float(eq_str)))
            except:
                continue
            if df_val is not None:
                try:
                    equipment_df[eq_key] = float(str(df_val).replace(',', '.'))
                except:
                    pass
            if sem_val is not None:
                try:
                    equipment_df_semana[eq_key] = float(str(sem_val).replace(',', '.'))
                except:
                    pass
            if dia_val is not None:
                try:
                    equipment_df_diario[eq_key] = float(str(dia_val).replace(',', '.'))
                except:
                    pass

        # Fleet accumulated DF (mensal, semanal, diario)
        df_acumulado = 0
        df_acumulado_semana = 0
        df_acumulado_diario = 0
        for r in range(3, 30):
            label = ws_df['F{}'.format(r)].value
            if label and "acumulado" in str(label).lower():
                v = ws_df['G{}'.format(r)].value
                if v is not None:
                    try:
                        df_acumulado = float(str(v).replace(',', '.'))
                    except:
                        pass
                v2 = ws_df['H{}'.format(r)].value
                if v2 is not None:
                    try:
                        df_acumulado_semana = float(str(v2).replace(',', '.'))
                    except:
                        pass
                v3 = ws_df['I{}'.format(r)].value
                if v3 is not None:
                    try:
                        df_acumulado_diario = float(str(v3).replace(',', '.'))
                    except:
                        pass
                break

        # === WEEKLY DF DATA for chart (read additional week columns if available) ===
        # Look for week headers in row 2 starting from column J onwards
        weekly_df_data = []  # list of {'week': label, 'df': value}
        # First add the main Semana column as latest week
        if df_acumulado_semana > 0:
            sem_header = ws_df['H2'].value or 'Semana'
            weekly_df_data.append({'week': str(sem_header), 'df': df_acumulado_semana})

        wb.close()

        # === ANALYSIS ===
        # System stats with total downtime
        sys_stats = []
        for s, flist in sorted(sys_failures.items(), key=lambda x: sum(f['duration_sec'] for f in x[1]), reverse=True):
            c = len(flist)
            pct = (c / total_month * 100) if total_month > 0 else 0
            total_sec = sum(f['duration_sec'] for f in flist)
            risk = 'ALTO' if c >= 4 else 'ALTO' if c >= 3 else 'MEDIO' if c >= 2 else 'BAIXO'
            sys_stats.append({
                'system': s, 'count': c, 'pct': round(pct, 1),
                'risk': risk, 'total_time': fmt_time(total_sec), 'total_sec': total_sec
            })

        # Equipment stats sorted by total downtime
        eq_stats = []
        for e in sorted(eq_failures.keys(), key=lambda x: sum(f['duration_sec'] for f in eq_failures[x]), reverse=True):
            flist = eq_failures[e]
            c = len(flist)
            pct = (c / total_month * 100) if total_month > 0 else 0
            df_val = equipment_df.get(e, 0)
            total_sec = sum(f['duration_sec'] for f in flist)
            if c >= 5:
                status = 'CRITICO'
            elif c >= 3:
                status = 'ALTO'
            elif c >= 2:
                status = 'MEDIO'
            else:
                status = 'BAIXO'
            eq_stats.append({
                'equipment': e, 'count': c, 'pct': round(pct, 1),
                'status': status, 'df': df_val,
                'total_time': fmt_time(total_sec), 'total_sec': total_sec
            })

        # All equipment DF
        all_eq_df = []
        all_eq_ids = set(list(eq_failures.keys()) + [str(k) for k in equipment_df.keys()])
        for e in sorted(all_eq_ids):
            df_val = equipment_df.get(e, 0)
            df_sem = equipment_df_semana.get(e, 0)
            df_dia = equipment_df_diario.get(e, 0)
            fc = len(eq_failures.get(e, []))
            total_sec = sum(f['duration_sec'] for f in eq_failures.get(e, []))
            if df_val == 0 and fc == 0:
                st = 'PREVENTIVA'
            elif df_val < 85:
                st = 'CRITICO'
            elif df_val < df_meta:
                st = 'ALERTA'
            else:
                st = 'OK'
            all_eq_df.append({
                'equipment': e, 'df': df_val, 'df_semana': df_sem, 'df_diario': df_dia,
                'failures': fc, 'status': st,
                'total_time': fmt_time(total_sec), 'total_sec': total_sec
            })

        # DF performance stats
        df_values = [v for v in equipment_df.values() if v > 0]
        df_max = max(df_values) if df_values else 0
        df_min = min(df_values) if df_values else 0
        df_avg = sum(df_values) / len(df_values) if df_values else 0
        above_meta = sum(1 for v in df_values if v >= df_meta)
        below_meta = sum(1 for v in df_values if v < df_meta)

        # Recurrent patterns
        subsys_dates = defaultdict(lambda: {'dates': [], 'count': 0, 'equipments': set(), 'system': ''})
        for f in failures_month:
            if f['subsystem']:
                key = f['subsystem']
                subsys_dates[key]['dates'].append(f['date'])
                subsys_dates[key]['count'] += 1
                subsys_dates[key]['equipments'].add(f['equipment'])
                subsys_dates[key]['system'] = f['system']

        recurrent = []
        for sub, data in sorted(subsys_dates.items(), key=lambda x: x[1]['count'], reverse=True):
            if data['count'] >= 2:
                dates_str = ', '.join(sorted(set(data['dates']))[:5])
                if data['count'] >= 3:
                    interpretation = 'Degradação PROGRESSIVA'
                else:
                    interpretation = 'Padrão de desgaste'
                recurrent.append({
                    'subsystem': sub, 'count': data['count'],
                    'dates': dates_str, 'system': data['system'],
                    'equipments': ', '.join(sorted(data['equipments'])),
                    'interpretation': interpretation
                })

        # === CRITICAL PATTERNS - Top 3 Month + Top 3 Year ===
        def build_critical_patterns(failure_list, equip_df, top_n=3):
            pattern_groups = defaultdict(lambda: {
                'failures': [], 'total_sec': 0, 'dates': [], 'faults': [],
                'system': '', 'equipment': '', 'subsystem': ''
            })
            for f in failure_list:
                key = (f['system'], f['equipment'])
                pattern_groups[key]['failures'].append(f)
                pattern_groups[key]['total_sec'] += f['duration_sec']
                pattern_groups[key]['dates'].append(f['date'])
                pattern_groups[key]['faults'].append(f['fault'])
                pattern_groups[key]['system'] = f['system']
                pattern_groups[key]['equipment'] = f['equipment']
                pattern_groups[key]['subsystem'] = f['subsystem']

            patterns = []
            for key, grp in pattern_groups.items():
                eq = grp['equipment']
                eq_df = equip_df.get(eq, 100)
                count = len(grp['failures'])
                total_sec = grp['total_sec']
                unique_dates = sorted(set(grp['dates']))
                date_span = ''
                if len(unique_dates) >= 2:
                    date_span = '{}x em {} dias'.format(len(unique_dates), len(unique_dates))
                if count >= 3:
                    pattern_type = 'PADRÃO RECORRENTE'
                elif count >= 2:
                    pattern_type = 'PADRÃO EMERGENTE'
                else:
                    pattern_type = ''
                score = 0
                if eq_df < 85:
                    score += 100
                score += total_sec / 60
                score += count * 10
                if total_sec >= 1800 or eq_df < 85 or count >= 3:
                    fault_counts = defaultdict(int)
                    for ft in grp['faults']:
                        if ft:
                            fault_counts[ft] += 1
                    main_fault = max(fault_counts.items(), key=lambda x: x[1])[0] if fault_counts else grp['system']
                    df_impact = ''
                    if eq_df < 85:
                        df_impact = '-{:.0f}% no equipamento'.format(100 - eq_df)
                    if eq_df < 50 or (count >= 3 and total_sec > 7200):
                        risk = 'Risco ALTO - intervenção imediata'
                    elif count >= 3:
                        risk = 'Degradação progressiva - falha completa iminente'
                    elif count >= 2:
                        risk = 'Falha recorrente - investigar causa raiz'
                    else:
                        risk = 'Parada significativa - investigar causa raiz'
                    patterns.append({
                        'system': grp['system'], 'equipment': eq,
                        'main_fault': main_fault, 'count': count,
                        'dates': ', '.join(unique_dates), 'date_span': date_span,
                        'total_duration': fmt_time(total_sec), 'total_sec': total_sec,
                        'pattern_type': pattern_type, 'df_impact': df_impact,
                        'risk': risk, 'score': score, 'df': eq_df
                    })
            patterns.sort(key=lambda x: -x['score'])
            return patterns[:top_n]

        critical_month = build_critical_patterns(failures_month, equipment_df, 3)
        critical_year = build_critical_patterns(failures_year, equipment_df, 3)

        # Critical failures list
        critical_failures = []
        for f in failures_month:
            eq = f['equipment']
            eq_df = equipment_df.get(eq, 100)
            eq_fc = len(eq_failures.get(eq, []))
            if eq_df < 85 or eq_fc >= 5:
                critical_failures.append(f)

        # Recommendations
        recommendations = {'CRITICA': [], 'ALTA': [], 'MEDIA': []}
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
        for ss in sys_stats[:2]:
            if ss['count'] >= 3:
                recommendations['ALTA'].append({
                    'num': action_num,
                    'title': 'Inspeção preventiva no sistema: {}'.format(ss['system']),
                    'detail': '{} falhas ({:.1f}% do total) | Tempo parado: {}'.format(ss['count'], ss['pct'], ss['total_time']),
                    'impact': 'Prevenir novas ocorrências'
                })
                action_num += 1
        for r in recurrent[:2]:
            recommendations['MEDIA'].append({
                'num': action_num,
                'title': 'Monitorar padrão emergente: {}'.format(r['subsystem']),
                'detail': '{}x ocorrências | Equipamentos: {}'.format(r['count'], r['equipments']),
                'impact': 'Evitar degradação progressiva'
            })
            action_num += 1

        # Pareto data
        pareto_labels = [s['system'] for s in sys_stats]
        pareto_values = [s['count'] for s in sys_stats]
        pareto_cumulative = []
        cum = 0
        for v in pareto_values:
            cum += v
            pareto_cumulative.append(round(cum / total_month * 100, 1) if total_month > 0 else 0)

        # Daily data for charts
        daily_labels = sorted(daily_failures.keys())
        daily_values = [len(daily_failures[d]) for d in daily_labels]

        # Equipment chart data
        eq_chart_labels = [e['equipment'] for e in eq_stats]
        eq_chart_values = [e['total_sec'] / 3600 for e in eq_stats]  # hours

        # Classify
        classified = {'RECORRENTE': 0, 'CRITICA': 0, 'POTENCIAL': 0, 'PONTUAL': 0}
        critical_set = set(id(f) for f in critical_failures)
        for f in failures_month:
            is_rec = f['fault'] and fault_tracker.get(f['fault'], {}).get('count', 0) >= 2
            is_crit = id(f) in critical_set
            is_pot = f['system'] and len(sys_failures.get(f['system'], [])) > 1
            if is_rec:
                classified['RECORRENTE'] += 1
            elif is_crit:
                classified['CRITICA'] += 1
            elif is_pot:
                classified['POTENCIAL'] += 1
            else:
                classified['PONTUAL'] += 1

        # Predictions based on trends
        if total_month > 0 and daily_labels:
            days_elapsed = len(set(daily_labels))
            days_in_month = 30
            projected_failures = int(total_month / max(days_elapsed, 1) * days_in_month)
            projected_hours = total_time_month / max(days_elapsed, 1) * days_in_month
        else:
            projected_failures = 0
            projected_hours = 0

        analysis = {
            'total_month': total_month,
            'total_year': total_year,
            'total_time_month': fmt_time(total_time_month),
            'total_time_month_sec': total_time_month,
            'total_time_7d': fmt_time(total_time_7d),
            'total_time_7d_sec': total_time_7d,
            'failures_yesterday_count': len(failures_yesterday),
            'df_acumulado': df_acumulado,
            'df_acumulado_semana': df_acumulado_semana,
            'df_acumulado_diario': df_acumulado_diario,
            'df_meta': df_meta,
            'df_status': 'OK' if df_acumulado >= df_meta else 'ALERTA' if df_acumulado >= 50 else 'CRITICO',
            'df_max': df_max,
            'df_min': df_min,
            'df_avg': df_avg,
            'above_meta': above_meta,
            'below_meta': below_meta,
            'total_eq': len(df_values),
            'sys_stats': sys_stats,
            'eq_stats': eq_stats,
            'all_eq_df': all_eq_df,
            'equipment_df': equipment_df,
            'recurrent': recurrent,
            'critical_failures': critical_failures,
            'critical_month': critical_month,
            'critical_year': critical_year,
            'recommendations': recommendations,
            'classified': classified,
            'pareto_labels': pareto_labels,
            'pareto_values': pareto_values,
            'pareto_cumulative': pareto_cumulative,
            'daily_labels': daily_labels,
            'daily_values': daily_values,
            'eq_chart_labels': eq_chart_labels,
            'eq_chart_values': [round(v, 2) for v in eq_chart_values],
            'weekly_df_data': weekly_df_data,
            'projected_failures': projected_failures,
            'projected_hours': fmt_time(projected_hours),
            'generation_time': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
            'month_year': datetime.now().strftime('%B %Y'),
            'today_str': today.strftime('%d/%m/%Y')
        }

        return analysis, None

    except Exception as e:
        import traceback
        return None, "Erro: {}".format(traceback.format_exc())


# ============================================================
# REPORT GENERATION
# ============================================================
def generate_report(a):
    p = []

    # ======================== HTML HEAD ========================
    p.append('<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">')
    p.append('<meta name="viewport" content="width=device-width, initial-scale=1.0">')
    p.append('<title>Relatório de Confiabilidade</title>')
    p.append('<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>')
    p.append('<script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>')
    p.append("<style>")
    p.append("*{margin:0;padding:0;box-sizing:border-box}")
    p.append("body{font-family:Segoe UI,Tahoma,sans-serif;background:#f0f2f5;color:#333;line-height:1.6}")
    p.append(".container{max-width:1200px;margin:0 auto;padding:20px}")
    p.append(".btn-pdf{position:fixed;top:20px;right:20px;z-index:999;background:linear-gradient(135deg,#c0392b,#e74c3c);color:#fff;border:none;padding:14px 28px;border-radius:8px;font-size:15px;font-weight:700;cursor:pointer;box-shadow:0 4px 15px rgba(0,0,0,.3);transition:all .3s}")
    p.append(".btn-pdf:hover{transform:translateY(-2px);box-shadow:0 6px 20px rgba(0,0,0,.4)}")
    p.append(".header{background:linear-gradient(135deg,#1a3a5c 0%,#2a5298 100%);color:#fff;padding:40px;border-radius:12px;margin-bottom:30px;text-align:center}")
    p.append(".header h1{font-size:32px;margin-bottom:8px}")
    p.append(".header .subtitle{font-size:16px;opacity:.9}")
    p.append(".header .meta{font-size:14px;opacity:.7;margin-top:5px}")
    p.append(".section{background:#fff;padding:30px;border-radius:12px;margin-bottom:25px;box-shadow:0 2px 8px rgba(0,0,0,.08)}")
    p.append(".section h2{font-size:22px;color:#1a3a5c;margin-bottom:20px;padding-bottom:12px;border-bottom:3px solid #2a5298}")
    p.append(".section h3{font-size:16px;color:#2a5298;margin:20px 0 12px}")
    p.append(".kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:15px;margin-bottom:25px}")
    p.append(".kpi{background:#fff;padding:20px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.08);border-left:5px solid #2a5298;text-align:center}")
    p.append(".kpi .value{font-size:28px;font-weight:700;color:#1a3a5c}")
    p.append(".kpi .label{font-size:11px;color:#888;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px}")
    p.append(".kpi .sub{font-size:11px;margin-top:4px}")
    p.append(".kpi.green{border-left-color:#27ae60}.kpi.green .value{color:#27ae60}.kpi.green .sub{color:#27ae60}")
    p.append(".kpi.red{border-left-color:#e74c3c}.kpi.red .value{color:#e74c3c}.kpi.red .sub{color:#e74c3c}")
    p.append(".kpi.orange{border-left-color:#f39c12}.kpi.orange .value{color:#f39c12}.kpi.orange .sub{color:#f39c12}")
    p.append("table{width:100%;border-collapse:collapse;margin:15px 0}")
    p.append("th{background:linear-gradient(135deg,#2a5298,#1a3a5c);color:#fff;padding:12px;text-align:left;font-weight:600;font-size:12px}")
    p.append("td{padding:10px 12px;border-bottom:1px solid #eee;font-size:12px}")
    p.append("tr:hover{background:#f8f9fa}")
    p.append(".badge{display:inline-block;padding:4px 12px;border-radius:20px;font-size:11px;font-weight:700}")
    p.append(".badge-red{background:#fce4e4;color:#c0392b}.badge-orange{background:#fef5e7;color:#e67e22}")
    p.append(".badge-green{background:#e8f8f0;color:#27ae60}.badge-blue{background:#eaf2f8;color:#2980b9}")
    p.append(".badge-gray{background:#f0f0f0;color:#666}")
    p.append(".critical-day{background:#fce4e4;border-left:4px solid #e74c3c;padding:15px;border-radius:6px;margin-bottom:12px}")
    p.append(".critical-day .title{color:#c0392b;font-weight:700;font-size:15px}")
    p.append(".critical-day .cause{color:#666;font-size:13px;margin-top:5px}")
    p.append(".critical-card{background:#fdf2f2;border:1px solid #f5c6cb;border-left:5px solid #c0392b;border-radius:10px;padding:20px;margin-bottom:18px}")
    p.append(".critical-card .card-title{font-size:16px;font-weight:700;color:#c0392b;margin-bottom:10px}")
    p.append(".critical-card .card-body{font-size:13px;color:#555;line-height:1.8}")
    p.append(".action-item{padding:15px;border-radius:6px;margin-bottom:12px;border-left:4px solid}")
    p.append(".action-critical{background:#fce4e4;border-left-color:#e74c3c}")
    p.append(".action-alta{background:#fef5e7;border-left-color:#f39c12}")
    p.append(".action-media{background:#eaf2f8;border-left-color:#3498db}")
    p.append(".action-title{font-weight:700;color:#1a3a5c;margin-bottom:5px}")
    p.append(".action-detail{font-size:13px;color:#555}")
    p.append(".status-box{background:#f0f8f0;border-left:4px solid #27ae60;padding:20px;border-radius:6px;margin-bottom:15px}")
    p.append(".status-box.alert{background:#fef5e7;border-left-color:#f39c12}")
    p.append(".status-box.critical{background:#fce4e4;border-left-color:#e74c3c}")
    p.append(".chart-container{position:relative;height:350px;margin:20px 0}")
    p.append(".chart-row{display:grid;grid-template-columns:1fr 1fr;gap:25px;margin:20px 0}")
    p.append("ul{margin:10px 0 10px 25px}li{margin-bottom:8px}")
    p.append(".footer{text-align:center;padding:20px;color:#999;font-size:12px;margin-top:30px}")
    p.append(".grd-box{background:linear-gradient(135deg,#1a3a5c,#2a5298);color:#fff;padding:20px;border-radius:8px;text-align:center;margin-top:20px}")
    p.append("@media(max-width:768px){.chart-row{grid-template-columns:1fr}.kpi-grid{grid-template-columns:1fr 1fr}}")
    p.append("@media print{.btn-pdf{display:none}body{background:#fff}.section{box-shadow:none;page-break-inside:avoid}}")
    p.append("</style></head><body>")

    # PDF Download button
    p.append('<button class="btn-pdf" onclick="downloadPDF()">📥 Baixar PDF</button>')
    p.append('<div id="report-content" class="container">')

    df_acum = a['df_acumulado']
    df_meta = a['df_meta']
    diff = df_acum - df_meta
    df_sem = a['df_acumulado_semana']
    df_dia = a['df_acumulado_diario']

    # ======================== HEADER ========================
    p.append('<div class="header">')
    p.append('<h1>Relatório de Confiabilidade</h1>')
    p.append('<div class="subtitle">{} | Frota de Escavadeiras CAT - GMO</div>'.format(a['month_year']))
    p.append('<div class="meta">Meta de DF: {:.2f}% | Atualizado: {}</div>'.format(df_meta, a['generation_time']))
    p.append('</div>')

    # ======================== RESUMO EXECUTIVO ========================
    p.append('<div class="section"><h2>Resumo Executivo</h2>')

    kpi_df_class = 'green' if df_acum >= df_meta else 'orange' if df_acum >= 50 else 'red'
    kpi_sem_class = 'green' if df_sem >= df_meta else 'orange' if df_sem >= 50 else 'red'
    kpi_dia_class = 'green' if df_dia >= df_meta else 'orange' if df_dia >= 50 else 'red'
    diff_txt = '{:+.2f}%'.format(diff)

    p.append('<div class="kpi-grid">')
    # DF Acumulada Mensal
    p.append('<div class="kpi {}"><div class="label">DF Acumulada Mensal</div><div class="value">{:.2f}%</div><div class="sub">{} da meta</div></div>'.format(
        kpi_df_class, df_acum, diff_txt))
    # DF Últimos 7 dias
    if df_sem > 0:
        diff_sem = df_sem - df_meta
        p.append('<div class="kpi {}"><div class="label">DF Últimos 7 Dias</div><div class="value">{:.2f}%</div><div class="sub">{:+.2f}% da meta</div></div>'.format(
            kpi_sem_class, df_sem, diff_sem))
    # DF Dia Anterior
    if df_dia > 0:
        diff_dia = df_dia - df_meta
        p.append('<div class="kpi {}"><div class="label">DF Dia Anterior</div><div class="value">{:.2f}%</div><div class="sub">{:+.2f}% da meta</div></div>'.format(
            kpi_dia_class, df_dia, diff_dia))
    # Total falhas
    p.append('<div class="kpi orange"><div class="label">Total Falhas (Mês)</div><div class="value">{}</div><div class="sub">Acumulado mensal</div></div>'.format(a['total_month']))
    # Total horas paradas mês
    p.append('<div class="kpi red"><div class="label">Horas Paradas (Mês)</div><div class="value">{}</div><div class="sub">Acumulado mensal</div></div>'.format(a['total_time_month']))
    # Total horas paradas 7d
    p.append('<div class="kpi orange"><div class="label">Horas Paradas (7 Dias)</div><div class="value">{}</div><div class="sub">Últimos 7 dias</div></div>'.format(a['total_time_7d']))
    p.append('</div>')

    # Status box
    if diff >= 0:
        p.append('<div class="status-box"><strong>Status Geral: DENTRO DA META</strong><br>')
        p.append('A frota está <strong>acima da meta</strong> de DF ({:.2f}% vs {:.2f}%).'.format(df_acum, df_meta))
    else:
        box_class = 'critical' if df_acum < 85 else 'alert'
        p.append('<div class="status-box {}"><strong>Status Geral: ABAIXO DA META</strong><br>'.format(box_class))
        p.append('A frota está <strong>{:.2f}% abaixo da meta</strong> ({:.2f}% vs {:.2f}%).'.format(abs(diff), df_acum, df_meta))
    n_rec = len(a['recurrent'])
    if n_rec > 0:
        p.append(' Identificados <strong>{} padrões de falhas recorrentes</strong>.'.format(n_rec))
    p.append('</div></div>')

    # ======================== DISPONIBILIDADE FÍSICA ========================
    p.append('<div class="section"><h2>Disponibilidade Física (DF)</h2>')

    p.append('<div class="kpi-grid">')
    p.append('<div class="kpi"><div class="label">DF Acumulada Mês</div><div class="value">{:.2f}%</div></div>'.format(df_acum))
    p.append('<div class="kpi"><div class="label">Máxima (Equipamento)</div><div class="value">{:.2f}%</div></div>'.format(a['df_max']))
    p.append('<div class="kpi"><div class="label">Mínima (Equipamento)</div><div class="value">{:.2f}%</div></div>'.format(a['df_min']))
    p.append('<div class="kpi green"><div class="label">Acima da Meta</div><div class="value">{}</div><div class="sub">equipamentos</div></div>'.format(a['above_meta']))
    p.append('<div class="kpi red"><div class="label">Abaixo da Meta</div><div class="value">{}</div><div class="sub">equipamentos</div></div>'.format(a['below_meta']))
    p.append('</div>')

    # DF Chart placeholder (Meta vs Realizado)
    p.append('<h3>DF Mensal - Meta vs Realizado</h3>')
    p.append('<div class="chart-container"><canvas id="dfMensalChart"></canvas></div>')

    # Equipment DF table
    p.append('<h3>DF por Equipamento</h3>')
    p.append('<table><tr><th>Equipamento</th><th>DF Mês (%)</th><th>DF Semana (%)</th><th>DF Dia (%)</th><th>Falhas</th><th>Tempo Parado</th><th>Status</th></tr>')
    for eq in a['all_eq_df']:
        if eq['status'] == 'CRITICO':
            badge = '<span class="badge badge-red">CRÍTICO</span>'
        elif eq['status'] == 'ALERTA':
            badge = '<span class="badge badge-orange">ALERTA</span>'
        elif eq['status'] == 'PREVENTIVA':
            badge = '<span class="badge badge-blue">PREVENTIVA</span>'
        else:
            badge = '<span class="badge badge-green">OK</span>'
        p.append('<tr><td><strong>{}</strong></td><td>{:.2f}%</td><td>{:.2f}%</td><td>{:.2f}%</td><td>{}</td><td>{}</td><td>{}</td></tr>'.format(
            eq['equipment'], eq['df'], eq['df_semana'], eq['df_diario'], eq['failures'], eq['total_time'], badge))
    p.append('</table></div>')

    # ======================== DIAS CRÍTICOS ========================
    critical_eq = [eq for eq in a['all_eq_df'] if eq['df'] > 0 and eq['df'] < df_meta]
    if critical_eq:
        p.append('<div class="section"><h2>Dias Críticos (DF Abaixo da Meta)</h2>')
        for eq in sorted(critical_eq, key=lambda x: x['df']):
            diff_eq = eq['df'] - df_meta
            p.append('<div class="critical-day">')
            p.append('<div class="title">Eq {}: {:.2f}% | {:.2f}% abaixo da meta</div>'.format(eq['equipment'], eq['df'], abs(diff_eq)))
            eq_f = [f for f in a['critical_failures'] if f['equipment'] == eq['equipment']]
            if eq_f:
                p.append('<div class="cause">Principal causa: {} ({}) - parada total {}</div>'.format(
                    eq_f[0]['fault'][:80] if eq_f[0]['fault'] else eq_f[0]['system'],
                    eq_f[0]['system'], eq['total_time']))
            p.append('</div>')
        p.append('</div>')

    # ======================== FALHAS CRÍTICAS ========================
    p.append('<div class="section"><h2>Falhas Críticas Identificadas</h2>')

    # Top 3 do mês
    if a['critical_month']:
        p.append('<h3>Top 3 - Mês Atual</h3>')
        for i, cp in enumerate(a['critical_month'], 1):
            p.append('<div class="critical-card">')
            p.append('<div class="card-title">CRÍTICA #{}: {} - Eq {}</div>'.format(i, cp['system'], cp['equipment']))
            p.append('<div class="card-body">')
            p.append('<strong>O quê:</strong> {}'.format(cp['main_fault']))
            if cp['pattern_type']:
                p.append(' <span style="color:#c0392b;font-weight:700">({})</span>'.format(cp['pattern_type']))
            p.append('<br><strong>Quando:</strong> {}'.format(cp['dates']))
            if cp['date_span']:
                p.append(' ({})'.format(cp['date_span']))
            p.append('<br><strong>Duração:</strong> {} total'.format(cp['total_duration']))
            if cp['count'] > 1:
                p.append(' ({} ocorrências)'.format(cp['count']))
            if cp['df_impact']:
                p.append('<br><strong style="color:#c0392b">Impacto DF:</strong> {}'.format(cp['df_impact']))
            p.append('<br><strong style="color:#c0392b">Risco:</strong> {}'.format(cp['risk']))
            p.append('</div></div>')

    # Top 3 do ano
    if a['critical_year']:
        p.append('<h3>Top 3 - Acumulado do Ano</h3>')
        for i, cp in enumerate(a['critical_year'], 1):
            p.append('<div class="critical-card">')
            p.append('<div class="card-title">CRÍTICA #{}: {} - Eq {}</div>'.format(i, cp['system'], cp['equipment']))
            p.append('<div class="card-body">')
            p.append('<strong>O quê:</strong> {}'.format(cp['main_fault']))
            if cp['pattern_type']:
                p.append(' <span style="color:#c0392b;font-weight:700">({})</span>'.format(cp['pattern_type']))
            p.append('<br><strong>Quando:</strong> {}'.format(cp['dates']))
            if cp['date_span']:
                p.append(' ({})'.format(cp['date_span']))
            p.append('<br><strong>Duração:</strong> {} total'.format(cp['total_duration']))
            if cp['count'] > 1:
                p.append(' ({} ocorrências)'.format(cp['count']))
            p.append('<br><strong style="color:#c0392b">Risco:</strong> {}'.format(cp['risk']))
            p.append('</div></div>')
    p.append('</div>')

    # ======================== PADRÕES E TENDÊNCIAS ========================
    p.append('<div class="section"><h2>Padrões e Tendências</h2>')

    # Falhas por Sistema com tempo parado
    p.append('<h3>Falhas por Sistema (Acumulado Mês)</h3>')
    p.append('<table><tr><th>Sistema</th><th>Ocorrências</th><th>%</th><th>Tempo Total Parado</th><th>Risco</th></tr>')
    for s in a['sys_stats']:
        badge = '<span class="badge badge-red">ALTO</span>' if s['risk'] == 'ALTO' else '<span class="badge badge-orange">MÉDIO</span>' if s['risk'] == 'MEDIO' else '<span class="badge badge-green">BAIXO</span>'
        p.append('<tr><td><strong>{}</strong></td><td>{}</td><td>{:.1f}%</td><td>{}</td><td>{}</td></tr>'.format(
            s['system'], s['count'], s['pct'], s['total_time'], badge))
    p.append('</table>')

    # Equipamentos Críticos (sorted by total downtime)
    p.append('<h3>Equipamentos Críticos (por Tempo Parado)</h3>')
    p.append('<table><tr><th>Equipamento</th><th>Falhas</th><th>Tempo Total Parado</th><th>DF (%)</th><th>Status</th></tr>')
    for e in a['eq_stats']:
        badge = '<span class="badge badge-red">CRÍTICO</span>' if e['status'] == 'CRITICO' else '<span class="badge badge-orange">ALTO</span>' if e['status'] == 'ALTO' else '<span class="badge badge-orange">MÉDIO</span>' if e['status'] == 'MEDIO' else '<span class="badge badge-green">BAIXO</span>'
        p.append('<tr><td><strong>{}</strong></td><td>{}</td><td>{}</td><td>{:.2f}%</td><td>{}</td></tr>'.format(
            e['equipment'], e['count'], e['total_time'], e['df'], badge))
    p.append('</table>')

    # Padrões Recorrentes
    if a['recurrent']:
        p.append('<h3>Padrões Recorrentes (2+ vezes)</h3>')
        for r in a['recurrent']:
            p.append('<div style="background:#fff8e1;border-left:4px solid #f39c12;padding:15px;border-radius:6px;margin-bottom:12px">')
            p.append('<strong>{}</strong> ({}) - <span class="badge badge-orange">{}x ocorrências</span>'.format(r['subsystem'], r['system'], r['count']))
            p.append('<br><span style="font-size:13px;color:#555">Equipamentos: {} | Datas: {} | {}</span>'.format(r['equipments'], r['dates'], r['interpretation']))
            p.append('</div>')
    p.append('</div>')

    # ======================== DASHBOARD ========================
    p.append('<div class="section"><h2>Dashboard</h2>')
    p.append('<div class="chart-row">')
    p.append('<div><h3>Pareto de Falhas por Sistema</h3><div class="chart-container"><canvas id="paretoChart"></canvas></div></div>')
    p.append('<div><h3>Tempo Parado por Equipamento (h)</h3><div class="chart-container"><canvas id="eqChart"></canvas></div></div>')
    p.append('</div>')
    p.append('<div class="chart-row">')
    p.append('<div><h3>Falhas por Dia</h3><div class="chart-container"><canvas id="dailyChart"></canvas></div></div>')
    p.append('<div><h3>DF por Equipamento</h3><div class="chart-container"><canvas id="dfChart"></canvas></div></div>')
    p.append('</div>')
    p.append('</div>')

    # ======================== AÇÕES RECOMENDADAS ========================
    p.append('<div class="section"><h2>Ações Recomendadas</h2>')
    if a['recommendations']['CRITICA']:
        p.append('<h3>CRÍTICAS (Executar HOJE)</h3>')
        for act in a['recommendations']['CRITICA']:
            p.append('<div class="action-item action-critical"><div class="action-title">{}. {}</div><div class="action-detail">{} | Impacto: {}</div></div>'.format(act['num'], act['title'], act['detail'], act['impact']))
    if a['recommendations']['ALTA']:
        p.append('<h3>ALTAS (Executar esta semana)</h3>')
        for act in a['recommendations']['ALTA']:
            p.append('<div class="action-item action-alta"><div class="action-title">{}. {}</div><div class="action-detail">{} | Impacto: {}</div></div>'.format(act['num'], act['title'], act['detail'], act['impact']))
    if a['recommendations']['MEDIA']:
        p.append('<h3>MÉDIAS (Próximas 2 semanas)</h3>')
        for act in a['recommendations']['MEDIA']:
            p.append('<div class="action-item action-media"><div class="action-title">{}. {}</div><div class="action-detail">{} | Impacto: {}</div></div>'.format(act['num'], act['title'], act['detail'], act['impact']))
    p.append('</div>')

    # ======================== PREVISÕES ========================
    p.append('<div class="section"><h2>Previsões</h2>')
    p.append('<div class="kpi-grid">')
    p.append('<div class="kpi orange"><div class="label">Projeção Falhas (Mês)</div><div class="value">{}</div><div class="sub">Se tendência mantiver</div></div>'.format(a['projected_failures']))
    p.append('<div class="kpi red"><div class="label">Projeção Horas Paradas</div><div class="value">{}</div><div class="sub">Estimativa fim do mês</div></div>'.format(a['projected_hours']))
    p.append('</div>')
    if a['recurrent']:
        p.append('<p><strong>Tendências identificadas:</strong></p><ul>')
        for r in a['recurrent'][:3]:
            p.append('<li><strong>{}</strong> ({}x): {} - se não tratado, risco de parada prolongada.</li>'.format(r['subsystem'], r['count'], r['interpretation']))
        p.append('</ul>')
    p.append('</div>')

    # ======================== CONCLUSÃO ========================
    p.append('<div class="section"><h2>Conclusão</h2>')
    if diff >= 0:
        p.append('<div class="status-box"><strong>Status Geral: DENTRO DA META</strong><br>')
        p.append('A frota de escavadeiras está acima da meta de DF ({:.2f}% vs {:.2f}%).'.format(df_acum, df_meta))
    else:
        box_class = 'critical' if df_acum < 85 else 'alert'
        p.append('<div class="status-box {}"><strong>Status Geral: ABAIXO DA META</strong><br>'.format(box_class))
        p.append('A frota de escavadeiras está {:.2f}% abaixo da meta ({:.2f}% vs {:.2f}%).'.format(abs(diff), df_acum, df_meta))
    if a['recurrent']:
        p.append(' Foram identificados {} padrões recorrentes:'.format(len(a['recurrent'])))
        p.append('<ol>')
        for r in a['recurrent'][:3]:
            p.append('<li><strong>{}</strong> - {} ({}x)</li>'.format(r['subsystem'], r['interpretation'], r['count']))
        p.append('</ol>')
    p.append('</div>')

    p.append('<p><strong>Recomendação Final:</strong></p><ul>')
    p.append('<li>Executar ações críticas nos próximos 2-3 dias</li>')
    p.append('<li>Monitorar equipamentos com DF abaixo da meta</li>')
    p.append('<li>Acompanhar padrões recorrentes de falhas</li>')
    p.append('<li>Meta: atingir/manter {:.2f}%+ de DF</li>'.format(df_meta))
    p.append('</ul>')
    p.append('<p><strong>Próxima análise:</strong> {} (próximo dia útil)</p>'.format(a['today_str']))

    # GRD Box
    p.append('<div class="grd-box">')
    p.append('<strong>Engenharia de Manutenção e Confiabilidade - Setor GRD</strong><br>')
    p.append('Relatório gerado automaticamente pelo Sistema de Análise de Falhas e Confiabilidade')
    p.append('</div>')
    p.append('</div>')

    # ======================== FOOTER ========================
    p.append('<div class="footer">')
    p.append('<p>Relatório gerado automaticamente | {} | Engenharia GRD - Sistema de Análise de Falhas e Confiabilidade</p>'.format(a['generation_time']))
    p.append('</div>')

    # ======================== CHART.JS SCRIPTS ========================
    p.append('<script>')

    # DF Mensal Chart (Meta vs Realizado per equipment)
    df_eq_labels = [str(e['equipment']) for e in a['all_eq_df']]
    df_eq_vals = [e['df'] for e in a['all_eq_df']]
    meta_line = [df_meta] * len(df_eq_labels)

    p.append('const dfMensalCtx=document.getElementById("dfMensalChart").getContext("2d");')
    p.append('new Chart(dfMensalCtx,{type:"bar",data:{')
    p.append('labels:' + _json.dumps(df_eq_labels, ensure_ascii=False) + ',')
    df_colors = ['#e74c3c' if v < 85 else '#f39c12' if v < df_meta else '#27ae60' for v in df_eq_vals]
    p.append('datasets:[{label:"DF Realizado %",data:' + _json.dumps(df_eq_vals) + ',backgroundColor:' + _json.dumps(df_colors) + '},')
    p.append('{label:"Meta ' + str(df_meta) + '%",data:' + _json.dumps(meta_line) + ',type:"line",borderColor:"#c0392b",borderWidth:2,borderDash:[5,5],pointRadius:0,fill:false}]')
    p.append('},options:{responsive:true,maintainAspectRatio:false,scales:{y:{min:0,max:100}},plugins:{legend:{position:"top"}}}});')

    # Pareto Chart
    p.append('const paretoCtx=document.getElementById("paretoChart").getContext("2d");')
    p.append('new Chart(paretoCtx,{type:"bar",data:{')
    p.append('labels:' + _json.dumps(a['pareto_labels'], ensure_ascii=False) + ',')
    p.append('datasets:[{label:"Falhas",data:' + _json.dumps(a['pareto_values']) + ',backgroundColor:"rgba(42,82,152,0.8)",yAxisID:"y"},')
    p.append('{label:"% Acumulado",data:' + _json.dumps(a['pareto_cumulative']) + ',type:"line",borderColor:"#e74c3c",borderWidth:2,pointRadius:4,yAxisID:"y1"}]')
    p.append('},options:{responsive:true,maintainAspectRatio:false,scales:{y:{beginAtZero:true,position:"left"},y1:{beginAtZero:true,max:100,position:"right",grid:{drawOnChartArea:false}}}}});')

    # Equipment Chart (hours)
    p.append('const eqCtx=document.getElementById("eqChart").getContext("2d");')
    p.append('new Chart(eqCtx,{type:"bar",data:{')
    p.append('labels:' + _json.dumps(a['eq_chart_labels'], ensure_ascii=False) + ',')
    eq_colors = ['#e74c3c' if v > 5 else '#f39c12' if v > 2 else '#27ae60' for v in a['eq_chart_values']]
    p.append('datasets:[{label:"Horas Paradas",data:' + _json.dumps(a['eq_chart_values']) + ',backgroundColor:' + _json.dumps(eq_colors) + '}]')
    p.append('},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}}}});')

    # Daily Chart
    p.append('const dailyCtx=document.getElementById("dailyChart").getContext("2d");')
    p.append('new Chart(dailyCtx,{type:"line",data:{')
    p.append('labels:' + _json.dumps(a['daily_labels'], ensure_ascii=False) + ',')
    p.append('datasets:[{label:"Falhas/Dia",data:' + _json.dumps(a['daily_values']) + ',borderColor:"#2a5298",backgroundColor:"rgba(42,82,152,0.1)",fill:true,tension:0.3}]')
    p.append('},options:{responsive:true,maintainAspectRatio:false}});')

    # DF per Equipment Chart
    df_labels = [str(e['equipment']) for e in a['all_eq_df']]
    df_vals = [e['df'] for e in a['all_eq_df']]
    df_colors2 = ['#e74c3c' if v < 85 else '#f39c12' if v < df_meta else '#27ae60' for v in df_vals]
    p.append('const dfCtx=document.getElementById("dfChart").getContext("2d");')
    p.append('new Chart(dfCtx,{type:"bar",data:{')
    p.append('labels:' + _json.dumps(df_labels, ensure_ascii=False) + ',')
    p.append('datasets:[{label:"DF %",data:' + _json.dumps(df_vals) + ',backgroundColor:' + _json.dumps(df_colors2) + '}]')
    p.append('},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{min:0,max:100}}}});')

    p.append('</script>')

    # PDF Download Script
    pdf_filename = 'Relatorio_Confiabilidade_' + datetime.now().strftime('%Y%m%d') + '.pdf'
    p.append('<script>')
    p.append('function downloadPDF(){')
    p.append('var btn=document.querySelector(".btn-pdf");')
    p.append('btn.textContent="Gerando PDF...";btn.disabled=true;')
    p.append('var el=document.getElementById("report-content");')
    p.append('var opt={margin:[10,10,10,10],')
    p.append('filename:"' + pdf_filename + '",')
    p.append('image:{type:"jpeg",quality:0.98},')
    p.append('html2canvas:{scale:2,useCORS:true,logging:false},')
    p.append('jsPDF:{unit:"mm",format:"a4",orientation:"portrait"},')
    p.append('pagebreak:{mode:["avoid-all","css","legacy"]}};')
    p.append('html2pdf().set(opt).from(el).save().then(function(){')
    p.append('btn.textContent="Baixar PDF";btn.disabled=false;});')
    p.append('}</script>')

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
    app.run(debug=True, host='0.0.0.0', port=5000)
