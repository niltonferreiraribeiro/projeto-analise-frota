from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.utils import get_column_letter
import os
from datetime import datetime
import traceback
from collections import defaultdict

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max
app.config['UPLOAD_FOLDER'] = 'uploads'

# Create uploads folder if it doesn't exist
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

ALLOWED_EXTENSIONS = {'xlsx', 'xlsm'}

def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def parse_time(time_obj):
    """Convert Excel time object to seconds"""
    if time_obj is None:
        return 0
    if isinstance(time_obj, str):
        try:
            parts = time_obj.split(':')
            hours = int(parts[0]) if len(parts) > 0 else 0
            minutes = int(parts[1]) if len(parts) > 1 else 0
            seconds = int(parts[2]) if len(parts) > 2 else 0
            return hours * 3600 + minutes * 60 + seconds
        except:
            return 0
    # Excel time as decimal day fraction or time object
    try:
        if hasattr(time_obj, 'hour'):
            return time_obj.hour * 3600 + time_obj.minute * 60 + time_obj.second
        return 0
    except:
        return 0

def seconds_to_hms(seconds):
    """Convert seconds to HH:MM:SS format"""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    return "{:02d}:{:02d}:{:02d}".format(hours, minutes, secs)

def parse_df_meta(df_meta_str):
    """Parse DF meta value, handling comma as decimal separator"""
    try:
        if isinstance(df_meta_str, str):
            df_meta_str = df_meta_str.replace(',', '.')
        return float(df_meta_str)
    except:
        return 90.50

def convert_to_serializable(obj):
    """Convert datetime/time objects to strings for JSON serialization"""
    if hasattr(obj, 'strftime'):  # datetime or date object
        return obj.strftime('%Y-%m-%d %H:%M:%S')
    elif hasattr(obj, 'hour'):  # time object
        return obj.strftime('%H:%M:%S')
    return str(obj) if obj is not None else ""

def normalize_system_name(system_name):
    """Normalize system names (fix capitalization issues)"""
    if not system_name:
        return system_name
    # Standardize common variations
    system_map = {
        'eletroeletrônico': 'Eletroeletrônico',
        'eletroeletronico': 'Eletroeletrônico',
        'hidráulico': 'Hidráulico',
    }
    lower_name = system_name.lower()
    return system_map.get(lower_name, system_name)

def process_excel(filepath, df_meta):
    """
    Process Excel file and generate comprehensive failure analysis.
    
    Expected structure:
    - Sheet "DF": Equipment DF data, accumulated DF
    - Sheet "Escavadeira": Failure records with skip handling for empty rows
    """
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)

        # Check required sheets
        if 'Escavadeira' not in wb.sheetnames:
            return None, "Aba 'Escavadeira' não encontrada"
        if 'DF' not in wb.sheetnames:
            return None, "Aba 'DF' não encontrada"

        ws_failures = wb['Escavadeira']
        ws_df = wb['DF']

        # ===== PROCESS FAILURE DATA =====
        failures = []
        equipment_failures = {}
        equipment_details = {}
        system_failures = {}
        subsystem_failures = {}
        fault_descriptions = {}
        daily_failures = defaultdict(int)

        row = 2
        consecutive_empty = 0
        MAX_EMPTY_ROWS = 5

        while consecutive_empty < MAX_EMPTY_ROWS:
            equipment = ws_failures[f'B{row}'].value

            # Count consecutive empty rows in equipment column
            if not equipment:
                consecutive_empty += 1
                row += 1
                continue

            # Reset counter when we find data
            consecutive_empty = 0

            # Filter only equipment starting with "94"
            if not str(equipment).startswith('94'):
                row += 1
                continue

            date_val = ws_failures[f'A{row}'].value
            description = ws_failures[f'C{row}'].value or ""
            system = ws_failures[f'D{row}'].value or ""
            subsystem = ws_failures[f'E{row}'].value or ""
            start_time = ws_failures[f'F{row}'].value
            end_time = ws_failures[f'G{row}'].value
            duration = ws_failures[f'H{row}'].value
            failure_description = ws_failures[f'I{row}'].value or ""

            # Calculate duration in seconds
            duration_seconds = parse_time(duration)

            # Store failure record (convert datetime/time to strings)
            failure_record = {
                'date': convert_to_serializable(date_val),
                'equipment': str(equipment),
                'description': description,
                'system': normalize_system_name(system),
                'subsystem': subsystem,
                'start_time': convert_to_serializable(start_time),
                'end_time': convert_to_serializable(end_time),
                'duration': seconds_to_hms(duration_seconds),
                'duration_seconds': duration_seconds,
                'failure_description': failure_description
            }

            failures.append(failure_record)

            # Track daily failures
            if date_val:
                date_key = convert_to_serializable(date_val)[:10]
                daily_failures[date_key] += 1

            # Count by equipment
            if equipment not in equipment_failures:
                equipment_failures[equipment] = 0
                equipment_details[equipment] = []
            equipment_failures[equipment] += 1
            equipment_details[equipment].append(failure_record)

            # Count by system
            if system:
                norm_system = normalize_system_name(system)
                if norm_system not in system_failures:
                    system_failures[norm_system] = 0
                system_failures[norm_system] += 1

            # Count by subsystem
            if subsystem:
                if subsystem not in subsystem_failures:
                    subsystem_failures[subsystem] = 0
                subsystem_failures[subsystem] += 1

            # Track fault descriptions for recurrence analysis
            if failure_description:
                if failure_description not in fault_descriptions:
                    fault_descriptions[failure_description] = {'count': 0, 'equipments': set(), 'dates': []}
                fault_descriptions[failure_description]['count'] += 1
                fault_descriptions[failure_description]['equipments'].add(str(equipment))
                if date_val:
                    fault_descriptions[failure_description]['dates'].append(date_val)

            row += 1

        # ===== PROCESS DF DATA =====
        df_acumulado = 0
        escavadeira_df_data = []
        equipment_df_map = {}

        # Read Escavadeira DF data starting from row 3 (row 2 has headers)
        for row in range(3, 10):
            equipment_df = ws_df[f'F{row}'].value
            df_value = ws_df[f'G{row}'].value
            week = ws_df[f'H{row}'].value

            if equipment_df and df_value is not None:
                if isinstance(df_value, str):
                    df_pct = float(df_value.replace('%', '').replace(',', '.'))
                else:
                    df_pct = float(df_value)

                escavadeira_df_data.append({
                    'equipment': str(equipment_df),
                    'df': df_pct,
                    'week': week
                })
                equipment_df_map[str(equipment_df)] = df_pct

        # Read accumulated DF - search for "acumulado" in any row
        for row in range(10, 20):
            label = ws_df[f'F{row}'].value
            if label and "acumulado" in str(label).lower():
                accumulated_df = ws_df[f'G{row}'].value
                if accumulated_df is not None:
                    if isinstance(accumulated_df, str):
                        df_acumulado = float(accumulated_df.replace('%', '').replace(',', '.'))
                    else:
                        df_acumulado = float(accumulated_df)
                    break

        # ===== FAULT CLASSIFICATION =====
        classified_faults = {
            'RECORRENTE': [],
            'CRITICA': [],
            'POTENCIAL': [],
            'PONTUAL': []
        }

        fault_classifications = {}

        for failure in failures:
            fault_desc = failure['failure_description']
            equipment = failure['equipment']
            system = failure['system']

            classification = None

            # Check for RECORRENTE (same fault 2+ times)
            if fault_desc and fault_descriptions.get(fault_desc, {}).get('count', 0) >= 2:
                classification = 'RECORRENTE'

            # Check for CRITICA
            elif equipment_failures.get(equipment, 0) >= 3:
                classification = 'CRITICA'
            elif equipment in equipment_df_map and equipment_df_map[equipment] < 85:
                classification = 'CRITICA'

            # Check for POTENCIAL
            elif system and system_failures.get(system, 0) > 1:
                fault_count = fault_descriptions.get(fault_desc, {}).get('count', 0)
                if fault_count == 1:
                    classification = 'POTENCIAL'

            # Default to PONTUAL
            if not classification:
                classification = 'PONTUAL'

            fault_key = "{0}|{1}|{2}".format(equipment, fault_desc, failure['date'])
            fault_classifications[fault_key] = {
                'classification': classification,
                'failure': failure
            }

            classified_faults[classification].append({
                'equipment': equipment,
                'fault': fault_desc,
                'date': failure['date'],
                'system': system,
                'count': fault_descriptions.get(fault_desc, {}).get('count', 1)
            })

        # ===== PARETO ANALYSIS =====
        sorted_systems = sorted(system_failures.items(), key=lambda x: x[1], reverse=True)
        total_system_failures = sum(count for _, count in sorted_systems)

        pareto_systems = []
        cumulative = 0
        for system, count in sorted_systems:
            cumulative += count
            cumulative_pct = (cumulative / total_system_failures * 100) if total_system_failures > 0 else 0
            pareto_systems.append({
                'system': system,
                'count': count,
                'percentage': (count / total_system_failures * 100) if total_system_failures > 0 else 0,
                'cumulative': cumulative,
                'cumulative_pct': cumulative_pct
            })

        # ===== EQUIPMENT STATUS =====
        equipment_status = {}
        all_equipment_ids = set(list(equipment_failures.keys()) + [str(eq) for eq in equipment_df_map.keys()])

        for eq_id in all_equipment_ids:
            failure_count = equipment_failures.get(eq_id, 0)
            df_value = equipment_df_map.get(str(eq_id), equipment_df_map.get(eq_id, 0))

            if df_value == 0 and failure_count == 0:
                status = 'PREVENTIVA'
            elif df_value < 50 and df_value > 0:
                status = 'PARADO'
            elif df_value < df_meta:
                status = 'ALERTA'
            elif failure_count > 0:
                status = 'OPERACIONAL'
            else:
                status = 'OPERACIONAL'

            equipment_status[str(eq_id)] = {
                'status': status,
                'df': df_value,
                'failures': failure_count
            }

        # ===== DEVELOPING PATTERNS =====
        developing_patterns = []
        for fault_desc, data in fault_descriptions.items():
            if data['count'] == 2:
                developing_patterns.append({
                    'fault': fault_desc,
                    'count': data['count'],
                    'equipments': ', '.join(sorted(data['equipments']))
                })

        # ===== SUGGESTED PROJECTS =====
        suggested_projects = []
        recurrent_faults = {desc: data for desc, data in fault_descriptions.items() if data['count'] >= 2}

        for fault_desc, data in sorted(recurrent_faults.items(), key=lambda x: x[1]['count'], reverse=True)[:3]:
            suggested_projects.append({
                'title': 'Eliminar falha recorrente: {0}'.format(fault_desc[:60]),
                'description': 'Ocorreu {0}x em {1} equipamento(s)'.format(data['count'], len(data['equipments'])),
                'equipments': ', '.join(sorted(data['equipments'])),
                'priority': 'ALTA' if data['count'] >= 3 else 'MEDIA'
            })

        # ===== TOP RANKINGS =====
        top_systems = sorted(system_failures.items(), key=lambda x: x[1], reverse=True)[:3]
        top_subsystems = sorted(subsystem_failures.items(), key=lambda x: x[1], reverse=True)[:3]
        top_equipment = sorted(equipment_failures.items(), key=lambda x: x[1], reverse=True)[:3]

        # ===== CRITICAL DAYS (DF < 85%) =====
        critical_days = []
        # For now, we only have weekly DF data, not daily
        # This will be populated when daily DF data is available

        # ===== GENERATE ALERTS =====
        alerts = []

        if df_acumulado < df_meta:
            severity = "CRÍTICO" if df_acumulado < 50 else "ALTO"
            alerts.append({
                'severity': severity,
                'title': 'Disponibilidade Abaixo da Meta',
                'message': 'DF: {0:.2f}% | Meta: {1:.2f}%'.format(df_acumulado, df_meta),
                'action': 'Análise urgente de causas raízes e plano de ação'
            })

        for equipment, count in top_equipment:
            if count >= 3:
                alerts.append({
                    'severity': 'ALTO',
                    'title': 'Equipamento Crítico: {0}'.format(equipment),
                    'message': '{0} falhas registradas'.format(count),
                    'action': 'Manutenção preventiva recomendada'
                })

        for fault_desc, data in recurrent_faults.items():
            if data['count'] >= 2:
                alerts.append({
                    'severity': 'MEDIO',
                    'title': 'Falha Recorrente',
                    'message': '{0}... ({1}x)'.format(fault_desc[:60], data['count']),
                    'action': 'Investigação técnica necessária'
                })

        # ===== CRITICAL FAILURES =====
        critical_failures = []
        for f in failures:
            is_critical = False
            eq = f['equipment']
            # Critical if equipment has 3+ failures or DF < 85%
            if equipment_failures.get(eq, 0) >= 3 or equipment_df_map.get(eq, 0) < 85:
                is_critical = True
            
            if is_critical:
                critical_failures.append({
                    'date': f['date'],
                    'equipment': f['equipment'],
                    'system': f['system'],
                    'fault': f['failure_description'],
                    'duration': f['duration'],
                    'duration_seconds': f['duration_seconds']
                })

        # ===== RECURRENT PATTERNS =====
        recurrent_patterns = []
        for fault_desc, data in fault_descriptions.items():
            if data['count'] >= 2:
                recurrent_patterns.append({
                    'fault': fault_desc,
                    'count': data['count'],
                    'equipments': list(data['equipments']),
                    'dates': [convert_to_serializable(d) for d in data['dates'][:5]]  # Last 5 dates
                })

        # ===== COMPILE ANALYSIS DATA =====
        analysis = {
            'total_failures': len(failures),
            'df_acumulado': df_acumulado,
            'df_meta': df_meta,
            'df_status': 'CRÍTICO' if df_acumulado < 50 else 'ALERTA' if df_acumulado < df_meta else 'OK',
            'equipments': equipment_failures,
            'systems': system_failures,
            'subsystems': subsystem_failures,
            'top_systems': top_systems,
            'top_subsystems': top_subsystems,
            'top_equipment': top_equipment,
            'alerts': alerts,
            'equipment_details': equipment_details,
            'equipment_df_map': equipment_df_map,
            'escavadeira_df': escavadeira_df_data,
            'equipment_status': equipment_status,
            'classified_faults': classified_faults,
            'pareto_systems': pareto_systems,
            'developing_patterns': developing_patterns,
            'suggested_projects': suggested_projects,
            'daily_failures': dict(daily_failures),
            'critical_failures': critical_failures,
            'recurrent_patterns': recurrent_patterns,
            'critical_days': critical_days,
            'generation_time': datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        }

        wb.close()
        return analysis, None

    except Exception as e:
        return None, "Erro ao processar arquivo: {0}".format(str(e))

@app.route('/')
def index():
    """Serve the main upload page"""
    return render_template('index.html')

@app.route('/api/upload', methods=['POST'])
def upload():
    """Handle file upload and analysis"""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'Nenhum arquivo enviado'}), 400

        file = request.files['file']

        if file.filename == '':
            return jsonify({'success': False, 'error': 'Nenhum arquivo selecionado'}), 400

        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': 'Apenas arquivos .xlsx são permitidos'}), 400

        df_meta = request.form.get('df_meta', '90.50')
        df_meta = parse_df_meta(df_meta)

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

        return jsonify({
            'success': True,
            'report': html_report,
            'analysis': analysis
        }), 200

    except Exception as e:
        return jsonify({'success': False, 'error': 'Erro no servidor: {0}'.format(str(e))}), 500

def generate_report(analysis):
    """Generate professional HTML report with strategic analysis"""
    
    parts = []
    
    # Extract data
    total_failures = analysis['total_failures']
    df_acumulado = analysis['df_acumulado']
    df_meta = analysis['df_meta']
    df_status = analysis['df_status']
    alerts = analysis['alerts']
    equipment_status = analysis['equipment_status']
    pareto_systems = analysis['pareto_systems']
    classified_faults = analysis['classified_faults']
    developing_patterns = analysis['developing_patterns']
    suggested_projects = analysis['suggested_projects']
    generation_time = analysis['generation_time']
    top_systems = analysis['top_systems']
    top_equipment = analysis['top_equipment']
    equipment_df_map = analysis['equipment_df_map']
    equipments = analysis['equipments']
    critical_failures = analysis['critical_failures']
    recurrent_patterns = analysis['recurrent_patterns']
    daily_failures = analysis['daily_failures']

    # Build equipment details for table
    equipment_details = [
        {
            'equipment': eq_id,
            'failures': data['failures'],
            'status': data['status'],
            'df': data['df']
        }
        for eq_id, data in sorted(equipment_status.items())
    ]

    # HTML Header and CSS
    parts.append('<!DOCTYPE html>')
    parts.append('<html lang="pt-BR">')
    parts.append('<head>')
    parts.append('    <meta charset="UTF-8">')
    parts.append('    <meta name="viewport" content="width=device-width, initial-scale=1.0">')
    parts.append('    <title>Relatório de Confiabilidade - Frota de Escavadeiras</title>')
    parts.append('    <style>')
    parts.append('        * { margin: 0; padding: 0; box-sizing: border-box; }')
    parts.append('        body { font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif; background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%); color: #333; line-height: 1.6; padding: 20px; }')
    parts.append('        .container { max-width: 1400px; margin: 0 auto; }')
    parts.append('        .header { background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); color: white; padding: 40px; border-radius: 10px; margin-bottom: 30px; box-shadow: 0 8px 16px rgba(0,0,0,0.2); }')
    parts.append('        .header h1 { font-size: 32px; margin-bottom: 10px; }')
    parts.append('        .header p { font-size: 14px; opacity: 0.9; }')
    parts.append('        .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }')
    parts.append('        .kpi-card { background: white; padding: 25px; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); border-left: 5px solid #2a5298; }')
    parts.append('        .kpi-card.critical { border-left-color: #e74c3c; }')
    parts.append('        .kpi-card.warning { border-left-color: #f39c12; }')
    parts.append('        .kpi-card.success { border-left-color: #27ae60; }')
    parts.append('        .kpi-value { font-size: 28px; font-weight: bold; color: #1e3c72; margin-bottom: 5px; }')
    parts.append('        .kpi-card.critical .kpi-value { color: #e74c3c; }')
    parts.append('        .kpi-card.warning .kpi-value { color: #f39c12; }')
    parts.append('        .kpi-card.success .kpi-value { color: #27ae60; }')
    parts.append('        .kpi-label { font-size: 12px; color: #999; text-transform: uppercase; letter-spacing: 1px; font-weight: 600; }')
    parts.append('        .section { background: white; padding: 30px; border-radius: 10px; margin-bottom: 30px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }')
    parts.append('        .section h2 { font-size: 22px; color: #1e3c72; margin-bottom: 20px; border-bottom: 3px solid #2a5298; padding-bottom: 10px; }')
    parts.append('        .section h3 { font-size: 16px; color: #2a5298; margin-top: 20px; margin-bottom: 15px; }')
    parts.append('        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }')
    parts.append('        table th { background: linear-gradient(135deg, #2a5298 0%, #1e3c72 100%); color: white; padding: 15px; text-align: left; font-weight: 600; }')
    parts.append('        table td { padding: 12px 15px; border-bottom: 1px solid #e0e0e0; }')
    parts.append('        table tbody tr:hover { background-color: #f8f9fa; }')
    parts.append('        table tbody tr:nth-child(even) { background-color: #f9f9f9; }')
    parts.append('        .badge { display: inline-block; padding: 6px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; }')
    parts.append('        .badge.recorrente { background-color: #e74c3c; color: white; }')
    parts.append('        .badge.critica { background-color: #c0392b; color: white; }')
    parts.append('        .badge.potencial { background-color: #f39c12; color: white; }')
    parts.append('        .badge.pontual { background-color: #3498db; color: white; }')
    parts.append('        .badge.operacional { background-color: #27ae60; color: white; }')
    parts.append('        .badge.preventiva { background-color: #2980b9; color: white; }')
    parts.append('        .badge.parado { background-color: #c0392b; color: white; }')
    parts.append('        .badge.alerta { background-color: #f39c12; color: white; }')
    parts.append('        .alert { padding: 15px; margin-bottom: 15px; border-radius: 5px; border-left: 4px solid; }')
    parts.append('        .alert.critical { background-color: #fadbd8; border-left-color: #e74c3c; color: #a93226; }')
    parts.append('        .alert.high { background-color: #fef5e7; border-left-color: #f39c12; color: #9a6c06; }')
    parts.append('        .alert.medium { background-color: #d6eaf8; border-left-color: #3498db; color: #1b4f72; }')
    parts.append('        .alert-title { font-weight: bold; font-size: 14px; margin-bottom: 5px; }')
    parts.append('        .alert-message { font-size: 13px; margin-bottom: 5px; }')
    parts.append('        .pareto-bar { display: flex; gap: 10px; margin-bottom: 15px; align-items: center; }')
    parts.append('        .pareto-system { flex: 0 0 150px; font-weight: 600; font-size: 12px; }')
    parts.append('        .pareto-visual { flex: 1; height: 30px; background: linear-gradient(90deg, #3498db 0%, #2980b9 100%); border-radius: 3px; display: flex; align-items: center; padding: 0 10px; color: white; font-weight: 600; font-size: 11px; }')
    parts.append('        .pareto-stats { flex: 0 0 200px; text-align: right; font-size: 11px; }')
    parts.append('        .no-data { text-align: center; padding: 20px; color: #999; font-style: italic; }')
    parts.append('        .action-item { background-color: #f0f8ff; padding: 15px; border-left: 4px solid #3498db; margin-bottom: 15px; border-radius: 3px; }')
    parts.append('        .action-critical { border-left-color: #e74c3c; background-color: #fadbd8; }')
    parts.append('        .action-high { border-left-color: #f39c12; background-color: #fef5e7; }')
    parts.append('        .action-medium { border-left-color: #3498db; background-color: #d6eaf8; }')
    parts.append('        .action-title { font-weight: 600; color: #1b4f72; margin-bottom: 5px; }')
    parts.append('        .action-desc { font-size: 13px; color: #34495e; margin-bottom: 8px; }')
    parts.append('        .footer { text-align: center; padding: 20px; color: #7f8c8d; font-size: 12px; border-top: 1px solid #ddd; margin-top: 40px; }')
    parts.append('        @media (max-width: 768px) { .header h1 { font-size: 24px; } .kpi-grid { grid-template-columns: 1fr; } table { font-size: 12px; } }')
    parts.append('        @media print { body { background: white; } .section { page-break-inside: avoid; } }')
    parts.append('    </style>')
    parts.append('</head>')
    parts.append('<body>')
    parts.append('    <div class="container">')

    # === HEADER ===
    parts.append('        <div class="header">')
    parts.append('            <h1>📊 Relatório de Confiabilidade - Frota de Escavadeiras CAT</h1>')
    parts.append('            <p>Análise de Falhas e Disponibilidade Física (DF)</p>')
    parts.append('            <p>Data/Hora: {0}</p>'.format(generation_time))
    parts.append('        </div>')

    # === KPI CARDS ===
    parts.append('        <div class="kpi-grid">')
    
    # DF Card
    kpi_class = 'success' if df_status == 'OK' else 'warning' if df_status == 'ALERTA' else 'critical'
    parts.append('            <div class="kpi-card {0}">'.format(kpi_class))
    parts.append('                <div class="kpi-value">{0:.2f}%</div>'.format(df_acumulado))
    parts.append('                <div class="kpi-label">Disponibilidade (DF)</div>')
    parts.append('            </div>')
    
    # Meta Card
    parts.append('            <div class="kpi-card">')
    parts.append('                <div class="kpi-value">{0:.2f}%</div>'.format(df_meta))
    parts.append('                <div class="kpi-label">Meta de DF</div>')
    parts.append('            </div>')
    
    # Total Failures Card
    parts.append('            <div class="kpi-card">')
    parts.append('                <div class="kpi-value">{0}</div>'.format(total_failures))
    parts.append('                <div class="kpi-label">Total de Falhas</div>')
    parts.append('            </div>')
    
    # Difference Card
    diff = df_acumulado - df_meta
    diff_class = 'success' if diff >= 0 else 'critical'
    parts.append('            <div class="kpi-card {0}">'.format(diff_class))
    parts.append('                <div class="kpi-value">{0:+.2f}%</div>'.format(diff))
    parts.append('                <div class="kpi-label">Diferença vs Meta</div>')
    parts.append('            </div>')
    
    parts.append('        </div>')

    # === EXECUTIVE SUMMARY ===
    parts.append('        <div class="section">')
    parts.append('            <h2>📋 Resumo Executivo</h2>')
    
    if df_status == 'OK':
        status_text = '✅ <strong>OK</strong> - Frota acima da meta'
    elif df_status == 'ALERTA':
        status_text = '⚠️ <strong>ALERTA</strong> - Frota abaixo da meta'
    else:
        status_text = '🚨 <strong>CRÍTICO</strong> - Situação crítica'
    
    parts.append('            <p><strong>Status Geral:</strong> {0}</p>'.format(status_text))
    parts.append('            <p>A frota apresenta DF de <strong>{0:.2f}%</strong> contra meta de <strong>{1:.2f}%</strong>, '
                 'uma diferença de <strong>{2:+.2f}%</strong>.</p>'.format(df_acumulado, df_meta, diff))
    
    # Top concerns
    parts.append('            <p><strong>Principais Preocupações:</strong></p>')
    parts.append('            <ul>')
    
    if len(top_equipment) > 0:
        eq_name, eq_count = top_equipment[0]
        eq_df = equipment_df_map.get(str(eq_name), 0)
        parts.append('                <li>Equipamento {0} com {1} falhas (DF={2:.2f}%)</li>'.format(eq_name, eq_count, eq_df))
    
    if len(top_systems) > 0:
        sys_name, sys_count = top_systems[0]
        parts.append('                <li>Sistema {0} com {1} falhas ({2:.1f}% do total)</li>'.format(
            sys_name, sys_count, (sys_count / total_failures * 100) if total_failures > 0 else 0))
    
    if len(recurrent_patterns) > 0:
        pattern = recurrent_patterns[0]
        parts.append('                <li>Padrão recorrente: "{0}" ({1}x)</li>'.format(pattern['fault'][:50], pattern['count']))
    
    parts.append('            </ul>')
    parts.append('        </div>')

    # === AVAILABILITY ANALYSIS ===
    parts.append('        <div class="section">')
    parts.append('            <h2>📈 Análise de Disponibilidade Física (DF)</h2>')
    parts.append('            <h3>Comparativo por Equipamento</h3>')
    parts.append('            <table>')
    parts.append('                <tr><th>Equipamento</th><th>DF (%)</th><th>Status</th><th>Falhas</th></tr>')
    
    for data in equipment_details:
        eq_id = data['equipment']
        df_val = data['df']
        status = data['status']
        failures = data['failures']
        
        if df_val == 0 and failures == 0:
            status_badge = '<span class="badge preventiva">{0}</span>'.format(status)
        elif df_val < 85:
            status_badge = '<span class="badge parado">{0}</span>'.format(status)
        elif df_val < df_meta:
            status_badge = '<span class="badge alerta">{0}</span>'.format(status)
        else:
            status_badge = '<span class="badge operacional">{0}</span>'.format(status)
        
        parts.append('                <tr><td>{0}</td><td>{1:.2f}%</td><td>{2}</td><td>{3}</td></tr>'.format(
            eq_id, df_val, status_badge, failures))
    
    parts.append('            </table>')
    parts.append('        </div>')

    # === CRITICAL FAILURES ===
    if len(critical_failures) > 0:
        parts.append('        <div class="section">')
        parts.append('            <h2>🚨 Falhas Críticas Identificadas</h2>')
        parts.append('            <table>')
        parts.append('                <tr><th>Data</th><th>Equipamento</th><th>Sistema</th><th>Descrição</th><th>Duração</th></tr>')
        
        for cf in critical_failures[:10]:  # Show top 10
            parts.append('                <tr>')
            parts.append('                    <td>{0}</td>'.format(cf['date']))
            parts.append('                    <td>{0}</td>'.format(cf['equipment']))
            parts.append('                    <td>{0}</td>'.format(cf['system']))
            parts.append('                    <td>{0}</td>'.format(cf['fault'][:60]))
            parts.append('                    <td>{0}</td>'.format(cf['duration']))
            parts.append('                </tr>')
        
        parts.append('            </table>')
        parts.append('        </div>')

    # === FAILURE CLASSIFICATION ===
    parts.append('        <div class="section">')
    parts.append('            <h2>🏷️ Classificação de Falhas</h2>')
    parts.append('            <table>')
    parts.append('                <tr><th>Tipo</th><th>Quantidade</th><th>Percentual</th></tr>')
    
    for fault_type in ['RECORRENTE', 'CRITICA', 'POTENCIAL', 'PONTUAL']:
        count = len(classified_faults.get(fault_type, []))
        pct = (count / total_failures * 100) if total_failures > 0 else 0
        badge_class = fault_type.lower()
        parts.append('                <tr>')
        parts.append('                    <td><span class="badge {0}">{1}</span></td>'.format(badge_class, fault_type))
        parts.append('                    <td>{0}</td>'.format(count))
        parts.append('                    <td>{0:.1f}%</td>'.format(pct))
        parts.append('                </tr>')
    
    parts.append('            </table>')
    parts.append('        </div>')

    # === PARETO ANALYSIS ===
    parts.append('        <div class="section">')
    parts.append('            <h2>📊 Análise de Pareto (Sistemas)</h2>')
    parts.append('            <p>Os 20% de sistemas que causam aproximadamente 80% das falhas:</p>')
    
    for s in pareto_systems[:5]:
        bar_width = max(30, min(500, s['percentage'] * 10))
        parts.append('            <div class="pareto-bar">')
        parts.append('                <div class="pareto-system">{0}</div>'.format(s['system']))
        parts.append('                <div class="pareto-visual" style="width: {0}px;">'.format(int(bar_width)))
        parts.append('                    {0} falhas'.format(s['count']))
        parts.append('                </div>')
        parts.append('                <div class="pareto-stats">{0:.1f}% | Acum: {1:.1f}%</div>'.format(
            s['percentage'], s['cumulative_pct']))
        parts.append('            </div>')
    
    parts.append('        </div>')

    # === RECURRENT PATTERNS ===
    if len(recurrent_patterns) > 0:
        parts.append('        <div class="section">')
        parts.append('            <h2>🔄 Padrões Recorrentes</h2>')
        
        for pattern in recurrent_patterns[:5]:
            parts.append('            <div style="background: #fff8e1; padding: 12px; border-left: 4px solid #f39c12; margin-bottom: 10px; border-radius: 3px;">')
            parts.append('                <div style="font-weight: 600; color: #333; margin-bottom: 5px;">{0}</div>'.format(pattern['fault']))
            parts.append('                <div style="font-size: 12px; color: #666;">Ocorrências: {0} | Equipamentos: {1}</div>'.format(
                pattern['count'], ', '.join(pattern['equipments'])))
            parts.append('            </div>')
        
        parts.append('        </div>')

    # === RECOMMENDED ACTIONS ===
    parts.append('        <div class="section">')
    parts.append('            <h2>✅ Ações Recomendadas</h2>')
    
    # Critical actions
    critical_actions = []
    for eq, failures_count in top_equipment:
        eq_df = equipment_df_map.get(str(eq), 0)
        if failures_count >= 3 or eq_df < 85:
            critical_actions.append((eq, failures_count, eq_df))
    
    if len(critical_actions) > 0:
        parts.append('            <h3>🚨 CRÍTICAS (Executar HOJE)</h3>')
        for eq, fcount, df_val in critical_actions[:3]:
            parts.append('            <div class="action-item action-critical">')
            parts.append('                <div class="action-title">Equipamento {0} - DF {1:.2f}%</div>'.format(eq, df_val))
            parts.append('                <div class="action-desc">{0} falhas registradas. Requer análise urgente e manutenção preventiva.</div>'.format(fcount))
            parts.append('            </div>')
    
    # High priority actions
    if len(top_systems) > 0:
        parts.append('            <h3>🔴 ALTAS (Esta Semana)</h3>')
        for sys, count in top_systems[:2]:
            parts.append('            <div class="action-item action-high">')
            parts.append('                <div class="action-title">Sistema: {0}</div>'.format(sys))
            parts.append('                <div class="action-desc">{0} falhas. Executar manutenção preventiva no sistema.</div>'.format(count))
            parts.append('            </div>')
    
    # Medium priority actions
    if len(developing_patterns) > 0:
        parts.append('            <h3>🟡 MÉDIAS (Próximas 2 Semanas)</h3>')
        for pattern in developing_patterns[:2]:
            parts.append('            <div class="action-item action-medium">')
            parts.append('                <div class="action-title">Padrão Emergente: {0}</div>'.format(pattern['fault'][:60]))
            parts.append('                <div class="action-desc">Ocorreu {0}x. Acompanhar evolução.</div>'.format(pattern['count']))
            parts.append('            </div>')
    
    parts.append('        </div>')

    # === CONCLUSION ===
    parts.append('        <div class="section">')
    parts.append('            <h2>🎯 Conclusão</h2>')
    
    if df_status == 'OK':
        conclusion = 'A frota está acima da meta de disponibilidade. Continue com as práticas atuais de manutenção preventiva.'
    else:
        conclusion = 'A frota está abaixo da meta de disponibilidade. Ações imediatas nos equipamentos críticos são necessárias.'
    
    parts.append('            <p>{0}</p>'.format(conclusion))
    parts.append('            <p><strong>Próximas Análises:</strong> Diariamente pela manhã, atualizando DF do dia anterior e analisando novas falhas.</p>')
    parts.append('        </div>')

    # === FOOTER ===
    parts.append('        <div class="footer">')
    parts.append('            <p>Relatório gerado automaticamente pelo Sistema de Análise de Falhas e Confiabilidade da Frota</p>')
    parts.append('            <p>Data/Hora: {0}</p>'.format(generation_time))
    parts.append('        </div>')

    parts.append('    </div>')
    parts.append('</body>')
    parts.append('</html>')

    return ''.join(parts)

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
