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
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"

def parse_df_meta(df_meta_str):
    """Parse DF meta value, handling comma as decimal separator"""
    try:
        if isinstance(df_meta_str, str):
            # Replace comma with dot for decimal
            df_meta_str = df_meta_str.replace(',', '.')
        return float(df_meta_str)
    except:
        return 90.50

def process_excel(filepath, df_meta):
    """
    Process Excel file and generate comprehensive failure analysis.

    Expected structure:
    - Sheet "DF": Equipment DF data (rows 3-9), accumulated DF (row 11)
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

        # =====================================================
        # PROCESS FAILURE DATA (Escavadeira sheet)
        # =====================================================
        failures = []
        equipment_failures = {}  # Count by equipment
        equipment_details = {}   # Store all failures per equipment
        system_failures = {}     # Count by system
        subsystem_failures = {}  # Count by subsystem
        fault_descriptions = {}  # Track all fault descriptions (for recurrence)

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

            # Store failure record
            failure_record = {
                'date': date_val,
                'equipment': str(equipment),
                'description': description,
                'system': system,
                'subsystem': subsystem,
                'start_time': start_time,
                'end_time': end_time,
                'duration': seconds_to_hms(duration_seconds),
                'duration_seconds': duration_seconds,
                'failure_description': failure_description
            }

            failures.append(failure_record)

            # Count by equipment
            if equipment not in equipment_failures:
                equipment_failures[equipment] = 0
                equipment_details[equipment] = []
            equipment_failures[equipment] += 1
            equipment_details[equipment].append(failure_record)

            # Count by system
            if system:
                if system not in system_failures:
                    system_failures[system] = 0
                system_failures[system] += 1

            # Count by subsystem
            if subsystem:
                if subsystem not in subsystem_failures:
                    subsystem_failures[subsystem] = 0
                subsystem_failures[subsystem] += 1

            # Track fault descriptions for recurrence analysis
            if failure_description:
                if failure_description not in fault_descriptions:
                    fault_descriptions[failure_description] = {'count': 0, 'equipments': set()}
                fault_descriptions[failure_description]['count'] += 1
                fault_descriptions[failure_description]['equipments'].add(str(equipment))

            row += 1

        # =====================================================
        # PROCESS DF DATA (DF sheet)
        # =====================================================
        df_acumulado = 0
        escavadeira_df_data = []
        equipment_df_map = {}  # Map equipment ID to DF value

        # Read Escavadeira DF data starting from row 3 (row 2 has headers)
        for row in range(3, 10):  # Rows 3-9 for individual equipment
            equipment_df = ws_df[f'F{row}'].value
            df_value = ws_df[f'G{row}'].value
            week = ws_df[f'H{row}'].value

            if equipment_df and df_value is not None:
                # Parse DF value (already percentages, not decimals)
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

        # Read accumulated DF from row 11
        accumulated_label = ws_df[f'F11'].value
        accumulated_df = ws_df[f'G11'].value

        if accumulated_label == "Acumulado da frota" and accumulated_df is not None:
            if isinstance(accumulated_df, str):
                df_acumulado = float(accumulated_df.replace('%', '').replace(',', '.'))
            else:
                df_acumulado = float(accumulated_df)

        # =====================================================
        # FAULT CLASSIFICATION
        # =====================================================
        classified_faults = {
            'RECORRENTE': [],      # 🔴 Same fault 2+ times
            'CRITICA': [],         # 🔴 Equipment with 3+ failures OR DF < 50%
            'POTENCIAL': [],       # 🟡 First occurrence in system with other failures
            'PONTUAL': []          # 🔵 One-time fault, no pattern
        }

        # Track which failures have been classified
        fault_classifications = {}

        for failure in failures:
            fault_desc = failure['failure_description']
            equipment = failure['equipment']
            system = failure['system']

            classification = None

            # Check for RECORRENTE (same fault 2+ times)
            if fault_desc and fault_descriptions.get(fault_desc, {}).get('count', 0) >= 2:
                classification = 'RECORRENTE'

            # Check for CRITICA (equipment with 3+ failures OR DF < 50%)
            elif equipment_failures.get(equipment, 0) >= 3:
                classification = 'CRITICA'
            elif equipment in equipment_df_map and equipment_df_map[equipment] < 50:
                classification = 'CRITICA'

            # Check for POTENCIAL (first occurrence in system with other failures)
            elif system and system_failures.get(system, 0) > 1:
                # Count how many failures of this description exist
                fault_count = fault_descriptions.get(fault_desc, {}).get('count', 0)
                if fault_count == 1:  # First occurrence
                    classification = 'POTENCIAL'

            # Default to PONTUAL
            if not classification:
                classification = 'PONTUAL'

            fault_key = f"{equipment}|{fault_desc}|{failure['date']}"
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

        # =====================================================
        # PARETO ANALYSIS (20% of systems cause 80% of failures)
        # =====================================================
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

        # =====================================================
        # EQUIPMENT STATUS (include all equipment from DF + failures)
        # =====================================================
        equipment_status = {}

        # Merge all equipment IDs from both DF sheet and failure records
        all_equipment_ids = set(list(equipment_failures.keys()) + [str(eq) for eq in equipment_df_map.keys()])

        for eq_id in all_equipment_ids:
            # Try both string and int keys
            failure_count = equipment_failures.get(eq_id, 0)
            if failure_count == 0:
                failure_count = equipment_failures.get(int(eq_id) if str(eq_id).isdigit() else eq_id, 0)
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

        # =====================================================
        # DEVELOP PATTERNS (emerging failures)
        # =====================================================
        developing_patterns = []
        for fault_desc, data in fault_descriptions.items():
            if data['count'] == 2:  # Just starting to repeat
                developing_patterns.append({
                    'fault': fault_desc,
                    'count': data['count'],
                    'equipments': ', '.join(sorted(data['equipments']))
                })

        # =====================================================
        # SUGGESTED PROJECTS (based on recurrent failures)
        # =====================================================
        suggested_projects = []
        recurrent_faults = {desc: data for desc, data in fault_descriptions.items() if data['count'] >= 2}

        for fault_desc, data in sorted(recurrent_faults.items(), key=lambda x: x[1]['count'], reverse=True)[:3]:
            suggested_projects.append({
                'title': f"Eliminar falha recorrente: {fault_desc[:60]}",
                'description': f"Ocorreu {data['count']}x em {len(data['equipments'])} equipamento(s)",
                'equipments': ', '.join(sorted(data['equipments'])),
                'priority': 'ALTA' if data['count'] >= 3 else 'MÉDIA'
            })

        # =====================================================
        # TOP RANKINGS
        # =====================================================
        top_systems = sorted(system_failures.items(), key=lambda x: x[1], reverse=True)[:3]
        top_subsystems = sorted(subsystem_failures.items(), key=lambda x: x[1], reverse=True)[:3]
        top_equipment = sorted(equipment_failures.items(), key=lambda x: x[1], reverse=True)[:3]

        # =====================================================
        # GENERATE ALERTS
        # =====================================================
        alerts = []

        # Alert: DF below target
        if df_acumulado < df_meta:
            severity = "CRÍTICO" if df_acumulado < 50 else "ALTO"
            alerts.append({
                'severity': severity,
                'title': 'Disponibilidade Abaixo da Meta',
                'message': f'DF: {df_acumulado:.2f}% | Meta: {df_meta:.2f}%',
                'action': 'Análise urgente de causas raízes e plano de ação'
            })

        # Alert: Critical equipment
        for equipment, count in top_equipment:
            if count >= 3:
                alerts.append({
                    'severity': 'ALTO',
                    'title': f'Equipamento Crítico: {equipment}',
                    'message': f'{count} falhas registradas',
                    'action': 'Manutenção preventiva recomendada'
                })

        # Alert: Recurrent failures
        for fault_desc, data in recurrent_faults.items():
            if data['count'] >= 2:
                alerts.append({
                    'severity': 'MÉDIO',
                    'title': 'Falha Recorrente',
                    'message': f'{fault_desc[:60]}... ({data["count"]}x)',
                    'action': 'Investigação técnica necessária'
                })

        # =====================================================
        # COMPILE ANALYSIS DATA
        # =====================================================
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
            'escavadeira_df': escavadeira_df_data,
            'equipment_status': equipment_status,
            'classified_faults': classified_faults,
            'pareto_systems': pareto_systems,
            'developing_patterns': developing_patterns,
            'suggested_projects': suggested_projects,
            'generation_time': datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        }

        wb.close()
        return analysis, None

    except Exception as e:
        return None, f"Erro ao processar arquivo: {str(e)}"

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

        # Parse DF meta (handle comma as decimal separator)
        df_meta = request.form.get('df_meta', '90.50')
        df_meta = parse_df_meta(df_meta)

        # Save file temporarily
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        # Process Excel
        analysis, error = process_excel(filepath, df_meta)

        # Delete file after processing
        try:
            os.remove(filepath)
        except:
            pass

        if error:
            return jsonify({'success': False, 'error': error}), 400

        # Generate HTML report
        html_report = generate_report(analysis)

        return jsonify({
            'success': True,
            'report': html_report,
            'analysis': analysis
        }), 200

    except Exception as e:
        return jsonify({'success': False, 'error': f"Erro no servidor: {str(e)}"}), 500

def generate_report(analysis):
    """Generate professional HTML report with all required sections"""

    # Prepare data for template
    df_acumulado = f"{analysis['df_acumulado']:.2f}%"
    df_meta = f"{analysis['df_meta']:.2f}%"
    df_diff = f"{analysis['df_acumulado'] - analysis['df_meta']:+.2f}%"
    df_status = analysis['df_status']
    generation_time = analysis['generation_time']
    total_failures = analysis['total_failures']

    # Equipment details (all equipment from DF + failures)
    equipment_details = [
        {
            'equipment': eq_id,
            'failures': data['failures'],
            'status': data['status'],
            'df': data['df']
        }
        for eq_id, data in sorted(analysis['equipment_status'].items())
    ]

    top_systems = [{'name': name, 'count': count} for name, count in analysis['top_systems']]
    top_equipment = [{'name': name, 'count': count} for name, count in analysis['top_equipment']]
    alerts = analysis['alerts']
    pareto_systems = analysis['pareto_systems']
    classified_faults = analysis['classified_faults']
    developing_patterns = analysis['developing_patterns']
    suggested_projects = analysis['suggested_projects']

    title = 'Análise de Falhas e Confiabilidade - Frota de Escavadeiras CAT'

    # Determine CSS classes based on status
    df_status_class = 'success' if df_status == 'OK' else 'warning' if df_status == 'ALERTA' else 'critical'
    df_diff_class = 'critical' if float(df_diff.replace('%', '')) < 0 else 'success'
    badge_status_class = 'success' if df_status == 'OK' else 'warning' if df_status == 'ALERTA' else 'critica'

    # Start building HTML
    parts = []

    # DOCTYPE and head
    parts.append('<!DOCTYPE html>')
    parts.append('<html lang="pt-BR">')
    parts.append('<head>')
    parts.append('    <meta charset="UTF-8">')
    parts.append('    <meta name="viewport" content="width=device-width, initial-scale=1.0">')
    parts.append(f'    <title>{title}</title>')
    parts.append('    <style>')

    # CSS styles (all unchanged)
    parts.append('''
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            color: #333;
            line-height: 1.6;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }

        .header {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            color: white;
            padding: 40px;
            border-radius: 10px;
            margin-bottom: 30px;
            box-shadow: 0 8px 16px rgba(0,0,0,0.2);
        }

        .header h1 {
            font-size: 32px;
            margin-bottom: 10px;
        }

        .header p {
            font-size: 14px;
            opacity: 0.9;
        }

        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }

        .kpi-card {
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            border-left: 5px solid #2a5298;
            transition: transform 0.3s, box-shadow 0.3s;
        }

        .kpi-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 20px rgba(0,0,0,0.15);
        }

        .kpi-card.critical {
            border-left-color: #e74c3c;
        }

        .kpi-card.warning {
            border-left-color: #f39c12;
        }

        .kpi-card.success {
            border-left-color: #27ae60;
        }

        .kpi-value {
            font-size: 28px;
            font-weight: bold;
            color: #1e3c72;
            margin-bottom: 5px;
        }

        .kpi-card.critical .kpi-value {
            color: #e74c3c;
        }

        .kpi-card.warning .kpi-value {
            color: #f39c12;
        }

        .kpi-card.success .kpi-value {
            color: #27ae60;
        }

        .kpi-label {
            font-size: 14px;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 1px;
            font-weight: 600;
        }

        .section {
            background: white;
            padding: 30px;
            border-radius: 10px;
            margin-bottom: 30px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }

        .section h2 {
            font-size: 22px;
            color: #1e3c72;
            margin-bottom: 20px;
            border-bottom: 3px solid #2a5298;
            padding-bottom: 10px;
        }

        .section h3 {
            font-size: 16px;
            color: #2a5298;
            margin-top: 20px;
            margin-bottom: 15px;
        }

        .alert {
            padding: 15px;
            margin-bottom: 15px;
            border-radius: 5px;
            border-left: 4px solid;
        }

        .alert.critical {
            background-color: #fadbd8;
            border-left-color: #e74c3c;
            color: #a93226;
        }

        .alert.high {
            background-color: #fef5e7;
            border-left-color: #f39c12;
            color: #9a6c06;
        }

        .alert.medium {
            background-color: #d6eaf8;
            border-left-color: #3498db;
            color: #1b4f72;
        }

        .alert-title {
            font-weight: bold;
            font-size: 14px;
            margin-bottom: 5px;
        }

        .alert-message {
            font-size: 13px;
            margin-bottom: 5px;
        }

        .alert-action {
            font-size: 12px;
            font-style: italic;
            margin-top: 8px;
            padding-top: 8px;
            border-top: 1px solid currentColor;
            opacity: 0.8;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }

        table th {
            background: linear-gradient(135deg, #2a5298 0%, #1e3c72 100%);
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
        }

        table td {
            padding: 12px 15px;
            border-bottom: 1px solid #e0e0e0;
        }

        table tbody tr:hover {
            background-color: #f8f9fa;
        }

        table tbody tr:nth-child(even) {
            background-color: #f9f9f9;
        }

        .badge {
            display: inline-block;
            padding: 6px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
        }

        .badge.recorrente {
            background-color: #e74c3c;
            color: white;
        }

        .badge.critica {
            background-color: #c0392b;
            color: white;
        }

        .badge.potencial {
            background-color: #f39c12;
            color: white;
        }

        .badge.pontual {
            background-color: #3498db;
            color: white;
        }

        .badge.operacional {
            background-color: #27ae60;
            color: white;
        }

        .badge.preventiva {
            background-color: #2980b9;
            color: white;
        }

        .badge.parado {
            background-color: #c0392b;
            color: white;
        }

        .badge.alerta {
            background-color: #f39c12;
            color: white;
        }

        .badge.info {
            background-color: #3498db;
            color: white;
        }

        .df-comparison {
            display: flex;
            gap: 20px;
            align-items: center;
            margin-bottom: 20px;
        }

        .df-bar {
            flex: 1;
            height: 40px;
            background-color: #ecf0f1;
            border-radius: 5px;
            overflow: hidden;
            position: relative;
            border: 1px solid #bdc3c7;
        }

        .df-bar-fill {
            height: 100%;
            background: linear-gradient(90deg, #e74c3c 0%, #f39c12 50%, #27ae60 100%);
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: bold;
            font-size: 12px;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.3);
        }

        .pareto-bar {
            display: flex;
            gap: 10px;
            margin-bottom: 15px;
            align-items: center;
        }

        .pareto-system {
            flex: 0 0 150px;
            font-weight: 600;
            font-size: 12px;
        }

        .pareto-visual {
            flex: 1;
            height: 30px;
            background: linear-gradient(90deg, #3498db 0%, #2980b9 100%);
            border-radius: 3px;
            display: flex;
            align-items: center;
            padding: 0 10px;
            color: white;
            font-weight: 600;
            font-size: 11px;
        }

        .pareto-stats {
            flex: 0 0 200px;
            text-align: right;
            font-size: 11px;
        }

        .fault-classification {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-bottom: 20px;
        }

        .fault-card {
            padding: 15px;
            border-radius: 8px;
            border-left: 4px solid;
        }

        .fault-card.recorrente {
            background-color: #ffebee;
            border-left-color: #e74c3c;
        }

        .fault-card.critica {
            background-color: #ffcdd2;
            border-left-color: #c0392b;
        }

        .fault-card.potencial {
            background-color: #fff3e0;
            border-left-color: #f39c12;
        }

        .fault-card.pontual {
            background-color: #e3f2fd;
            border-left-color: #3498db;
        }

        .fault-count {
            font-size: 20px;
            font-weight: bold;
            margin-bottom: 5px;
        }

        .fault-label {
            font-size: 12px;
            color: #666;
        }

        .project-card {
            background-color: #f0f8ff;
            padding: 20px;
            border-radius: 8px;
            border-left: 4px solid #3498db;
            margin-bottom: 15px;
        }

        .project-title {
            font-weight: 600;
            color: #1b4f72;
            margin-bottom: 8px;
        }

        .project-description {
            font-size: 13px;
            color: #34495e;
            margin-bottom: 8px;
        }

        .project-equipment {
            font-size: 12px;
            color: #666;
            font-style: italic;
            margin-bottom: 8px;
        }

        .project-priority {
            display: inline-block;
            font-size: 11px;
            font-weight: 600;
        }

        .recommendations {
            background: linear-gradient(135deg, #f0f8ff 0%, #e8f4f8 100%);
            padding: 20px;
            border-radius: 8px;
            border-left: 4px solid #3498db;
        }

        .recommendations h3 {
            color: #1b4f72;
            margin-bottom: 15px;
        }

        .recommendations ul {
            margin-left: 20px;
        }

        .recommendations li {
            margin-bottom: 10px;
            color: #34495e;
            line-height: 1.5;
        }

        .developing-pattern {
            background-color: #fff8e1;
            padding: 12px;
            border-left: 4px solid #f39c12;
            margin-bottom: 10px;
            border-radius: 3px;
        }

        .developing-pattern-fault {
            font-weight: 600;
            color: #333;
            margin-bottom: 5px;
        }

        .developing-pattern-equipment {
            font-size: 12px;
            color: #666;
        }

        .footer {
            text-align: center;
            padding: 20px;
            color: #7f8c8d;
            font-size: 12px;
            border-top: 1px solid #ddd;
            margin-top: 40px;
        }

        .no-data {
            text-align: center;
            padding: 20px;
            color: #999;
            font-style: italic;
        }

        @media print {
            body {
                background: white;
            }
            .container {
                max-width: 100%;
            }
            .kpi-card, .section {
                page-break-inside: avoid;
            }
        }

        @media (max-width: 768px) {
            .header {
                padding: 20px;
            }
            .header h1 {
                font-size: 24px;
            }
            .kpi-grid {
                grid-template-columns: 1fr;
            }
            .section {
                padding: 15px;
            }
            .pareto-bar {
                flex-direction: column;
                gap: 5px;
            }
            .pareto-system, .pareto-visual, .pareto-stats {
                flex: 1;
                width: 100%;
            }
        }
    ''')

    parts.append('    </style>')
    parts.append('</head>')
    parts.append('<body>')
    parts.append('    <div class="container">')

    # Header
    parts.append('        <!-- Header -->')
    parts.append('        <div class="header">')
    parts.append(f'            <h1>📊 {title}</h1>')
    parts.append(f'            <p>Relatório gerado em {generation_time}</p>')
    parts.append('        </div>')

    # KPI Cards
    parts.append('        <!-- KPI Cards -->')
    parts.append('        <div class="kpi-grid">')
    parts.append(f'            <div class="kpi-card {df_status_class}">')
    parts.append(f'                <div class="kpi-value">{df_acumulado}</div>')
    parts.append('                <div class="kpi-label">Disponibilidade (DF)</div>')
    parts.append('            </div>')
    parts.append('            <div class="kpi-card">')
    parts.append(f'                <div class="kpi-value">{df_meta}</div>')
    parts.append('                <div class="kpi-label">Meta de DF</div>')
    parts.append('            </div>')
    parts.append('            <div class="kpi-card">')
    parts.append(f'                <div class="kpi-value">{total_failures}</div>')
    parts.append('                <div class="kpi-label">Total de Falhas</div>')
    parts.append('            </div>')
    parts.append(f'            <div class="kpi-card {df_diff_class}">')
    parts.append(f'                <div class="kpi-value">{df_diff}</div>')
    parts.append('                <div class="kpi-label">Diferença vs Meta</div>')
    parts.append('            </div>')
    parts.append('        </div>')

    # DF Analysis
    df_percent = float(df_acumulado.replace('%', ''))
    df_bar_width = min(100, max(0, df_percent))

    df_status_message = '✅ A frota está acima da meta.' if df_status == 'OK' else '⚠️ A frota está abaixo da meta - ação necessária.' if df_status == 'ALERTA' else '🚨 Situação crítica - intervenção urgente recomendada.'

    parts.append('        <!-- DF Analysis -->')
    parts.append('        <div class="section">')
    parts.append('            <h2>📈 Análise de Disponibilidade Física (DF)</h2>')
    parts.append('            <div class="df-comparison">')
    parts.append('                <div style="min-width: 100px;">')
    parts.append('                    <div style="color: #666; font-size: 12px;">Atual vs Meta</div>')
    parts.append('                </div>')
    parts.append('                <div class="df-bar">')
    parts.append(f'                    <div class="df-bar-fill" style="width: {df_bar_width}%;">')
    parts.append(f'                        {df_acumulado}')
    parts.append('                    </div>')
    parts.append('                </div>')
    parts.append('                <div style="min-width: 80px; text-align: right;">')
    parts.append(f'                    <div style="color: #666; font-size: 12px;">Meta: {df_meta}</div>')
    parts.append('                </div>')
    parts.append('            </div>')
    parts.append('            <div style="background: #f9f9f9; padding: 15px; border-radius: 5px; margin-top: 15px;">')
    parts.append('                <strong>Status:</strong>')
    parts.append(f'                <span class="badge {badge_status_class}">')
    parts.append(f'                    {df_status}')
    parts.append('                </span>')
    parts.append(f'                <p style="margin-top: 10px; color: #666; font-size: 13px;">')
    parts.append(f'                    {df_status_message}')
    parts.append('                </p>')
    parts.append('            </div>')
    parts.append('        </div>')

    # Critical Alerts
    if alerts:
        parts.append('        <!-- Critical Alerts -->')
        parts.append('        <div class="section">')
        parts.append('            <h2>⚠️ Alertas Críticos e Recomendações</h2>')
        for data in alerts:
            severity_lower = data['severity'].lower()
            parts.append(f'            <div class="alert {severity_lower}">')
            parts.append(f'                <div class="alert-title">{data["severity"]} - {data["title"]}</div>')
            parts.append(f'                <div class="alert-message">{data["message"]}</div>')
            parts.append(f'                <div class="alert-action">✓ Ação: {data["action"]}</div>')
            parts.append('            </div>')
        parts.append('        </div>')

    # Fault Classification
    parts.append('        <!-- Fault Classification -->')
    parts.append('        <div class="section">')
    parts.append('            <h2>🏷️ Classificação de Falhas</h2>')
    parts.append('            <div class="fault-classification">')
    parts.append('                <div class="fault-card recorrente">')
    parts.append(f'                    <div class="fault-count">🔴 {len(classified_faults["RECORRENTE"])}</div>')
    parts.append('                    <div class="fault-label">Falhas Recorrentes</div>')
    parts.append('                </div>')
    parts.append('                <div class="fault-card critica">')
    parts.append(f'                    <div class="fault-count">🔴 {len(classified_faults["CRITICA"])}</div>')
    parts.append('                    <div class="fault-label">Falhas Críticas</div>')
    parts.append('                </div>')
    parts.append('                <div class="fault-card potencial">')
    parts.append(f'                    <div class="fault-count">🟡 {len(classified_faults["POTENCIAL"])}</div>')
    parts.append('                    <div class="fault-label">Falhas Potenciais</div>')
    parts.append('                </div>')
    parts.append('                <div class="fault-card pontual">')
    parts.append(f'                    <div class="fault-count">🔵 {len(classified_faults["PONTUAL"])}</div>')
    parts.append('                    <div class="fault-label">Falhas Pontuais</div>')
    parts.append('                </div>')
    parts.append('            </div>')

    # Fault details table
    parts.append('            <h3>Detalhes das Falhas Recorrentes e Críticas</h3>')
    if classified_faults['RECORRENTE'] or classified_faults['CRITICA']:
        parts.append('            <table>')
        parts.append('                <tr><th>Tipo</th><th>Equipamento</th><th>Falha</th><th>Data</th><th>Ocorrências</th></tr>')

        for f in classified_faults['RECORRENTE']:
            fault_text = f['fault'][:40] + ('...' if len(f['fault']) > 40 else '')
            parts.append(f"                <tr><td><span class='badge recorrente'>RECORRENTE</span></td><td>{f['equipment']}</td><td>{fault_text}</td><td>{f['date']}</td><td>{f['count']}</td></tr>")

        for f in classified_faults['CRITICA']:
            fault_text = f['fault'][:40] + ('...' if len(f['fault']) > 40 else '')
            parts.append(f"                <tr><td><span class='badge critica'>CRÍTICA</span></td><td>{f['equipment']}</td><td>{fault_text}</td><td>{f['date']}</td><td>{f['count']}</td></tr>")

        parts.append('            </table>')
    else:
        parts.append('            <div class="no-data">Nenhuma falha recorrente ou crítica identificada.</div>')

    parts.append('        </div>')

    # Pareto Analysis
    parts.append('        <!-- Pareto Analysis -->')
    parts.append('        <div class="section">')
    parts.append('            <h2>📊 Análise de Pareto (Sistemas)</h2>')
    parts.append('            <p style="margin-bottom: 20px; color: #666; font-size: 13px;">Os 20% de sistemas que causam 80% das falhas:</p>')

    if pareto_systems:
        for s in pareto_systems[:5]:
            bar_width = s['percentage'] * 3
            parts.append('            <div class="pareto-bar">')
            parts.append(f'                <div class="pareto-system">{s["system"]}</div>')
            parts.append(f'                <div class="pareto-visual" style="width: {bar_width}px;">')
            parts.append(f'                    {s["count"]} falhas')
            parts.append('                </div>')
            parts.append('                <div class="pareto-stats">')
            parts.append(f'                    {s["percentage"]:.1f}% | Acum: {s["cumulative_pct"]:.1f}%')
            parts.append('                </div>')
            parts.append('            </div>')
    else:
        parts.append('            <div class="no-data">Nenhum dado de sistema disponível.</div>')

    parts.append('        </div>')

    # Developing Patterns
    if developing_patterns:
        parts.append('        <!-- Developing Patterns -->')
        parts.append('        <div class="section">')
        parts.append('            <h2>📈 Falhas em Desenvolvimento (Padrões Emergentes)</h2>')
        for p in developing_patterns:
            parts.append('            <div class="developing-pattern">')
            parts.append(f'                <div class="developing-pattern-fault">{p["fault"]}</div>')
            parts.append(f'                <div class="developing-pattern-equipment">Equipamentos: {p["equipments"]} | Ocorrências: {p["count"]}</div>')
            parts.append('            </div>')
        parts.append('        </div>')

    # Equipment Status
    parts.append('        <!-- Equipment Status -->')
    parts.append('        <div class="section">')
    parts.append('            <h2>🚜 Status dos Equipamentos</h2>')
    parts.append('            <table>')
    parts.append('                <tr><th>Equipamento</th><th>Status</th><th>DF</th><th>Falhas</th></tr>')
    for data in equipment_details:
        status_lower = data['status'].lower()
        parts.append(f"                <tr><td>{data['equipment']}</td><td><span class='badge {status_lower}'>{data['status']}</span></td><td>{data['df']:.2f}%</td><td>{data['failures']}</td></tr>")
    parts.append('            </table>')
    parts.append('        </div>')

    # Top Systems
    parts.append('        <!-- Top Systems -->')
    parts.append('        <div class="section">')
    parts.append('            <h2>🔧 TOP 3 Sistemas com Mais Falhas</h2>')
    if top_systems:
        parts.append('            <table>')
        parts.append('                <tr><th>Posição</th><th>Sistema</th><th>Quantidade de Falhas</th></tr>')
        for i, data in enumerate(top_systems):
            parts.append(f"                <tr><td>#{i+1}</td><td>{data['name']}</td><td><span class='badge info'>{data['count']}</span></td></tr>")
        parts.append('            </table>')
    else:
        parts.append('            <div class="no-data">Nenhum sistema com falhas.</div>')
    parts.append('        </div>')

    # Top Equipment
    parts.append('        <!-- Top Equipment -->')
    parts.append('        <div class="section">')
    parts.append('            <h2>🚜 TOP 3 Equipamentos com Mais Falhas</h2>')
    if top_equipment:
        parts.append('            <table>')
        parts.append('                <tr><th>Posição</th><th>Equipamento</th><th>Quantidade de Falhas</th><th>Severity</th></tr>')
        for i, data in enumerate(top_equipment):
            severity_badge = 'critica' if data['count'] >= 3 else 'warning' if data['count'] == 2 else 'info'
            severity_text = 'CRÍTICO' if data['count'] >= 3 else 'ATENÇÃO' if data['count'] == 2 else 'OK'
            parts.append(f"                <tr><td>#{i+1}</td><td>{data['name']}</td><td><span class='badge info'>{data['count']}</span></td><td><span class='badge {severity_badge}'>{severity_text}</span></td></tr>")
        parts.append('            </table>')
    else:
        parts.append('            <div class="no-data">Nenhum equipamento com falhas.</div>')
    parts.append('        </div>')

    # Suggested Projects
    if suggested_projects:
        parts.append('        <!-- Suggested Projects -->')
        parts.append('        <div class="section">')
        parts.append('            <h2>💼 Projetos Sugeridos (Baseados em Falhas Recorrentes)</h2>')
        for p in suggested_projects:
            priority_badge = 'critica' if p['priority'] == 'ALTA' else 'warning'
            parts.append('            <div class="project-card">')
            parts.append(f'                <div class="project-title">{p["title"]}</div>')
            parts.append(f'                <div class="project-description">{p["description"]}</div>')
            parts.append(f'                <div class="project-equipment">Equipamentos afetados: {p["equipments"]}</div>')
            parts.append(f'                <span class="badge {priority_badge} project-priority">{p["priority"]}</span>')
            parts.append('            </div>')
        parts.append('        </div>')

    # Recommendations
    parts.append('        <!-- Recommendations -->')
    parts.append('        <div class="section recommendations">')
    parts.append('            <h3>💡 Recomendações Estratégicas</h3>')
    parts.append('            <ul>')
    parts.append('                <li><strong>Manutenção Preventiva:</strong> Implementar calendário baseado nos TOP 3 sistemas com falhas para evitar paradas não planejadas.</li>')
    parts.append('                <li><strong>Análise de Raiz Causa:</strong> Investigar padrões recorrentes e em desenvolvimento para eliminar causas raízes.</li>')
    parts.append('                <li><strong>Priorização de Equipamentos:</strong> Concentrar esforços nos equipamentos críticos (≥3 falhas ou DF < 50%).</li>')
    parts.append('                <li><strong>Monitoramento de Padrões:</strong> Acompanhar falhas em desenvolvimento antes que se tornem críticas.</li>')
    parts.append(f'                <li><strong>Meta de DF:</strong> Alcançar {df_meta} de disponibilidade através de manutenção planejada e projetos direcionados.</li>')
    parts.append('                <li><strong>Acompanhamento Contínuo:</strong> Utilizar este relatório semanalmente para acompanhar tendências e validar efetividade das ações.</li>')
    parts.append('            </ul>')
    parts.append('        </div>')

    # Monthly Trend Placeholder
    parts.append('        <!-- Monthly Trend Placeholder -->')
    parts.append('        <div class="section">')
    parts.append('            <h2>📅 Dashboard de Tendências Mensais (Em Desenvolvimento)</h2>')
    parts.append('            <div style="background: #f0f0f0; padding: 30px; border-radius: 5px; text-align: center;">')
    parts.append('                <p style="color: #999;">Este gráfico será preenchido conforme mais dados históricos forem coletados.</p>')
    parts.append('                <p style="color: #999; font-size: 12px; margin-top: 10px;">Será exibido um gráfico de tendência de DF e falhas ao longo do tempo.</p>')
    parts.append('            </div>')
    parts.append('        </div>')

    # Footer
    parts.append('        <!-- Footer -->')
    parts.append('        <div class="footer">')
    parts.append('            <p>Relatório gerado automaticamente pelo Sistema de Análise de Falhas e Confiabilidade da Frota</p>')
    parts.append(f'            <p>Data/Hora: {generation_time}</p>')
    parts.append('            <p style="margin-top: 10px; opacity: 0.6;">Esse documento é confidencial e destinado apenas para análise interna.</p>')
    parts.append('        </div>')

    parts.append('    </div>')

    # Script
    parts.append('    <script>')
    parts.append('        function printReport() {')
    parts.append('            window.print();')
    parts.append('        }')
    parts.append('    </script>')
    parts.append('</body>')
    parts.append('</html>')

    return '\n'.join(parts)

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
