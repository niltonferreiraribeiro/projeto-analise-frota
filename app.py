from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.utils import get_column_letter
import os
from datetime import datetime
import traceback

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max
app.config['UPLOAD_FOLDER'] = 'uploads'

# Create uploads folder if it doesn't exist
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

ALLOWED_EXTENSIONS = {'xlsx', 'xlsm'}

def allowed_file(filename):
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
    # Excel time as decimal day fraction
    try:
        if hasattr(time_obj, 'hour'):
            return time_obj.hour * 3600 + time_obj.minute * 60 + time_obj.second
        return 0
    except:
        return 0

def seconds_to_hms(seconds):
    """Convert seconds to HH:MM:SS"""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"

def process_excel(filepath, df_meta):
    """Process Excel file and generate analysis"""
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)

        # Check required sheets
        if 'Escavadeira' not in wb.sheetnames:
            return None, "Aba 'Escavadeira' não encontrada"
        if 'DF' not in wb.sheetnames:
            return None, "Aba 'DF' não encontrada"

        ws_failures = wb['Escavadeira']
        ws_df = wb['DF']

        # Process failures data
        failures = []
        equipment_failures = {}  # Count by equipment
        system_failures = {}  # Count by system
        subsystem_failures = {}  # Count by subsystem
        pattern_failures = {}  # Identify recurrent patterns
        equipment_details = {}  # Store equipment details

        row = 2
        while True:
            equipment = ws_failures[f'B{row}'].value
            if not equipment:
                break

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

            # Track patterns (same failure description)
            pattern_key = f"{equipment}|{failure_description}"
            if pattern_key not in pattern_failures:
                pattern_failures[pattern_key] = {'count': 0, 'details': failure_record}
            pattern_failures[pattern_key]['count'] += 1

            row += 1

        # Read DF data from DF sheet
        df_acumulado = 0
        escavadeira_df_data = []

        # Read Escavadeira DF data (columns F-H)
        for row in range(2, 12):
            equipment_df = ws_df[f'F{row}'].value
            df_value = ws_df[f'G{row}'].value
            week = ws_df[f'H{row}'].value

            if equipment_df and df_value is not None:
                if equipment_df == "Acumulado da frota":
                    # Parse percentage
                    if isinstance(df_value, str):
                        df_acumulado = float(df_value.replace('%', '').replace(',', '.'))
                    else:
                        df_acumulado = float(df_value) * 100 if df_value < 1 else float(df_value)
                else:
                    if isinstance(df_value, str):
                        df_pct = float(df_value.replace('%', '').replace(',', '.'))
                    else:
                        df_pct = float(df_value) * 100 if df_value < 1 else float(df_value)

                    escavadeira_df_data.append({
                        'equipment': str(equipment_df),
                        'df': df_pct,
                        'week': week
                    })

        # Identify critical patterns
        critical_failures = {}
        for pattern_key, data in pattern_failures.items():
            if data['count'] >= 2:
                equipment, desc = pattern_key.split('|', 1)
                critical_failures[pattern_key] = data

        # TOP 3 rankings
        top_systems = sorted(system_failures.items(), key=lambda x: x[1], reverse=True)[:3]
        top_subsystems = sorted(subsystem_failures.items(), key=lambda x: x[1], reverse=True)[:3]
        top_equipment = sorted(equipment_failures.items(), key=lambda x: x[1], reverse=True)[:3]

        # Generate alerts
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
        for pattern_key, data in critical_failures.items():
            if data['count'] >= 2:
                equipment, desc = pattern_key.split('|', 1)
                alerts.append({
                    'severity': 'MÉDIO',
                    'title': 'Falha Recorrente',
                    'message': f'{desc[:50]}... ({data["count"]}x)',
                    'action': 'Investigação técnica necessária'
                })

        # Prepare analysis data
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
            'critical_patterns': critical_failures,
            'alerts': alerts,
            'equipment_details': equipment_details,
            'escavadeira_df': escavadeira_df_data,
            'generation_time': datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        }

        wb.close()
        return analysis, None

    except Exception as e:
        return None, f"Erro ao processar arquivo: {str(e)}"

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

        df_meta = request.form.get('df_meta', 90.50)
        try:
            df_meta = float(df_meta)
        except:
            df_meta = 90.50

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
    """Generate professional HTML report"""

    # Prepare data for template
    report_data = {
        'title': 'Análise de Falhas e Confiabilidade - Frota de Escavadeiras CAT',
        'generation_time': analysis['generation_time'],
        'total_failures': analysis['total_failures'],
        'df_acumulado': f"{analysis['df_acumulado']:.2f}%",
        'df_meta': f"{analysis['df_meta']:.2f}%",
        'df_status': analysis['df_status'],
        'df_diff': f"{analysis['df_acumulado'] - analysis['df_meta']:.2f}%",
        'top_systems': [{'name': name, 'count': count} for name, count in analysis['top_systems']],
        'top_subsystems': [{'name': name, 'count': count} for name, count in analysis['top_subsystems']],
        'top_equipment': [{'name': name, 'count': count} for name, count in analysis['top_equipment']],
        'alerts': analysis['alerts'],
        'equipment_details': [
            {'equipment': eq, 'failures': count}
            for eq, count in sorted(analysis['equipments'].items())
        ],
        'critical_patterns': [
            {'pattern': key.split('|')[1], 'equipment': key.split('|')[0], 'count': data['count']}
            for key, data in analysis['critical_patterns'].items()
        ]
    }

    html = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{report_data['title']}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f5f5f5;
            color: #333;
            line-height: 1.6;
        }}

        .container {{
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }}

        .header {{
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            color: white;
            padding: 40px;
            border-radius: 10px;
            margin-bottom: 30px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}

        .header h1 {{
            font-size: 32px;
            margin-bottom: 10px;
        }}

        .header p {{
            font-size: 14px;
            opacity: 0.9;
        }}

        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}

        .kpi-card {{
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            border-left: 5px solid #2a5298;
            transition: transform 0.3s;
        }}

        .kpi-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
        }}

        .kpi-card.critical {{
            border-left-color: #e74c3c;
        }}

        .kpi-card.warning {{
            border-left-color: #f39c12;
        }}

        .kpi-card.success {{
            border-left-color: #27ae60;
        }}

        .kpi-value {{
            font-size: 28px;
            font-weight: bold;
            color: #1e3c72;
            margin-bottom: 5px;
        }}

        .kpi-card.critical .kpi-value {{
            color: #e74c3c;
        }}

        .kpi-card.warning .kpi-value {{
            color: #f39c12;
        }}

        .kpi-card.success .kpi-value {{
            color: #27ae60;
        }}

        .kpi-label {{
            font-size: 14px;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}

        .section {{
            background: white;
            padding: 30px;
            border-radius: 10px;
            margin-bottom: 30px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}

        .section h2 {{
            font-size: 22px;
            color: #1e3c72;
            margin-bottom: 20px;
            border-bottom: 3px solid #2a5298;
            padding-bottom: 10px;
        }}

        .alert {{
            padding: 15px;
            margin-bottom: 15px;
            border-radius: 5px;
            border-left: 4px solid;
        }}

        .alert.critical {{
            background-color: #fadbd8;
            border-left-color: #e74c3c;
            color: #a93226;
        }}

        .alert.high {{
            background-color: #fef5e7;
            border-left-color: #f39c12;
            color: #9a6c06;
        }}

        .alert.medium {{
            background-color: #d6eaf8;
            border-left-color: #3498db;
            color: #1b4f72;
        }}

        .alert-title {{
            font-weight: bold;
            font-size: 14px;
            margin-bottom: 5px;
        }}

        .alert-message {{
            font-size: 13px;
            margin-bottom: 5px;
        }}

        .alert-action {{
            font-size: 12px;
            font-style: italic;
            margin-top: 8px;
            padding-top: 8px;
            border-top: 1px solid currentColor;
            opacity: 0.8;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }}

        table th {{
            background-color: #2a5298;
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
        }}

        table td {{
            padding: 12px 15px;
            border-bottom: 1px solid #ddd;
        }}

        table tbody tr:hover {{
            background-color: #f8f9fa;
        }}

        table tbody tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}

        .badge {{
            display: inline-block;
            padding: 5px 10px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
        }}

        .badge.critical {{
            background-color: #e74c3c;
            color: white;
        }}

        .badge.warning {{
            background-color: #f39c12;
            color: white;
        }}

        .badge.success {{
            background-color: #27ae60;
            color: white;
        }}

        .badge.info {{
            background-color: #3498db;
            color: white;
        }}

        .df-comparison {{
            display: flex;
            gap: 20px;
            align-items: center;
            margin-bottom: 20px;
        }}

        .df-bar {{
            flex: 1;
            height: 40px;
            background-color: #ecf0f1;
            border-radius: 5px;
            overflow: hidden;
            position: relative;
        }}

        .df-bar-fill {{
            height: 100%;
            background: linear-gradient(90deg, #e74c3c 0%, #f39c12 50%, #27ae60 100%);
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: bold;
            font-size: 12px;
        }}

        .recommendations {{
            background-color: #f0f8ff;
            padding: 20px;
            border-radius: 5px;
            border-left: 4px solid #3498db;
        }}

        .recommendations h3 {{
            color: #1b4f72;
            margin-bottom: 15px;
        }}

        .recommendations ul {{
            margin-left: 20px;
        }}

        .recommendations li {{
            margin-bottom: 10px;
            color: #34495e;
        }}

        .footer {{
            text-align: center;
            padding: 20px;
            color: #7f8c8d;
            font-size: 12px;
            border-top: 1px solid #ddd;
            margin-top: 40px;
        }}

        .no-data {{
            text-align: center;
            padding: 20px;
            color: #999;
        }}

        @media print {{
            body {{
                background-color: white;
            }}
            .container {{
                max-width: 100%;
            }}
            .kpi-card {{
                page-break-inside: avoid;
            }}
            .section {{
                page-break-inside: avoid;
            }}
        }}

        @media (max-width: 768px) {{
            .header {{
                padding: 20px;
            }}
            .header h1 {{
                font-size: 24px;
            }}
            .kpi-grid {{
                grid-template-columns: 1fr;
            }}
            .section {{
                padding: 15px;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <!-- Header -->
        <div class="header">
            <h1>📊 {report_data['title']}</h1>
            <p>Relatório gerado em {report_data['generation_time']}</p>
        </div>

        <!-- KPI Cards -->
        <div class="kpi-grid">
            <div class="kpi-card {'success' if report_data['df_status'] == 'OK' else 'warning' if report_data['df_status'] == 'ALERTA' else 'critical'}">
                <div class="kpi-value">{report_data['df_acumulado']}</div>
                <div class="kpi-label">Disponibilidade (DF)</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-value">{report_data['df_meta']}</div>
                <div class="kpi-label">Meta de DF</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-value">{report_data['total_failures']}</div>
                <div class="kpi-label">Total de Falhas</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-value">{report_data['df_diff']}</div>
                <div class="kpi-label">Diferença vs Meta</div>
            </div>
        </div>

        <!-- DF Analysis -->
        <div class="section">
            <h2>📈 Análise de Disponibilidade Física (DF)</h2>
            <div class="df-comparison">
                <div style="min-width: 100px;">
                    <div style="color: #666; font-size: 12px;">Atual vs Meta</div>
                </div>
                <div class="df-bar">
                    <div class="df-bar-fill" style="width: {min(100, max(0, float(report_data['df_acumulado'].replace('%', ''))))}%;">
                        {report_data['df_acumulado']}
                    </div>
                </div>
                <div style="min-width: 80px; text-align: right;">
                    <div style="color: #666; font-size: 12px;">Meta: {report_data['df_meta']}</div>
                </div>
            </div>
            <div style="background: #f9f9f9; padding: 15px; border-radius: 5px; margin-top: 15px;">
                <strong>Status:</strong>
                <span class="badge {'success' if report_data['df_status'] == 'OK' else 'warning' if report_data['df_status'] == 'ALERTA' else 'critical'}">
                    {report_data['df_status']}
                </span>
                <p style="margin-top: 10px; color: #666; font-size: 13px;">
                    {'✅ A frota está acima da meta.' if report_data['df_status'] == 'OK' else '⚠️ A frota está abaixo da meta - ação necessária.' if report_data['df_status'] == 'ALERTA' else '🚨 Situação crítica - intervenção urgente recomendada.'}
                </p>
            </div>
        </div>

        <!-- Critical Alerts -->
        {f"""<div class="section">
            <h2>⚠️ Alertas Críticos e Recomendações</h2>
            {''.join([f'''<div class="alert {data['severity'].lower()}">
                <div class="alert-title">{data['severity']} - {data['title']}</div>
                <div class="alert-message">{data['message']}</div>
                <div class="alert-action">✓ Ação: {data['action']}</div>
            </div>''' for data in report_data['alerts']])}
            {'' if report_data['alerts'] else '<div class="no-data">Nenhum alerta crítico identificado.</div>'}
        </div>""" if report_data['alerts'] else ""}

        <!-- Top Systems -->
        <div class="section">
            <h2>🔧 TOP 3 Sistemas com Mais Falhas</h2>
            {''.join([f"""<table>
                <tr><th>Posição</th><th>Sistema</th><th>Quantidade de Falhas</th></tr>
                {''.join([f"<tr><td>#{i+1}</td><td>{data['name']}</td><td><span class='badge info'>{data['count']}</span></td></tr>" for i, data in enumerate(report_data['top_systems'])])}
            </table>"""])}
        </div>

        <!-- Top Subsystems -->
        <div class="section">
            <h2>⚙️ TOP 3 Sub-Sistemas com Mais Falhas</h2>
            {''.join([f"""<table>
                <tr><th>Posição</th><th>Sub-Sistema</th><th>Quantidade de Falhas</th></tr>
                {''.join([f"<tr><td>#{i+1}</td><td>{data['name']}</td><td><span class='badge info'>{data['count']}</span></td></tr>" for i, data in enumerate(report_data['top_subsystems'])])}
            </table>"""])}
        </div>

        <!-- Top Equipment -->
        <div class="section">
            <h2>🚜 TOP 3 Equipamentos com Mais Falhas</h2>
            {''.join([f"""<table>
                <tr><th>Posição</th><th>Equipamento</th><th>Quantidade de Falhas</th></tr>
                {''.join([f"<tr><td>#{i+1}</td><td>{data['name']}</td><td><span class='badge {'critical' if data['count'] >= 3 else 'warning' if data['count'] == 2 else 'info'}'>{data['count']}</span></td></tr>" for i, data in enumerate(report_data['top_equipment'])])}
            </table>"""])}
        </div>

        <!-- All Equipment Summary -->
        <div class="section">
            <h2>📋 Resumo por Equipamento</h2>
            <table>
                <tr><th>Equipamento</th><th>Quantidade de Falhas</th><th>Status</th></tr>
                {''.join([f"<tr><td>{data['equipment']}</td><td>{data['failures']}</td><td><span class='badge {'critical' if data['failures'] >= 3 else 'warning' if data['failures'] >= 2 else 'success'}'>{data['failures']} falha(s)</span></td></tr>" for data in report_data['equipment_details']])}
            </table>
        </div>

        <!-- Recurrent Patterns -->
        {f"""<div class="section">
            <h2>🔄 Padrões Recorrentes Identificados</h2>
            {f'''<table>
                <tr><th>Equipamento</th><th>Falha Recorrente</th><th>Ocorrências</th><th>Severidade</th></tr>
                {''.join([f"<tr><td>{data['equipment']}</td><td>{data['pattern'][:50]}{'...' if len(data['pattern']) > 50 else ''}</td><td><span class='badge info'>{data['count']}x</span></td><td><span class='badge warning'>ATENÇÃO</span></td></tr>" for data in report_data['critical_patterns']])}
            </table>''' if report_data['critical_patterns'] else '<div class="no-data">Nenhum padrão recorrente identificado.</div>'}
        </div>""" if report_data['critical_patterns'] else ""}

        <!-- Recommendations -->
        <div class="section recommendations">
            <h3>💡 Recomendações Estratégicas</h3>
            <ul>
                <li><strong>Manutenção Preventiva:</strong> Implementar calendário baseado nos TOP 3 sistemas com falhas para evitar paradas não planejadas.</li>
                <li><strong>Análise de Raiz Causa:</strong> Investigar padrões recorrentes para eliminar causas raízes.</li>
                <li><strong>Priorização de Equipamentos:</strong> Concentrar esforços nos equipamentos críticos (≥3 falhas).</li>
                <li><strong>Meta de DF:</strong> Alcançar {report_data['df_meta']} de disponibilidade através de manutenção planejada.</li>
                <li><strong>Monitoramento Contínuo:</strong> Utilizar este relatório semanalmente para acompanhar tendências.</li>
            </ul>
        </div>

        <!-- Footer -->
        <div class="footer">
            <p>Relatório gerado automaticamente pelo Sistema de Análise de Falhas e Confiabilidade da Frota</p>
            <p>Data/Hora: {report_data['generation_time']}</p>
            <p style="margin-top: 10px; opacity: 0.6;">Esse documento é confidencial e destinado apenas para análise interna.</p>
        </div>
    </div>

    <script>
        function printReport() {{
            window.print();
        }}

        function downloadPDF() {{
            const element = document.querySelector('.container');
            const opt = {{
                margin: 10,
                filename: 'Analise_Falhas_Escavadeiras.pdf',
                image: {{ type: 'jpeg', quality: 0.98 }},
                html2canvas: {{ scale: 2 }},
                jsPDF: {{ orientation: 'portrait', unit: 'mm', format: 'a4' }}
            }};
        }}
    </script>
</body>
</html>
"""
    return html

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
