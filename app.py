"""
Bilant Generator - Flask Web Application
Generates Romanian Balance Sheet (Bilant) from Trial Balance (Balanta)
"""

from flask import Flask, request, render_template, send_file, jsonify
import pandas as pd
import numpy as np
import re
import io
import os
from werkzeug.utils import secure_filename
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

app = Flask(__name__, static_folder='static')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = 'uploads'

# Ensure folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('static', exist_ok=True)

# =============================================================================
# CONSTANTS - Match template layout
# =============================================================================

# Balanta layout (0-indexed for pandas)
# Template: Cont, SFD, SFC (3 columns only)
COL_BAL_ACCOUNT = 0   # A = Cont (account number)
COL_BAL_SFD = 1       # B = SFD
COL_BAL_SFC = 2       # C = SFC

# Bilant layout
COL_BIL_DESC = 0      # A = Denumirea elementului
COL_BIL_NR_RD = 1     # B = Nr. rd.
COL_BIL_VAL = 2       # C = Sold Final


# =============================================================================
# STEP 1: Prepare Balanta data
# =============================================================================

def prepare_balanta(df_balanta):
    """
    Prepare Balanta dataframe by skipping header row if present.
    """
    df = df_balanta.copy()

    # Skip header row if present
    if len(df) > 0 and df.iloc[0, COL_BAL_ACCOUNT] == 'Cont':
        df = df.iloc[1:].reset_index(drop=True)

    return df


# =============================================================================
# STEP 2: Extract CT formulas from Bilant descriptions
# =============================================================================

def extract_ct_formula(description):
    """
    Extract account formula from description text.
    Example: "1.Cheltuieli de constituire (ct.201-2801)" -> "201-2801"
    Based on VBA GetCtExpression: finds "ct." then extracts until ")"
    """
    if pd.isna(description):
        return ""

    text = str(description)

    # Find "ct." position (case insensitive)
    # MUST have the dot to avoid matching "ct" in words like "active"
    match = re.search(r'ct\.\s*', text, re.IGNORECASE)
    if not match:
        return ""

    # Start position after "ct." and any spaces
    start_pos = match.end()

    # Find closing parenthesis
    paren_pos = text.find(')', start_pos)
    if paren_pos == -1:
        paren_pos = len(text)

    # Extract the expression
    expr = text[start_pos:paren_pos].strip()

    # Clean up the expression (NormalizeCtFormula from VBA)
    expr = expr.replace('*', '')     # Remove asterisks
    expr = re.sub(r'\s+', '', expr)  # Remove whitespace
    expr = expr.replace('\r', '')    # Remove carriage return
    expr = expr.replace('\n', '')    # Remove newline

    return expr


# =============================================================================
# STEP 3: Extract row formulas from Bilant descriptions
# =============================================================================

def extract_row_formula(description):
    """
    Extract row formula from description text.
    Example: "TOTAL (rd. 01 la 06)" -> "01+02+03+04+05+06"
    Example: "TOTAL (rd. 31 la 35 +35a)" -> "31+32+33+34+35+35a"
    """
    if pd.isna(description):
        return ""

    text = str(description).lower()

    # Find "rd." followed by formula
    match = re.search(r'rd\.?\s*([^)]+)', text)
    if not match:
        return ""

    raw = match.group(1).strip()
    raw = re.sub(r'\s+', '', raw)

    # Handle "01 la 06" format - expand the range first
    la_match = re.search(r'(\d+)la(\d+)', raw)
    if la_match:
        start = int(la_match.group(1))
        end = int(la_match.group(2))
        width = len(la_match.group(1))

        if end >= start:
            parts = [str(i).zfill(width) for i in range(start, end + 1)]
            expanded = '+'.join(parts)
            # Replace the "XXlaYY" with expanded form, keep any additional terms
            raw = raw[:la_match.start()] + expanded + raw[la_match.end():]

    # Extract row references (numbers with optional letter suffix) and signs
    # Keep alphanumeric row references like "35a"
    result = re.sub(r'[^0-9a-z+\-]', '', raw)

    # Convert alphanumeric references like "35a" to numeric "36"
    # This handles template errors where "35a" should actually be "36"
    result = re.sub(r'35a', '36', result)

    return result


# =============================================================================
# STEP 4: Evaluate CT expressions
# =============================================================================

def parse_ct_formula(expr):
    """
    Parse CT formula into list of (prefix, sign_type) tuples.
    sign_type: 'normal_plus', 'normal_minus', 'dynamic'

    Handles: 345+346-2801+/-348-dinct.4428
    """
    if not expr:
        return []

    items = []
    i = 0
    sign = 1  # 1 for plus, -1 for minus

    while i < len(expr):
        # Check for "+/-" dynamic sign
        if expr[i:i+3] == '+/-':
            i += 3
            # Read the number after +/-
            num = ''
            while i < len(expr) and expr[i].isdigit():
                num += expr[i]
                i += 1
            if num:
                items.append((num, 'dynamic'))
            continue

        # Check for "dinct." special case
        if expr[i:i+6].lower() == 'dinct.':
            i += 6
            num = ''
            while i < len(expr) and expr[i].isdigit():
                num += expr[i]
                i += 1
            if num:
                items.append((num, 'normal_minus'))  # dinct is always subtracted
            continue

        # Handle signs
        if expr[i] == '+':
            sign = 1
            i += 1
            continue
        elif expr[i] == '-':
            sign = -1
            i += 1
            continue

        # Read number
        if expr[i].isdigit():
            num = ''
            while i < len(expr) and expr[i].isdigit():
                num += expr[i]
                i += 1
            if num:
                sign_type = 'normal_plus' if sign == 1 else 'normal_minus'
                items.append((num, sign_type))
            sign = 1  # Reset sign
            continue

        # Skip other characters
        i += 1

    return items


def sum_accounts_by_prefix(df_balanta, prefix, use_net=False):
    """
    Sum all accounts starting with the given prefix.
    Returns (total, account_details)
    """
    total = 0
    details = []

    for idx, row in df_balanta.iterrows():
        acct = str(row.iloc[COL_BAL_ACCOUNT])
        # Clean up account number (remove .0 suffix if present)
        if acct.endswith('.0'):
            acct = acct[:-2]

        if acct.startswith(prefix):
            sfd = pd.to_numeric(row.iloc[COL_BAL_SFD], errors='coerce') or 0
            sfc = pd.to_numeric(row.iloc[COL_BAL_SFC], errors='coerce') or 0

            if use_net:
                # For dynamic +/- terms: use SFD - SFC
                val = sfd - sfc
            else:
                # For normal terms: use abs(SFD) + abs(SFC)
                val = abs(sfd) + abs(sfc)

            total += val
            details.append((acct, val))

    return total, details


def eval_ct_expression(expr, df_balanta):
    """
    Evaluate CT expression and return (result, verification_details).
    """
    items = parse_ct_formula(expr)
    total = 0
    all_details = []

    for prefix, sign_type in items:
        if sign_type == 'dynamic':
            # +/- means SFD - SFC per account
            subtotal, details = sum_accounts_by_prefix(df_balanta, prefix, use_net=True)
            if not details:
                # No accounts found for this prefix
                all_details.append((prefix, 'No Val.', prefix, 'dynamic'))
            else:
                for acct, val in details:
                    all_details.append((acct, val, prefix, 'dynamic'))
            total += subtotal
        elif sign_type == 'normal_plus':
            subtotal, details = sum_accounts_by_prefix(df_balanta, prefix, use_net=False)
            if not details:
                # No accounts found for this prefix
                all_details.append((prefix, 'No Val.', prefix, '+'))
            else:
                for acct, val in details:
                    all_details.append((acct, val, prefix, '+'))
            total += subtotal
        elif sign_type == 'normal_minus':
            subtotal, details = sum_accounts_by_prefix(df_balanta, prefix, use_net=False)
            if not details:
                # No accounts found for this prefix
                all_details.append((prefix, 'No Val.', prefix, '-'))
            else:
                for acct, val in details:
                    all_details.append((acct, -val, prefix, '-'))
            total -= subtotal

    return total, all_details


# =============================================================================
# STEP 5: Evaluate row formulas
# =============================================================================

def eval_row_formula(expr, bilant_values):
    """
    Evaluate row formula referencing other Bilant rows.
    bilant_values: dict mapping Nr.rd -> value
    Handles alphanumeric row references like "35a"
    """
    if not expr:
        return 0

    total = 0
    sign = 1
    row_ref = ''

    for ch in expr + '+':  # Add + to flush last reference
        if ch.isdigit() or ch.isalpha():
            row_ref += ch
        elif ch in '+-':
            if row_ref:
                # Strip leading zeros from numeric part but keep letters
                # "035a" -> "35a", "01" -> "1"
                match = re.match(r'^0*(\d+[a-z]*)$', row_ref)
                if match:
                    row_num = match.group(1) or '0'
                else:
                    row_num = row_ref
                val = bilant_values.get(row_num, 0)
                total += sign * val
                row_ref = ''
            sign = 1 if ch == '+' else -1

    return total


# =============================================================================
# MAIN PROCESSING FUNCTION
# =============================================================================

def process_bilant(df_balanta, df_bilant):
    """
    Process Balanta and generate Bilant with calculations and verification.
    """
    # Step 1: Prepare Balanta data
    df_balanta = prepare_balanta(df_balanta)

    # Prepare Bilant dataframe
    df_bilant = df_bilant.copy()

    # Extract formulas
    df_bilant['Formula_CT'] = df_bilant.iloc[:, COL_BIL_DESC].apply(extract_ct_formula)
    df_bilant['Formula_RD'] = df_bilant.iloc[:, COL_BIL_DESC].apply(extract_row_formula)

    # Initialize results
    results = []
    verifications = []
    bilant_values = {}  # For row formula references

    # First pass: calculate CT formulas
    for idx, row in df_bilant.iterrows():
        nr_rd = str(row.iloc[COL_BIL_NR_RD]).replace('.0', '') if pd.notna(row.iloc[COL_BIL_NR_RD]) else ''
        expr_ct = row['Formula_CT']
        expr_rd = row['Formula_RD']

        val = 0
        verification = ""

        if expr_ct:
            val, details = eval_ct_expression(expr_ct, df_balanta)

            # Build verification string
            verif_lines = []
            for acct, acct_val, prefix, sign_type in details:
                if acct_val == 'No Val.':
                    verif_lines.append(f"{acct} = No Val.")
                else:
                    verif_lines.append(f"{acct} = {acct_val:.2f}")
            verification = '\n'.join(verif_lines)

        results.append(val)
        verifications.append(verification)

        # Store for row formula references
        if nr_rd:
            bilant_values[nr_rd] = val

    # Second pass: calculate RD formulas (for TOTAL rows)
    for idx, row in df_bilant.iterrows():
        expr_rd = row['Formula_RD']
        expr_ct = row['Formula_CT']

        # Only use RD formula if no CT formula
        if expr_rd and not expr_ct:
            val = eval_row_formula(expr_rd, bilant_values)
            results[idx] = val
            verifications[idx] = f"Sum of rows: {expr_rd}"

            # Update bilant_values
            nr_rd = str(row.iloc[COL_BIL_NR_RD]).replace('.0', '') if pd.notna(row.iloc[COL_BIL_NR_RD]) else ''
            if nr_rd:
                bilant_values[nr_rd] = val

    # Write results to the Sold Final column (COL_BIL_VAL)
    col_name = df_bilant.columns[COL_BIL_VAL]
    df_bilant[col_name] = results

    # Add verification as new column
    df_bilant['Verification'] = verifications

    return df_balanta, df_bilant


# =============================================================================
# EXCEL OUTPUT
# =============================================================================

def create_output_excel(df_balanta, df_bilant):
    """
    Create output Excel file with processed Balanta, Bilant, and Dashboard.
    """
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Write Balanta with Sold Final
        df_balanta.to_excel(writer, sheet_name='Balanta', index=False)

        # Write Bilant with calculations (exclude Formula_CT and Formula_RD columns)
        columns_to_exclude = ['Formula_CT', 'Formula_RD']
        df_bilant_export = df_bilant.drop(columns=[col for col in columns_to_exclude if col in df_bilant.columns])
        df_bilant_export.to_excel(writer, sheet_name='Bilant', index=False)

        # Calculate and write Dashboard
        metrics = calculate_bi_metrics(df_bilant)

        # Create Dashboard dataframe
        dashboard_data = []

        # Summary section
        dashboard_data.append(['SUMAR FINANCIAR', '', ''])
        dashboard_data.append(['Indicator', 'Valoare', ''])
        dashboard_data.append(['Total Active', metrics['summary']['total_active'], ''])
        dashboard_data.append(['Active Imobilizate', metrics['summary']['active_imobilizate'], ''])
        dashboard_data.append(['Active Circulante', metrics['summary']['active_circulante'], ''])
        dashboard_data.append(['Capitaluri Proprii', metrics['summary']['capitaluri_proprii'], ''])
        dashboard_data.append(['Total Datorii', metrics['summary']['total_datorii'], ''])
        dashboard_data.append(['', '', ''])

        # Ratios section
        dashboard_data.append(['INDICATORI FINANCIARI', '', ''])
        dashboard_data.append(['Indicator', 'Valoare', 'Interpretare'])

        ratio_labels = {
            'lichiditate_curenta': ('Lichiditate Curenta', 'Ideal > 1'),
            'lichiditate_rapida': ('Lichiditate Rapida', 'Ideal > 0.8'),
            'lichiditate_imediata': ('Lichiditate Imediata', 'Ideal > 0.2'),
            'solvabilitate': ('Solvabilitate (%)', 'Ideal > 50%'),
            'indatorare': ('Indatorare (%)', 'Ideal < 50%'),
            'autonomie_financiara': ('Autonomie Financiara (%)', 'Ideal > 50%')
        }

        for key, (label, interpretation) in ratio_labels.items():
            val = metrics['ratios'].get(key)
            val_str = str(val) if val is not None else 'N/A'
            dashboard_data.append([label, val_str, interpretation])

        dashboard_data.append(['', '', ''])

        # Asset structure section
        dashboard_data.append(['STRUCTURA ACTIVELOR', '', ''])
        dashboard_data.append(['Component', 'Valoare', 'Procent'])
        for item in metrics['structure']['assets']:
            dashboard_data.append([item['name'], item['value'], f"{item['percent']}%"])

        dashboard_data.append(['', '', ''])

        # Liability structure section
        dashboard_data.append(['STRUCTURA PASIVELOR', '', ''])
        dashboard_data.append(['Component', 'Valoare', 'Procent'])
        for item in metrics['structure']['liabilities']:
            dashboard_data.append([item['name'], item['value'], f"{item['percent']}%"])

        df_dashboard = pd.DataFrame(dashboard_data, columns=['A', 'B', 'C'])
        df_dashboard.to_excel(writer, sheet_name='Dashboard', index=False, header=False)

        # Format Dashboard sheet
        workbook = writer.book
        worksheet = writer.sheets['Dashboard']

        # Set column widths
        worksheet.column_dimensions['A'].width = 30
        worksheet.column_dimensions['B'].width = 20
        worksheet.column_dimensions['C'].width = 20

    output.seek(0)
    return output


# =============================================================================
# BI ANALYTICS
# =============================================================================

def calculate_bi_metrics(df_bilant):
    """
    Calculate Business Intelligence metrics from processed Bilant.
    Returns key financial ratios and structure data for dashboard.
    """
    # Build nr_rd to value mapping
    nr_rd_map = {}
    for i, row in df_bilant.iterrows():
        nr = str(row.iloc[1]).replace('.0', '') if pd.notna(row.iloc[1]) else ''
        if nr:
            val = row.iloc[2] if pd.notna(row.iloc[2]) else 0
            nr_rd_map[nr] = val

    # Key balance sheet positions (based on Romanian Bilant structure)
    # Assets
    active_imobilizate = nr_rd_map.get('25', 0)  # TOTAL ACTIVE IMOBILIZATE
    active_circulante = nr_rd_map.get('40', 0)   # TOTAL ACTIVE CIRCULANTE
    stocuri = nr_rd_map.get('30', 0)             # TOTAL Stocuri
    creante = nr_rd_map.get('37', 0)             # TOTAL Creante
    disponibilitati = nr_rd_map.get('39', 0)     # Casa si conturi la banci
    total_active = nr_rd_map.get('41', 0)        # TOTAL ACTIVE

    # Liabilities
    datorii_termen_scurt = nr_rd_map.get('54', 0)  # Datorii < 1 an
    datorii_termen_lung = nr_rd_map.get('55', 0)   # Datorii > 1 an
    total_datorii = datorii_termen_scurt + datorii_termen_lung

    # Equity
    capitaluri_proprii = nr_rd_map.get('101', 0)   # CAPITALURI PROPRII TOTAL
    capital_social = nr_rd_map.get('81', 0)       # Capital subscris varsat

    # Calculate ratios
    metrics = {
        'summary': {
            'total_active': total_active,
            'active_imobilizate': active_imobilizate,
            'active_circulante': active_circulante,
            'capitaluri_proprii': capitaluri_proprii,
            'total_datorii': total_datorii
        },
        'ratios': {},
        'structure': {
            'assets': [],
            'liabilities': []
        }
    }

    # Financial ratios
    # 1. Lichiditate curenta (Current Ratio) = Active Circulante / Datorii < 1 an
    if datorii_termen_scurt > 0:
        metrics['ratios']['lichiditate_curenta'] = round(active_circulante / datorii_termen_scurt, 2)
    else:
        metrics['ratios']['lichiditate_curenta'] = None

    # 2. Lichiditate rapida (Quick Ratio) = (Active Circulante - Stocuri) / Datorii < 1 an
    if datorii_termen_scurt > 0:
        metrics['ratios']['lichiditate_rapida'] = round((active_circulante - stocuri) / datorii_termen_scurt, 2)
    else:
        metrics['ratios']['lichiditate_rapida'] = None

    # 3. Lichiditate imediata (Cash Ratio) = Disponibilitati / Datorii < 1 an
    if datorii_termen_scurt > 0:
        metrics['ratios']['lichiditate_imediata'] = round(disponibilitati / datorii_termen_scurt, 2)
    else:
        metrics['ratios']['lichiditate_imediata'] = None

    # 4. Rata solvabilitatii (Solvency) = Capitaluri Proprii / Total Active
    if total_active > 0:
        metrics['ratios']['solvabilitate'] = round(capitaluri_proprii / total_active * 100, 1)
    else:
        metrics['ratios']['solvabilitate'] = None

    # 5. Rata indatorarii (Debt Ratio) = Total Datorii / Total Active
    if total_active > 0:
        metrics['ratios']['indatorare'] = round(total_datorii / total_active * 100, 1)
    else:
        metrics['ratios']['indatorare'] = None

    # 6. Autonomie financiara = Capitaluri Proprii / (Capitaluri Proprii + Datorii)
    total_pasive = capitaluri_proprii + total_datorii
    if total_pasive > 0:
        metrics['ratios']['autonomie_financiara'] = round(capitaluri_proprii / total_pasive * 100, 1)
    else:
        metrics['ratios']['autonomie_financiara'] = None

    # Asset structure for charts
    if total_active > 0:
        metrics['structure']['assets'] = [
            {'name': 'Active Imobilizate', 'value': active_imobilizate, 'percent': round(active_imobilizate / total_active * 100, 1)},
            {'name': 'Stocuri', 'value': stocuri, 'percent': round(stocuri / total_active * 100, 1)},
            {'name': 'Creante', 'value': creante, 'percent': round(creante / total_active * 100, 1)},
            {'name': 'Disponibilitati', 'value': disponibilitati, 'percent': round(disponibilitati / total_active * 100, 1)}
        ]

    # Liability structure for charts
    if total_pasive > 0:
        metrics['structure']['liabilities'] = [
            {'name': 'Capitaluri Proprii', 'value': capitaluri_proprii, 'percent': round(capitaluri_proprii / total_pasive * 100, 1)},
            {'name': 'Datorii < 1 an', 'value': datorii_termen_scurt, 'percent': round(datorii_termen_scurt / total_pasive * 100, 1)},
            {'name': 'Datorii > 1 an', 'value': datorii_termen_lung, 'percent': round(datorii_termen_lung / total_pasive * 100, 1)}
        ]

    return metrics


# =============================================================================
# FLASK ROUTES
# =============================================================================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Invalid file type. Please upload an Excel file.'}), 400

    try:
        # Read the Excel file
        xlsx = pd.ExcelFile(file)

        # Check for required sheets
        if 'Balanta' not in xlsx.sheet_names:
            return jsonify({'error': 'Sheet "Balanta" not found in the file'}), 400
        if 'Bilant' not in xlsx.sheet_names:
            return jsonify({'error': 'Sheet "Bilant" not found in the file'}), 400

        # Read sheets
        df_balanta = pd.read_excel(xlsx, sheet_name='Balanta')
        df_bilant = pd.read_excel(xlsx, sheet_name='Bilant')

        # Process
        df_balanta_processed, df_bilant_processed = process_bilant(df_balanta, df_bilant)

        # Create output file
        output = create_output_excel(df_balanta_processed, df_bilant_processed)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Bilant_Generated.xlsx'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/health')
def health():
    return jsonify({'status': 'healthy'})


@app.route('/analyze', methods=['POST'])
def analyze_file():
    """Process file and return BI metrics for dashboard."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Invalid file type. Please upload an Excel file.'}), 400

    try:
        # Read the Excel file
        xlsx = pd.ExcelFile(file)

        # Check for required sheets
        if 'Balanta' not in xlsx.sheet_names:
            return jsonify({'error': 'Sheet "Balanta" not found in the file'}), 400
        if 'Bilant' not in xlsx.sheet_names:
            return jsonify({'error': 'Sheet "Bilant" not found in the file'}), 400

        # Read sheets
        df_balanta = pd.read_excel(xlsx, sheet_name='Balanta')
        df_bilant = pd.read_excel(xlsx, sheet_name='Bilant')

        # Process
        df_balanta_processed, df_bilant_processed = process_bilant(df_balanta, df_bilant)

        # Calculate BI metrics
        metrics = calculate_bi_metrics(df_bilant_processed)

        return jsonify(metrics)

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/template')
def download_template():
    """Download the template Excel file."""
    template_path = os.path.join(app.static_folder, 'template_balanta.xlsx')
    if os.path.exists(template_path):
        return send_file(
            template_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Template_Balanta_Bilant.xlsx'
        )
    else:
        return jsonify({'error': 'Template file not found'}), 404


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
