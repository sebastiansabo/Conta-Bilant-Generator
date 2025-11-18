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

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = 'uploads'

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# =============================================================================
# CONSTANTS - Match VBA macro layout
# =============================================================================

# Balanta layout (0-indexed for pandas)
COL_BAL_ACCOUNT = 1   # B = Nr. cont (RAD1)
COL_BAL_SFD = 4       # E = SFD
COL_BAL_SFC = 5       # F = SFC

# Bilant layout
COL_BIL_DESC = 0      # A = Denumirea elementului
COL_BIL_NR_RD = 2     # C = Nr. rd.
COL_BIL_VAL = 4       # E = Sold Final
COL_BIL_FORM_CT = 5   # F = Formula Calcul
COL_BIL_FORM_RD = 6   # G = Formula Randuri


# =============================================================================
# STEP 1: Calculate Sold Final in Balanta
# =============================================================================

def calculate_sold_final(df_balanta):
    """
    Calculate Sold Final for each account in Balanta.
    Sold Final = SFD + SFC (absolute sum for display purposes)
    Also calculates Net Balance = SFD - SFC for accounting purposes
    """
    df = df_balanta.copy()

    # Skip header row if present
    if df.iloc[0, COL_BAL_ACCOUNT] == 'Cont':
        df = df.iloc[1:].reset_index(drop=True)

    sold_final = []
    net_balance = []

    for idx, row in df.iterrows():
        sfd = pd.to_numeric(row.iloc[COL_BAL_SFD], errors='coerce')
        sfc = pd.to_numeric(row.iloc[COL_BAL_SFC], errors='coerce')

        sfd = 0 if pd.isna(sfd) else sfd
        sfc = 0 if pd.isna(sfc) else sfc

        # Sold Final as absolute sum (for prefix matching)
        sf = abs(sfd) + abs(sfc)
        sold_final.append(sf)

        # Net balance for actual accounting
        net = sfd - sfc
        net_balance.append(net)

    df['Sold_Final'] = sold_final
    df['Net_Balance'] = net_balance

    return df


# =============================================================================
# STEP 2: Extract CT formulas from Bilant descriptions
# =============================================================================

def extract_ct_formula(description):
    """
    Extract account formula from description text.
    Example: "1.Cheltuieli de constituire (ct.201-2801)" -> "201-2801"
    """
    if pd.isna(description):
        return ""

    text = str(description)

    # Find "ct." or "ct " followed by formula
    match = re.search(r'ct\.?\s*([^)]+)', text, re.IGNORECASE)
    if not match:
        return ""

    expr = match.group(1).strip()

    # Clean up the expression
    expr = re.sub(r'\s+', '', expr)  # Remove whitespace
    expr = expr.replace('*', '')     # Remove asterisks

    return expr


# =============================================================================
# STEP 3: Extract row formulas from Bilant descriptions
# =============================================================================

def extract_row_formula(description):
    """
    Extract row formula from description text.
    Example: "TOTAL (rd. 01 la 06)" -> "01+02+03+04+05+06"
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

    # Handle "01 la 06" format
    la_match = re.search(r'(\d+)la(\d+)', raw)
    if la_match:
        start = int(la_match.group(1))
        end = int(la_match.group(2))
        width = len(la_match.group(1))

        if end >= start:
            parts = [str(i).zfill(width) for i in range(start, end + 1)]
            return '+'.join(parts)

    # Otherwise extract numbers and signs
    result = re.sub(r'[^0-9+\-]', '', raw)
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
        if acct.startswith(prefix):
            if use_net:
                # For dynamic +/- terms: use SFD - SFC
                sfd = pd.to_numeric(row.iloc[COL_BAL_SFD], errors='coerce') or 0
                sfc = pd.to_numeric(row.iloc[COL_BAL_SFC], errors='coerce') or 0
                val = sfd - sfc
            else:
                # For normal terms: use Sold Final (SFD + SFC)
                val = row.get('Sold_Final', 0) or 0

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
            for acct, val in details:
                all_details.append((acct, val, prefix, 'dynamic'))
            total += subtotal
        elif sign_type == 'normal_plus':
            subtotal, details = sum_accounts_by_prefix(df_balanta, prefix, use_net=False)
            for acct, val in details:
                all_details.append((acct, val, prefix, '+'))
            total += subtotal
        elif sign_type == 'normal_minus':
            subtotal, details = sum_accounts_by_prefix(df_balanta, prefix, use_net=False)
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
    """
    if not expr:
        return 0

    total = 0
    sign = 1
    num = ''

    for ch in expr + '+':  # Add + to flush last number
        if ch.isdigit():
            num += ch
        elif ch in '+-':
            if num:
                row_num = num.lstrip('0') or '0'
                val = bilant_values.get(row_num, 0)
                total += sign * val
                num = ''
            sign = 1 if ch == '+' else -1

    return total


# =============================================================================
# MAIN PROCESSING FUNCTION
# =============================================================================

def process_bilant(df_balanta, df_bilant):
    """
    Process Balanta and generate Bilant with calculations and verification.
    """
    # Step 1: Calculate Sold Final in Balanta
    df_balanta = calculate_sold_final(df_balanta)

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

    # Add results to dataframe
    df_bilant['Calculated_Value'] = results
    df_bilant['Verification'] = verifications

    return df_balanta, df_bilant


# =============================================================================
# EXCEL OUTPUT
# =============================================================================

def create_output_excel(df_balanta, df_bilant):
    """
    Create output Excel file with processed Balanta and Bilant.
    """
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Write Balanta with Sold Final
        df_balanta.to_excel(writer, sheet_name='Balanta', index=False)

        # Write Bilant with calculations
        df_bilant.to_excel(writer, sheet_name='Bilant', index=False)

    output.seek(0)
    return output


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


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
