import PyPDF2
import pandas as pd
import numpy as np
import re

# Define the common terms that may be used in the headers for invoice numbers, amounts, CUIN numbers, and remarks
invoice_terms = ['Invoice', 'Number', 'Referance No', 'Doc/Inv # ', 'INV. NO:-', 'INVOICE NO']
amount_terms = ['Invoice amount', 'Original', 'Amount Paid (Ksh)', 'Net Pay', 'TO SUPPLIER', 'AMOUNT', 'CFP AMT']
cuin_terms = ['CUIN Number','Suplr Inv # ', 'INVOICE #']
remark_terms = ['Remark', 'Remarks', 'REMARKS', 'REMARK']
date_terms = ['Date', 'DATE']
start_terms = ['Date', 'Amount']

# Function to find the correct column based on exact match of possible terms
def find_column(headers, possible_terms):
    for term in possible_terms:
        for header in headers:
            if term.lower() == header.lower():
                return header
    return None

# Function to find the starting row based on possible terms
def find_starting_row(df, possible_terms):
    for index, row in df.iterrows():
        for cell in row:
            if any(term.lower() in str(cell).lower() for term in possible_terms):
                return index
    return 0

# Function to clean invoice numbers (remove special characters, remove rows with letters)
def clean_invoice_number(invoice_value):
    # Remove all non-numeric and non-alphabetic characters
    cleaned_invoice = re.sub(r'[^a-zA-Z0-9]', '', invoice_value)
    
    # Check if the cleaned invoice number contains any letters; if so, return None (to be removed later)
    if any(c.isalpha() for c in cleaned_invoice):
        return None
    return cleaned_invoice

# Function to clean remittance advice for Quick Mart Limited
def clean_quick_mart_remittance(file_path):
    df = pd.read_excel(file_path, header=None)

    # Select columns based on fixed positions
    date_column = df.iloc[:, 0]  # Column A
    invoice_column = df.iloc[:, 3]  # Column D
    cuin_column = df.iloc[:, 4]  # Column E
    amount_column = df.iloc[:, 7]  # Column H

    # Create cleaned DataFrame
    cleaned_df = pd.DataFrame(columns=['Date', 'Invoice Number', 'CUIN', 'Amount'])

    last_valid_invoice = None
    last_valid_cuin = None

    rows = []  # Use a list to accumulate rows for the DataFrame

    for i in range(len(invoice_column)):
        date_value = date_column.iloc[i]
        invoice_value = str(invoice_column.iloc[i])
        cuin_value = str(cuin_column.iloc[i])
        amount_value = amount_column.iloc[i]

        # Clean the invoice number by removing special characters and invalid rows
        cleaned_invoice = clean_invoice_number(invoice_value)

        # Check if CUIN contains any special character or letter, if so, skip the row
        if re.search(r'[a-zA-Z]', cuin_value):
            continue

        if cleaned_invoice:
            # Positive amount for valid rows
            rows.append({
                'Date': date_value,
                'Invoice Number': cleaned_invoice,
                'CUIN': cuin_value,
                'Amount': f"{abs(float(amount_value)):.2f}" if pd.notna(amount_value) else np.nan  # Ensure positive amounts
            })
            last_valid_invoice = cleaned_invoice
            last_valid_cuin = cuin_value
        else:
            # If the invoice number is invalid, check the CUIN number for specific terms
            if "CREDIT NOTE B" in cuin_value.upper() or "IMPART TO CREDITOR" in cuin_value.upper():
                # Negative amount for specific rows
                rows.append({
                    'Date': date_value,
                    'Invoice Number': last_valid_invoice,
                    'CUIN': last_valid_cuin,
                    'Amount': f"{-abs(float(amount_value)):.2f}" if pd.notna(amount_value) else np.nan
                })

    # Convert the list of rows into a DataFrame
    cleaned_df = pd.DataFrame(rows)

    return cleaned_df

# Function to clean remittance advice for Chandarana Supermarket
def clean_chandarana_remittance(file_path):
    df = pd.read_excel(file_path, header=None)
    start_row = find_starting_row(df, start_terms)
    df = pd.read_excel(file_path, skiprows=start_row)

    headers = df.columns
    invoice_column_name = find_column(headers, invoice_terms)
    amount_column_name = find_column(headers, amount_terms)
    cuin_column_name = find_column(headers, cuin_terms)
    remark_column_name = find_column(headers, remark_terms)
    date_column_name = find_column(headers, date_terms)

    # If CUIN column is not found, initialize it with empty values
    if cuin_column_name is None:
        df['CUIN'] = pd.NA
        cuin_column_name = 'CUIN'

    # Create cleaned DataFrame
    cleaned_df = pd.DataFrame()
    cleaned_df['Invoice Number'] = df[invoice_column_name]
    cleaned_df['Amount'] = df[amount_column_name]
    cleaned_df['CUIN'] = df[cuin_column_name]
    cleaned_df['Remark'] = df[remark_column_name]
    cleaned_df['Date'] = df[date_column_name]

    # List to accumulate rows
    rows = []
    last_valid_invoice = None
    last_valid_cuin = None
    last_valid_date = None

    # Flags to identify "less credits" and "to pay" ranges
    in_range = False

    for i in range(len(cleaned_df)):
        invoice_value = cleaned_df['Invoice Number'].iloc[i]
        amount_value = cleaned_df['Amount'].iloc[i]
        cuin_value = cleaned_df['CUIN'].iloc[i]
        remark_value = cleaned_df['Remark'].iloc[i]
        date_value = cleaned_df['Date'].iloc[i]

        # Handle the range between "less credits" or "less returns" and "to pay"
        if "less credits" in str(invoice_value).lower() or "less returns" in str(invoice_value).lower():
            in_range = True

        if in_range and pd.notna(amount_value):
            rows.append({
                'Invoice Number': invoice_value,
                'Amount': amount_value,
                'CUIN': cuin_value,
                'Date': date_value
            })

        if "to pay" in str(invoice_value).lower():
            in_range = False

        # Retain valid rows with invoice numbers and not in the "less credits" to "to pay" range or "less returns" to "to pay" range
        if pd.notna(invoice_value) and "less credits" not in str(invoice_value).lower() and "to pay" not in str(invoice_value).lower() or pd.notna(invoice_value) and "less returns" not in str(invoice_value).lower() and "to pay" not in str(invoice_value).lower():
            last_valid_invoice = invoice_value
            last_valid_cuin = cuin_value
            last_valid_date = date_value
            rows.append({
                'Invoice Number': invoice_value,
                'Amount': amount_value,
                'CUIN': cuin_value,
                'Date': date_value
            })
        # Process rows with 'ss' in remark and a valid amount but no invoice number
        elif pd.isna(invoice_value) and pd.notna(amount_value) and remark_value == 'ss':
            if last_valid_invoice is not None and last_valid_cuin is not None:
                # Add the amount of this row to the last valid invoice above it
                rows[-1]['Amount'] = rows[-1]['Amount'] + amount_value

                # Retain this row by applying the last valid invoice, CUIN, and Date
                rows.append({
                    'Invoice Number': last_valid_invoice,
                    'Amount': amount_value,
                    'CUIN': last_valid_cuin,
                    'Date': last_valid_date
                })
        elif pd.isna(invoice_value) and pd.notna(amount_value) and remark_value != 'ss' and not in_range:
            # Skip rows with amount but no invoice number and no 'ss' in remark, unless within the "less credits" to "to pay" range
            continue

    # Convert the list of rows into a DataFrame
    cleaned_result_df = pd.DataFrame(rows)

    # Remove rows containing 'less credits', 'to pay' and the row before 'to pay'
    drop_indexes = []
    for i in range(len(cleaned_result_df)):
        if "to pay" in str(cleaned_result_df['Invoice Number'].iloc[i]).lower():
            drop_indexes.append(i - 1 if i > 0 else i)  # Drop the row before 'to pay'
            drop_indexes.append(i)  # Drop the 'to pay' row
        elif "less credits" in str(cleaned_result_df['Invoice Number'].iloc[i]).lower():
            drop_indexes.append(i)  # Drop the 'less credits' row
        elif "less returns" in str(cleaned_result_df['Invoice Number'].iloc[i]).lower():
            drop_indexes.append(i)  # Drop the 'less returns' row

    cleaned_result_df = cleaned_result_df.drop(drop_indexes).reset_index(drop=True)

    return cleaned_result_df

# Function to clean remittance advice for Majid Al Futaim Hypermarkets Ltd
def clean_majid_al_futaim_remittance(file_path):
    # Extract text from the PDF file
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ''
        for page in reader.pages:
            text += page.extract_text()

    # Define regex pattern for CUIN, date (dd-mm-yy), remark (optional), and amount with optional trailing text
    pattern = r"([\d\w/]+)\s*(\d{2}-\d{2}-\d{2})\s*(.*?)\s([\d,]+\.\d{2}-?)"

    # Find all occurrences of the pattern
    matches = re.findall(pattern, text, re.MULTILINE)

    # Create lists to store extracted data
    cuins = []
    dates = []
    remarks = []
    amounts = []

    for match in matches:
        cuin, date, remark, amount = match

        # Clean and convert amount, moving the hyphen to the front if present
        amount = str(amount.replace(',', ''))
        if amount.endswith('-'):
            amount = '-' + amount[:-1]

        # Append extracted data to lists
        cuins.append(cuin)
        dates.append(date)
        remarks.append(remark.strip())  # Remove any extra whitespace
        amounts.append(amount)

    # Create a DataFrame
    data = {
        'Invoice Number': [''] * len(cuins),  # Empty Invoice Number column
        'CUIN': cuins,
        'Date': dates,
        'Remark': remarks,
        'Amount': amounts
    }
    df = pd.DataFrame(data)

    return df

# Default cleaning logic
def clean_remittance_advice(file_path, specific_logic=False):
    if specific_logic:
        return clean_quick_mart_remittance(file_path)

    df = pd.read_excel(file_path, header=None)
    start_row = find_starting_row(df, start_terms)
    df = pd.read_excel(file_path, skiprows=start_row)

    headers = df.columns
    invoice_column = find_column(headers, invoice_terms)
    amount_column = find_column(headers, amount_terms)
    cuin_column = find_column(headers, cuin_terms)
    date_column = find_column(headers, date_terms)

    # Create cleaned DataFrame with necessary columns, initializing missing ones with pd.NA
    cleaned_df = pd.DataFrame()
    cleaned_df['Invoice Number'] = df[invoice_column] if invoice_column else pd.NA
    cleaned_df['CUIN'] = df[cuin_column] if cuin_column else pd.NA
    cleaned_df['Amount'] = df[amount_column] if amount_column else pd.NA
    cleaned_df['Date'] = df[date_column] if date_column else pd.NA

    if amount_column:
        # Remove rows based on specific conditions
        cleaned_df = cleaned_df[~cleaned_df['Invoice Number'].astype(str).str.contains('INVOICE #', case=False, na=False)]
        cleaned_df = cleaned_df[~cleaned_df['Invoice Number'].astype(str).str.contains('less credits ', case=False, na=False)]
        cleaned_df = cleaned_df[~cleaned_df['Invoice Number'].astype(str).str.contains('to pay ', case=False, na=False)]
        cleaned_df = cleaned_df[~cleaned_df['CUIN'].astype(str).str.contains('INVOICE #', case=False, na=False)]

        # Find index of rows containing "Payment Date" or "Total"
        payment_date_index = cleaned_df[cleaned_df['Invoice Number'].astype(str).str.contains('Payment Date', case=False, na=False)].index
        total_index = cleaned_df[cleaned_df['Invoice Number'].astype(str).str.contains('Total', case=False, na=False)].index

        # Remove rows containing "Payment Date" or "Total" and all rows after them
        if not payment_date_index.empty:
            cleaned_df = cleaned_df[:payment_date_index[0]]
        elif not total_index.empty:
            cleaned_df = cleaned_df[:total_index[0]]

        # Trim the DataFrame at the row where either "Invoice Number" and "CUIN Number" columns are empty
        cleaned_df = cleaned_df[(cleaned_df['Invoice Number'].notna()) | (cleaned_df['CUIN'].notna())]

        # Remove commas from 'Amount' column, convert to numeric, round to 2 decimal places, and convert back to string
        cleaned_df['Amount'] = cleaned_df['Amount'].astype(str).str.replace(',', '').astype(float).round(2).apply(lambda x: f"{x:.2f}") #If amount is negative, leave as negative. If positive, leave as positive.

        return cleaned_df
    else:
        raise ValueError(f"Could not find required columns in {file_path}")

def main(input_file_path, output_file_path):
    try:
        # Check if the input file path contains "Quick Mart Limited" or "Chandarana Supermarket"
        if "Quick Mart Limited" in input_file_path:
            cleaned_df = clean_quick_mart_remittance(input_file_path)
        elif "Chandarana Supermarket" in input_file_path:
            cleaned_df = clean_chandarana_remittance(input_file_path)
        elif "Majid Al Futaim Hypermarkets Ltd" in input_file_path:
            cleaned_df = clean_majid_al_futaim_remittance(input_file_path)
        else:
            cleaned_df = clean_remittance_advice(input_file_path)

        cleaned_df.to_excel(output_file_path, index=False)
        print(f"Processed and saved cleaned data to {output_file_path}")
    except ValueError as e:
        print(f"Error processing file {input_file_path}: {e}")

# Example usage within UiPath
# main(r"C:\Users\jeff.nganga\OneDrive - PWANI OIL PRODUCTS LIMITED\Documents\Pwani RPA\Projects Documentation\Receipt Allocation\Naivas LTD\Naivas Remittance.xlsx", r"C:\Users\jeff.nganga\OneDrive - PWANI OIL PRODUCTS LIMITED\Documents\Pwani RPA\Projects Documentation\Receipt Allocation\Naivas LTD\Cleaned Naivas Remittance.xlsx")