import re
import PyPDF2
import pandas as pd

def extract_text_from_pdf(pdf_path):
    # Open the PDF file
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ''
        
        # Extract text from each page
        for page in reader.pages:
            text += page.extract_text()
    
    return text

def parse_pdf_text(text):
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

# Test the function
pdf_path = r"C:\Users\jeff.nganga\OneDrive - PWANI OIL PRODUCTS LIMITED\Documents\Pwani RPA\Projects Documentation\Receipt Allocation\Majid\Majid RA 01.08.24.pdf"
output_path = r"C:\Users\jeff.nganga\OneDrive - PWANI OIL PRODUCTS LIMITED\Documents\Pwani RPA\Projects Documentation\Receipt Allocation\Majid\Cleaned Majid RA 01.08.24.xlsx"
extracted_text = extract_text_from_pdf(pdf_path)
parsed_data = parse_pdf_text(extracted_text)

parsed_data.to_excel(output_path, index=False)