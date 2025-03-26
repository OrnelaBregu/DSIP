import pdfplumber
import fitz  # PyMuPDF
import pandas as pd
from openpyxl import load_workbook

def extract_main_text(pdf_path):
    """Extract text from the main content using pdfplumber."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text


def extract_widget_text(pdf_path):
    """Extract text from form fields/widgets using PyMuPDF."""
    doc = fitz.open(pdf_path)
    widget_text = ""
    for page in doc:
        widgets = page.widgets()
        if widgets:
            for widget in widgets:
                value = widget.field_value
                if value:
                    widget_text += f"{widget.field_name or ''}: {value}\n"
    return widget_text


def extract_fields(text):
    """
    Extract fields from given text using simple line-by-line parsing.
    Fields are expected to be in the format 'Field Name: Value'.
    """
    fields = {}
    for line in text.splitlines():
        line = line.strip()
        if ":" in line:
            key, value = line.split(":", 1)
            fields[key.strip()] = value.strip()
    return fields


def parse_pdf_data(pdf_path):
    """
    Extract text from both the main content and widget fields,
    then combine the extracted fields (with widget values taking precedence)
    to build the final data dictionary matching our Excel template.
    """
    # Extract texts from both sources
    main_text = extract_main_text(pdf_path)
    widget_text = extract_widget_text(pdf_path)

    # Extract fields dictionaries from each
    main_fields = extract_fields(main_text)
    widget_fields = extract_fields(widget_text)

    # Set up default values based on your Excel template
    default_data = {
        "Student Name": "Ornela",
        "Thesis Title": "LLM",
        "Student ID": "26898580",
        "Department": "CIISE",
        "TH05": "TH05",  # default value
        "Thesis Defence Date": "MArch 20, 2025",
        "Oral Defence": "Yes",
        "Room": "309",
        "Thesis Ranking": "Outstanding",
        "Coded on SIS": "",
        "Embargo Date": "",
        "Examining Committee decision": "Accepted"
    }

    # Define mapping between expected field names and keys in our default_data
    mapping = {
        "Student Name": "Student Name",
        "Thesis Title": "Thesis Title",
        "Student ID": "Student ID",
        "Department": "Department",
        "Defence date": "Thesis Defence Date",
        "Room": "Room",
        "Thesis ranking": "Thesis Ranking",
        "Decision": "Examining Committee decision",
        "Oral Defence": "Oral Defence",
        "Date": "Coded on SIS"
    }

    # Start with defaults, then override with main_fields, then with widget_fields (if available)
    combined_data = default_data.copy()

    # Helper function to update fields (case-insensitive matching)
    def update_fields(source_fields):
        for key, value in source_fields.items():
            for map_key, target_field in mapping.items():
                if key.lower() == map_key.lower() and value:
                    combined_data[target_field] = value

    update_fields(main_fields)
    update_fields(widget_fields)

    return combined_data


def save_to_excel(data, output_excel,template_path):
    """Save the extracted data into an Excel file using the specified template."""
    columns = [
        "Student Name", "Thesis Title", "Student ID", "Department",
        "TH05", "Thesis Defence Date", "Oral Defence", "Room"
        , "Coded on SIS", "Embargo Date", "Examining Committee decision","Thesis Ranking"
    ]
    df = pd.DataFrame([data], columns=columns)

    # Load the macro-enabled template workbook with macros preserved
    wb = load_workbook(template_path, keep_vba=True)

    # Create an ExcelWriter using the openpyxl engine with keep_vba=True
    writer = pd.ExcelWriter(output_excel, engine='openpyxl', keep_vba=True)
    writer.book = wb
    # Write the DataFrame to a specified sheet (e.g., "Sheet1")
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.save()
    writer.close()
    print("Data saved to:", output_excel)

    # df.to_excel(output_excel, index=False)
    # print("Data saved to:", output_excel)


if __name__ == "__main__":
    # Update these paths as needed
    pdf_path = r"C:\Users\umroot\Downloads\DSIP\MA Committee Report - Youssef Maghrebi_40259660.pdf"
    output_excel = r"C:\Users\umroot\Downloads\DSIP\ExtractPdfData.xlsm"
    template_path = r"C:\Users\umroot\Downloads\DSIP\Template.xlsm"
    data = parse_pdf_data(pdf_path)
    print("Extracted Data:", data)
    save_to_excel(data, output_excel,template_path)