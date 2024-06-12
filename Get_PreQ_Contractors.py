import fitz  # PyMuPDF
import re
import pandas as pd
from datetime import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename

# List of US state abbreviations
STATE_ABBREVIATIONS = {
    'AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA', 'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 
    'MD', 'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ', 'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 
    'RI', 'SC', 'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY'
}

# Function to standardize phone and fax numbers
def standardize_phone_fax(number):
    number = re.sub(r'\D', '', number)  # Remove non-digit characters
    if len(number) == 10:
        return f'({number[:3]}){number[3:6]}-{number[6:]}'
    else:
        return "N/A"

# Function to check if text is bold
def is_bold(text_item):
    return 'bold' in text_item['font'].lower()

# Function to process a page text
def process_page(page, contractor_data):
    blocks = page.get_text("dict")["blocks"]
    contractor = {}
    previous_bold_text = None
    for block in blocks:
        for line in block["lines"]:
            for span in line["spans"]:
                text = span["text"].strip()
                if "TDOT Prequalified Contractors As Of" in text or "Contractor" in text or "See last page of report" in text:
                    continue  # Ignore title, headers, and footers
                if text in ["Mailing Address", "Phone", "State", "City", "Fax", "Zip"]:
                    continue  # Ignore unnecessary headers

                if is_bold(span) and text:
                    if previous_bold_text:
                        # If the previous bold text exists, it means the name continued on this line
                        contractor["Contractor"] += " " + text
                        previous_bold_text = None
                    else:
                        if contractor:  # If there is already a contractor being processed, save it first
                            contractor_data.append(contractor)
                            contractor = {}  # Reset for the next contractor
                        contractor["Contractor"] = text
                        previous_bold_text = text
                else:
                    previous_bold_text = None
                    if "Contractor" in contractor:
                        # Capture details based on known patterns
                        phone_match = re.search(r'Phone:\s*([\(\)\-\d\s]+)', text)
                        fax_match = re.search(r'Fax:\s*([\(\)\-\d\s\*]+)', text)
                        expiration_date_match = re.search(r'Expiration Date:\s*(\d{2}/\d{2}/\d{4})', text)
                        vendor_id_match = re.search(r'Vendor ID:\s*(\d+)', text)
                        address_match = re.search(r'(\d+\s+.*?)\s{2,}', text)
                        city_state_zip_match = re.search(r'\s{2,}(.+?),\s*(\w{2})\s*(\d{5})', text)
                        
                        if expiration_date_match:
                            contractor["Expiration Date"] = expiration_date_match.group(1).strip()
                        if vendor_id_match:
                            contractor["Vendor ID"] = vendor_id_match.group(1).strip()
                        if phone_match:
                            contractor["Phone"] = standardize_phone_fax(phone_match.group(1).strip())
                        if fax_match:
                            contractor["Fax"] = standardize_phone_fax(fax_match.group(1).strip())
                        if address_match:
                            contractor["Mailing Address"] = address_match.group(1).strip()
                        if city_state_zip_match:
                            contractor["City"] = city_state_zip_match.group(1).strip()
                            contractor["State"] = city_state_zip_match.group(2).strip()
                            contractor["Zip"] = city_state_zip_match.group(3).strip()
                        if "* NO FAX *" in text:
                            contractor["Fax"] = "N/A"
                        if "Certified SBE" in text:
                            contractor["Certified SBE"] = "Yes"
                        if "Certified DBE" in text:
                            contractor["Certified DBE"] = "Yes"
                        if "Limited Prequalification" in text:
                            contractor["Limited Prequalification"] = "Yes"
                        else:
                            contractor["Certified SBE"] = contractor.get("Certified SBE", "No")
                            contractor["Certified DBE"] = contractor.get("Certified DBE", "No")
                            contractor["Limited Prequalification"] = contractor.get("Limited Prequalification", "No")
                        
                        work_classes = ["ASPH", "BASE", "CONC", "ENGR", "ERTH", "FNCE", "HAUL", "ITS", "LITE", "NONR", "RIPR", "RR", "SGNL", 
                                        "SLLE", "STBR", "SWPD", "TRFT", "UTIL", "BARR", "BRPT", "DRNG", "EROS", "FLAT", "GAT", "INCI", 
                                        "LAND", "MOVL", "PVWK", "RMVL", "SALT", "SLIP", "SNOW", "STDR", "TRFP", "TUNN", "WARW"]
                        work_classes_block = re.search(r'Work Class:\s*(.*)', text)
                        if work_classes_block:
                            work_classes_text = work_classes_block.group(1).strip()
                            for wc in work_classes:
                                contractor[wc] = "Yes" if wc in work_classes_text else "No"
                        else:
                            for wc in work_classes:
                                contractor[wc] = "No"

    if contractor:  # Append the last contractor if there is one
        contractor_data.append(contractor)

def main():
    # Create a Tkinter root window
    root = Tk()
    root.withdraw()  # Hide the root window

    # Prompt user to select a PDF file
    pdf_path = askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])
    if not pdf_path:
        print("No file selected. Exiting.")
        return

    # Step 1: Open PDF and extract text
    pdf_document = fitz.open(pdf_path)

    # Extract initial "As Of" date from the first page
    first_page_text = pdf_document[0].get_text()
    as_of_date_match = re.search(r'As Of\s*(\w+ \d{1,2}, \d{4})', first_page_text)
    if as_of_date_match:
        as_of_date = datetime.strptime(as_of_date_match.group(1), '%B %d, %Y').strftime('%m/%d/%Y')
    else:
        print("Failed to extract 'As Of' date. Exiting.")
        return

    # Initialize a list to hold contractor data
    contractor_data = []

    # Step 2: Loop through all pages except the last one
    for page_num in range(len(pdf_document) - 1):
        page = pdf_document.load_page(page_num)
        process_page(page, contractor_data)

    # Step 3: Create DataFrame and clean data
    work_classes = ["ASPH", "BASE", "CONC", "ENGR", "ERTH", "FNCE", "HAUL", "ITS", "LITE", "NONR", "RIPR", "RR", "SGNL", 
                    "SLLE", "STBR", "SWPD", "TRFT", "UTIL", "BARR", "BRPT", "DRNG", "EROS", "FLAT", "GAT", "INCI", 
                    "LAND", "MOVL", "PVWK", "RMVL", "SALT", "SLIP", "SNOW", "STDR", "TRFP", "TUNN", "WARW"]
    columns = ["Contractor", "Expiration Date", "Vendor ID", "Mailing Address", "City", "State", "Zip", "Phone", "Fax",
               "Certified SBE", "Certified DBE", "Limited Prequalification"] + work_classes
    data_df = pd.DataFrame(contractor_data, columns=columns)

    # Step 4: Prompt user to save the CSV file
    output_file_name = f'TDOT_Prequalified_Contractors_As_Of_{as_of_date.replace("/", "_")}.csv'
    save_path = asksaveasfilename(defaultextension=".csv", initialfile=output_file_name, filetypes=[("CSV Files", "*.csv")])
    if save_path:
        data_df.to_csv(save_path, index=False)
        print(f"Data extraction complete. As Of Date: {as_of_date}")
        print(f"Total number of contractors: {len(contractor_data)}")
    else:
        print("Save operation cancelled. Exiting.")

if __name__ == "__main__":
    main()
