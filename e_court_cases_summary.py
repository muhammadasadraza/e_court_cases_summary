import os
import pandas as pd
from PyPDF2 import PdfReader
from pathlib import Path
from collections import defaultdict

def count_pdf_pages(pdf_path):
    """Count the number of pages in a PDF file."""
    try:
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            return len(reader.pages)
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
        return None

def scan_directory(base_dir):
    """Scan the directory and collect PDF data from all subdirectories, logging duplicates."""
    pdf_data = []
    failed_files = []
    file_tracker = defaultdict(list)  # Track duplicates by Case Number and Date

    base_path = Path(base_dir)
    print(f"Scanning base directory: {base_path}")

    for root, dirs, files in os.walk(base_path):
        root_path = Path(root)
        path_parts = root_path.parts

        if len(path_parts) > 1:
            date_folder = path_parts[1]
            extra_folders = path_parts[2:]
            
            for filename in files:
                if filename.endswith('.pdf'):
                    pdf_path = root_path / filename
                    case_number = filename.split(', ')[-1].strip('.pdf')
                    num_pages = count_pdf_pages(pdf_path)
                    
                    if num_pages is not None:
                        data_entry = {
                            'Case Number': case_number,
                            'Number of Pages': num_pages,
                            'Date': date_folder
                        }

                        # Add columns for each extra subdirectory level
                        for i, folder in enumerate(extra_folders, start=1):
                            data_entry[f'Sub-Folder Level {i}'] = folder

                        # Track duplicates
                        key = (case_number, date_folder)
                        if key in file_tracker:
                            print(f"Duplicate found for Case Number: {case_number} on Date: {date_folder}")
                        file_tracker[key].append(pdf_path)  # Log path of duplicates
                        
                        pdf_data.append(data_entry)
                    else:
                        failed_files.append(pdf_path)

    print(f"Total PDFs found and processed: {len(pdf_data)}")
    print(f"Total PDFs that failed to process: {len(failed_files)}")
    print(f"Total duplicates detected: {sum(len(paths) - 1 for paths in file_tracker.values() if len(paths) > 1)}")
    
    if failed_files:
        print("Files that failed to process:", failed_files)
    
    return pdf_data

def save_to_excel(pdf_data, output_file):
    """Save the collected PDF data to an Excel file, with dynamic columns for subdirectories."""
    if pdf_data:
        df = pd.DataFrame(pdf_data)
        df.to_excel(output_file, index=False)
        print(f"Data saved to {output_file}")
    else:
        print("No PDF data found to save.")

def main():
    base_dir = 'D:/Scanned_Files'
    output_file = 'e_court_cases_summary.xlsx'
    
    pdf_data = scan_directory(base_dir)
    save_to_excel(pdf_data, output_file)

if __name__ == "__main__":
    main()