# E-Court Case Summary Script

## Project Overview
This project is a Python-based script designed to automate the scanning of directories and subdirectories that contain E-Court session PDF case files. The script counts the pages in each PDF file and generates an organized Excel summary that includes essential details like case number, page count, date, and folder structure, which simplifies analysis and management of large case files.

## Features
- **Recursive Directory Scanning**: Scans all date-named folders and their nested subdirectories.
- **PDF Page Counting**: Accurately counts pages in each PDF file.
- **Excel Report Generation**: Exports data to an Excel file with columns for date, case number, page count, and the nested directory structure.
- **Duplicate Detection**: Logs duplicate case files to avoid redundancy in reporting.

## Directory Structure
Organize your files as follows, with the base directory containing folders named by dates, and each date folder containing case-type subfolders (Urgent, Regular, Fresh Regular, Supplementary, etc.):

```
D:/Scanned_Files
│
└─── 02-09-24/
     ├── Urgent/
     │   ├── 15. Crl. Misc. 6343-B-24.pdf
     │   └── ...
     ├── Regular/
     ├── Fresh Regular/
     └── Supplementary/
```

## Setup Instructions
### Prerequisites
- **Python 3.7 or higher**
- Required Python libraries:
  - `PyPDF2` for reading PDF files
  - `pandas` for handling data and exporting to Excel
  - `openpyxl` for Excel file creation

Install dependencies with:
```bash
pip install PyPDF2 pandas openpyxl
```

### How to Run
1. **Clone this Repository**
   ```bash
   git clone https://github.com/muhammadasadraza/e_court_cases_summary.git
   cd e_court_cases_summary
   ```

2. **Update the Base Directory**
   - Open `e_court_cases_summary.py` and set the `base_dir` variable to your directory path:
   ```python
   base_dir = r'D:/Scanned_Files'
   ```

3. **Execute the Script**
   - Run the script to generate the Excel report:
   ```bash
   python e_court_cases_summary.py
   ```

4. **Output**
   - The script creates an Excel file named `e_court_cases_summary.xlsx` containing columns:
     - **Case Number**: Extracted from each PDF filename.
     - **Number of Pages**: Total page count for each PDF.
     - **Date**: The name of the date folder.
     - **Sub-Folder Levels**: Each nested subfolder’s name.

### Sample Output
| Case Number     | Number of Pages | Date     | Sub-Folder Level 1 | Sub-Folder Level 2 |
|-----------------|-----------------|----------|---------------------|---------------------|
| 6343-B-24       | 12              | 02-09-24 | Urgent             | -                  |
| 1021-C-24       | 5               | 03-09-24 | Regular            | Supplementary      |

## Logging and Error Handling
- **Duplicate Files**: The script logs duplicate entries by `Case Number` and `Date`.
- **File Errors**: Files that can’t be processed are noted in the log for review.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contributing
Contributions are welcome! Feel free to open issues or submit pull requests with improvements.
