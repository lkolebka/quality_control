# Document Quality Checker

This project provides a tool to check the quality of Word documents (.docx) based on specific criteria. The script scans documents in a specified directory and generates a report in an Excel file.

## Features

- **1. Check for Blue Text**: Verifies if there is any blue-colored text in the document.
- **2. Check Headers**: Ensures all headers (Heading 1 and Heading 2) have valid content following them.
- **3. Extract Version Number**: Extracts the version number from the header of the document.
- **4. Check Approvers Table**: Validates the presence and completeness of the approvers table.
- **5. Check WRICEF Tables**: Validates WRICEF tables to ensure all IDs are correct integers.
- **6. Check Open Points Section**: Ensures the Open Points section contains a hyperlink.
- **7. Check Template Points**: Verifies all required template points are present in the document.

## Setup
### 1. Clone the Repository

```bash
git clone https://github.com/yourusername/document-quality-checker.git
cd document-quality-checker
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```
### 3. Directory Structure
Ensure your document are placed in the correct directory. For this script, place all ".docx" file in a directory **doc_directory** variable in [main.py](https://github.com/lkolebka/quality_control/main.py)

## Usage
### 1. Running the Script
Run the **main.py** script to start the quality check:

```bash
python main.py
```

### 2. Viewing the Report
After running the script, an Excel file named document_quality_report.xlsx will be generated in the specified directory. This file contains detailed information on the quality checks performed on each document.

### 3. Debug
In the [debug folder]([url](https://github.com/lkolebka/quality_control/tree/master/debug)), there are multipes sub-script you can exectute to see the execution of the different features. 



