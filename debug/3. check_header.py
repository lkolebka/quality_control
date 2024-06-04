import os
from docx import Document

def read_headers(doc_path):
    print(f"Reading headers from document: {doc_path}")
    doc = Document(doc_path)
    header_texts = []

    for section in doc.sections:
        header = section.header
        header_content = ""
        for table in header.tables:
            for row in table.rows:
                for cell in row.cells:
                    header_content += cell.text + "\n"
        header_texts.append(header_content.strip())

    return header_texts

# Path to the directory containing the Word documents
doc_directory = r"C:\Users\lazare.kolebka\OneDrive - Accenture\Documents\Dev\Yara"

# Iterate over all Word documents in the directory
for filename in os.listdir(doc_directory):
    if filename.endswith(".docx"):
        doc_path = os.path.join(doc_directory, filename)
        headers = read_headers(doc_path)
        print(f"Headers in document {filename}:")
        for i, header in enumerate(headers):
            print(f"Header {i+1}:\n{header}\n")

print("Header reading completed.")
