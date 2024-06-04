import os
from docx import Document

def print_open_points_links(doc_path):
    print(f"_______________Checking Open Points links in document: {doc_path}")
    doc = Document(doc_path)
    open_points_found = False

    def has_hyperlink(paragraph):
        """
        Check if a paragraph contains a hyperlink.
        """
        for item in paragraph._element.xpath('.//w:hyperlink'):
            return True
        return False

    for paragraph in doc.paragraphs:
        if paragraph.text.strip().startswith("Open Points"):
            open_points_found = True
            print("Open Points section found:")
            continue
        if open_points_found and paragraph.text.strip().startswith("Appendix"):
            print("End of Open Points section found.")
            break
        if open_points_found:
            paragraph_text = paragraph.text.strip()
            hyperlink_present = has_hyperlink(paragraph)
            print(f"Text: {paragraph_text}, Hyperlink present: {hyperlink_present}")
            if "http" in paragraph_text:
                print(f"Plain text URL found: {paragraph_text}")

# Path to the directory containing the Word documents
doc_directory = r"C:\Users\lazare.kolebka\OneDrive - Accenture\Documents\Dev\Yara"

# Iterate over all Word documents in the directory
for filename in os.listdir(doc_directory):
    if filename.endswith(".docx"):
        doc_path = os.path.join(doc_directory, filename)
        print_open_points_links(doc_path)
        print(f"Finished checking Open Points links in document: {doc_path}")

print("Open Points links check completed.")
