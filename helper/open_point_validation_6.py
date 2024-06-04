from docx import Document

def check_open_points(doc):
    open_points_found = False
    link_found = False

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
            continue
        if open_points_found and paragraph.text.strip().startswith("Appendix"):
            break
        if open_points_found:
            if "http" in paragraph.text or has_hyperlink(paragraph):
                link_found = True
                break

    if link_found:
        print("Check 6: ✅ Link found in Open Points section.")
        return True
    else:
        print("Check 6: ❌ No link found in Open Points section.")
        return False
