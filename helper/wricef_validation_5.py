import os
from docx import Document
from collections import defaultdict

def read_wricef_tables(doc):
    wricef_tables = []

    for table in doc.tables:
        if table.rows and len(table.rows) > 1 and table.rows[1].cells:
            # Check second row for WRICEF ID header
            second_row_cells = [cell.text.strip() for cell in table.rows[1].cells]
            if second_row_cells[0].lower() == "wricef id":
                wricef_tables.append(table)

    return wricef_tables

def validate_wricef_table(wricef_table):
    correct_count = 0
    value_counts = defaultdict(int)
    # Start from the third row to skip the header and ensure it is valid
    for row in wricef_table.rows[2:]:
        wricef_id = row.cells[0].text.strip().lower()
        try:
            # Check if WRICEF ID is a valid integer
            int(wricef_id)
            correct_count += 1
        except ValueError:
            value_counts[wricef_id] += 1
    return correct_count, value_counts

def print_table(table):
    table_content = []
    for row in table.rows:
        row_content = [cell.text.strip() for cell in row.cells]
        table_content.append(" | ".join(row_content))
    return "\n".join(table_content)
