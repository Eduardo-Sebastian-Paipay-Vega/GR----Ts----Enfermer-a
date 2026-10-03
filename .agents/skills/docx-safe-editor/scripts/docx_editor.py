"""
Docx Safe Editor Utility Script
Safely replaces text in Word (.docx) documents preserving formatting at the run level.
"""

import sys
import os
import docx

def safe_replace_in_docx(file_path: str, old_text: str, new_text: str, output_path: str = None) -> str:
    """
    Safely replaces occurrences of old_text with new_text in paragraphs,
    tables, headers, and footers without wiping formatting.
    """
    if output_path is None:
        base, ext = os.path.splitext(file_path)
        output_path = f"{base}_modified{ext}"

    doc = docx.Document(file_path)

    def process_paragraph(p):
        if old_text not in p.text:
            return

        # Attempt 1: Fast direct run-level replacement
        replaced = False
        for run in p.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
                replaced = True

        # Attempt 2: If old_text spans across multiple runs (split by Word)
        if not replaced and old_text in p.text:
            full_text = p.text
            new_full_text = full_text.replace(old_text, new_text)
            if p.runs:
                first_run = p.runs[0]
                first_run.text = new_full_text
                for r in p.runs[1:]:
                    r.text = ""

    def process_paragraphs(paragraphs):
        for p in paragraphs:
            process_paragraph(p)

    # 1. Main body paragraphs
    process_paragraphs(doc.paragraphs)

    # 2. Tables (cells paragraphs and nested tables)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_paragraphs(cell.paragraphs)

    # 3. Headers and Footers
    for section in doc.sections:
        process_paragraphs(section.header.paragraphs)
        for t in section.header.tables:
            for row in t.rows:
                for cell in row.cells:
                    process_paragraphs(cell.paragraphs)

        process_paragraphs(section.footer.paragraphs)
        for t in section.footer.tables:
            for row in t.rows:
                for cell in row.cells:
                    process_paragraphs(cell.paragraphs)

    doc.save(output_path)
    return output_path

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: py -3.13 docx_editor.py <file.docx> <old_text> <new_text> [output.docx]")
        sys.exit(1)

    in_file = sys.argv[1]
    o_text = sys.argv[2]
    n_text = sys.argv[3]
    out_file = sys.argv[4] if len(sys.argv) > 4 else in_file

    saved = safe_replace_in_docx(in_file, o_text, n_text, out_file)
    print(f"Successfully updated document saved to: {saved}")
