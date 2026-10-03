---
name: docx-safe-editor
description: >-
  Safely edit Microsoft Word (.docx) files without losing formats or styles.
  Use when asked to modify, fill templates, replace text, or edit a Word document.
---

# Docx Safe Editor (Antigravity)

A skill to safely edit and replace text in Microsoft Word (`.docx`) documents without breaking existing formatting (bold, colors, fonts, margins, line spacing, and table layouts).

---

## When to Use
* Modifying, filling templates, or replacing text in any `.docx` file using `python-docx`.
* Editing research protocols, thesis templates, or institutional forms (such as `plantilla/PROTOCOLO PROYECTOS CUALITATIVOS FINAL.docx`) preserving university formatting.

---

## Core Principles & Gotchas

1. **NEVER use `paragraph.text = new_text`:**  
   This completely wipes out all inline formatting (`runs`), hyperlink references, fonts, colors, and resets the paragraph to default plain text.
2. **ALWAYS operate at the Run level:**  
   Word paragraphs are split into `runs` (formatting boundaries). Replacements must be applied inside `run.text`.
3. **Handling Split Runs:**  
   If Word splits a placeholder (like `{{NAME}}`) or word across multiple runs, direct replacement on a single run may fail. To safely resolve this:
   * First, attempt run-level replacement if the target string exists entirely within a single run.
   * If the string spans across multiple runs in the same paragraph, preserve the first run's formatting, assign the updated text to it, and empty subsequent runs in that phrase.
4. **Tables and Headers/Footers:**  
   Many Word templates store form fields inside table cells, headers, and footers. The replacement logic must traverse `doc.tables`, `doc.sections[].header`, and `doc.sections[].footer`.

---

## Safe Replacement Implementation

When executing Python to modify a Word document, use this robust approach or invoke [scripts/docx_editor.py](./scripts/docx_editor.py):

```python
import docx

def safe_replace_in_docx(file_path: str, old_text: str, new_text: str, output_path: str):
    doc = docx.Document(file_path)
    
    def process_paragraph(p):
        if old_text not in p.text:
            return
            
        # Fast direct run-level replacement
        replaced = False
        for run in p.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
                replaced = True
                
        # Fallback: if old_text is split across consecutive runs
        if not replaced and old_text in p.text:
            full_text = p.text
            new_full_text = full_text.replace(old_text, new_text)
            if p.runs:
                p.runs[0].text = new_full_text
                for r in p.runs[1:]:
                    r.text = ""
                    
    def process_paragraphs(paragraphs):
        for p in paragraphs:
            process_paragraph(p)
            
    # 1. Main document body
    process_paragraphs(doc.paragraphs)
    
    # 2. Tables (cells, nested paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_paragraphs(cell.paragraphs)
                
    # 3. Headers and Footers
    for section in doc.sections:
        process_paragraphs(section.header.paragraphs)
        process_paragraphs(section.footer.paragraphs)
        
    doc.save(output_path)
```

---

## Command Line Helper

You can also run the built-in script directly from PowerShell / Terminal:

```powershell
py -3.13 "c:\GRESLY\.agents\skills\docx-safe-editor\scripts\docx_editor.py" "ruta\documento.docx" "TEXTO_VIEJO" "TEXTO_NUEVO" "ruta\documento_modificado.docx"
```

---

## Steps for the Agent

1. **Inspect the user's `.docx` file** and identify placeholders or exact phrases to modify.
2. **Execute Python with `python-docx`** (using Python 3.13: `py -3.13`).
3. **Implement run-level replacement** to preserve font size, family, color, bold, and margins.
4. **Save the modified `.docx`** and report the saved location to the user with a clickable link.
