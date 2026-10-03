import win32com.client as win32
import os
import fitz

word = win32.Dispatch('Word.Application')
word.Visible = False
try:
    doc_path = os.path.abspath(r'c:\GRESLY\DOCUMENTOS\Generación\protocolo_cualitativo.docx')
    pdf_path = os.path.abspath(r'c:\GRESLY\DOCUMENTOS\Generación\build\word_rendered.pdf')
    doc = word.Documents.Open(doc_path)
    doc.SaveAs(pdf_path, FileFormat=17)
    doc.Close(False)
finally:
    word.Quit()

pdf = fitz.open(pdf_path)
print('Exported PDF pages:', len(pdf))
for i in range(len(pdf)):
    hw = ' '.join(w[4] for w in pdf[i].get_text('words') if w[1] < 60)
    fw = ' '.join(w[4] for w in pdf[i].get_text('words') if w[1] > 780)
    print(f'Page {i+1}: Header: "{hw}" | Footer: "{fw}"')
