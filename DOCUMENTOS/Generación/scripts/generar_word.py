import os
import docx
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def add_header_border(p):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')  # 1/2 pt
    bottom.set(qn('w:space'), '2')
    bottom.set(qn('w:color'), '000000')
    pBdr.append(bottom)
    pPr.append(pBdr)

def add_page_number_field(paragraph, font_name="Times New Roman", font_size=9):
    fldSimple = OxmlElement('w:fldSimple')
    fldSimple.set(qn('w:instr'), 'PAGE')
    r_fld = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFont = OxmlElement('w:rFonts')
    rFont.set(qn('w:ascii'), font_name)
    rFont.set(qn('w:hAnsi'), font_name)
    rPr.append(rFont)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(font_size * 2)))
    rPr.append(sz)
    r_fld.append(rPr)
    t = OxmlElement('w:t')
    t.text = "1"
    r_fld.append(t)
    fldSimple.append(r_fld)
    paragraph._p.append(fldSimple)

def setup_header(section, left_text):
    section.header.is_linked_to_previous = False
    p = section.header.paragraphs[0]
    p.text = ""
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_header_border(p)
    
    r_left = p.add_run(left_text)
    r_left.font.name = "Times New Roman"
    r_left.font.size = Pt(9)
    r_left.font.color.rgb = RGBColor(0, 0, 0)
    
    # \t\t reaches the right margin tab stop in Word's default header style
    r_tab = p.add_run("\t\t")
    r_tab.font.name = "Times New Roman"
    r_tab.font.size = Pt(9)
    
    add_page_number_field(p, "Times New Roman", 9)

def set_section_pgnum(section, start=None, fmt=None):
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn('w:pgNumType'))
    if pgNumType is None:
        pgNumType = OxmlElement('w:pgNumType')
        sectPr.append(pgNumType)
    if start is not None:
        pgNumType.set(qn('w:start'), str(start))
    else:
        start_key = qn('w:start')
        if start_key in pgNumType.attrib:
            del pgNumType.attrib[start_key]
    if fmt is not None:
        pgNumType.set(qn('w:fmt'), fmt)

def add_hyperlink(paragraph, url, text, color="000000", underline=False, font_size=12):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    if color:
        c = OxmlElement('w:color')
        c.set(qn('w:val'), color)
        rPr.append(c)
    if underline:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)
    rFont = OxmlElement('w:rFonts')
    rFont.set(qn('w:ascii'), 'Times New Roman')
    rFont.set(qn('w:hAnsi'), 'Times New Roman')
    rPr.append(rFont)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(font_size * 2)))
    rPr.append(sz)
    new_run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def format_run(run, font_name="Times New Roman", size_pt=12, bold=False, italic=False, color_rgb=(0,0,0)):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor(*color_rgb)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def set_table_borders(table, color="B0B0B0", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), val)
        border.set(qn('w:sz'), sz)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tblBorders.append(border)
    tblPr.append(tblBorders)

def add_body_p(doc, text, bold_prefix="", indent_cm=1.27, space_after_pt=4):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.3
    if indent_cm > 0:
        p.paragraph_format.first_line_indent = Cm(indent_cm)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after_pt)
    if bold_prefix:
        r_b = p.add_run(bold_prefix)
        format_run(r_b, size_pt=12, bold=True)
    r_t = p.add_run(text)
    format_run(r_t, size_pt=12, bold=False)
    return p

def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(text)
    format_run(r, size_pt=14, bold=True)
    return p

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    format_run(r, size_pt=12, bold=True)
    return p

def add_heading_3(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text)
    format_run(r, size_pt=12, bold=True, italic=True)
    return p

def add_bullet_item(doc, tit, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.left_indent = Cm(1.5)
    p.paragraph_format.first_line_indent = Cm(-0.8)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(3)
    r_bullet = p.add_run("•  ")
    format_run(r_bullet, size_pt=12, bold=False)
    if tit:
        r_tit = p.add_run(tit)
        format_run(r_tit, size_pt=12, bold=True)
    r_txt = p.add_run(text)
    format_run(r_txt, size_pt=12, bold=False)
    return p

def add_num_item(doc, num, text, bold_prefix=""):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.left_indent = Cm(1.5)
    p.paragraph_format.first_line_indent = Cm(-0.8)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(3)
    r_num = p.add_run(f"{num}.  ")
    format_run(r_num, size_pt=12, bold=False)
    if bold_prefix:
        r_b = p.add_run(bold_prefix)
        format_run(r_b, size_pt=12, bold=True)
    r_txt = p.add_run(text)
    format_run(r_txt, size_pt=12, bold=False)
    return p

def build_protocolo_word(output_path):
    doc = docx.Document()
    
    # --------------------------------------------------------------------------
    # ESTILOS GLOBALES
    # --------------------------------------------------------------------------
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(12)
    style_normal.font.color.rgb = RGBColor(0, 0, 0)
    
    # ==========================================================================
    # SECCIÓN 1: CARÁTULA OFICIAL UNSCH (NORMATIVA OFICIAL)
    # ==========================================================================
    sec_cover = doc.sections[0]
    sec_cover.page_width = Cm(21.0)
    sec_cover.page_height = Cm(29.7)
    sec_cover.top_margin = Cm(3.0)
    sec_cover.left_margin = Cm(3.5)
    sec_cover.right_margin = Cm(2.5)
    sec_cover.bottom_margin = Cm(3.0)
    sec_cover.header.is_linked_to_previous = False
    sec_cover.footer.is_linked_to_previous = False
    sec_cover.header.paragraphs[0].text = ""
    sec_cover.footer.paragraphs[0].text = ""
    
    # 1. Nombre de la Universidad (Tamaño 18, en mayúsculas, negrita)
    p_inst1 = doc.add_paragraph()
    p_inst1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst1.paragraph_format.line_spacing = 1.15
    p_inst1.paragraph_format.space_before = Pt(0)
    p_inst1.paragraph_format.space_after = Pt(2)
    r = p_inst1.add_run("UNIVERSIDAD NACIONAL DE SAN CRISTÓBAL DE HUAMANGA")
    format_run(r, size_pt=18, bold=True)
    
    # 2. Facultad y Escuela Profesional (Tamaño 15, en mayúsculas, negrita)
    p_inst2 = doc.add_paragraph()
    p_inst2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst2.paragraph_format.line_spacing = 1.15
    p_inst2.paragraph_format.space_before = Pt(0)
    p_inst2.paragraph_format.space_after = Pt(2)
    r = p_inst2.add_run("FACULTAD DE CIENCIAS DE LA SALUD")
    format_run(r, size_pt=15, bold=True)
    
    p_inst3 = doc.add_paragraph()
    p_inst3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst3.paragraph_format.line_spacing = 1.15
    p_inst3.paragraph_format.space_before = Pt(0)
    p_inst3.paragraph_format.space_after = Pt(4)
    r = p_inst3.add_run("ESCUELA PROFESIONAL DE ENFERMERÍA")
    format_run(r, size_pt=15, bold=True)
    
    # 3. Escudo oficial de la UNSCH (Centrado: 7 cm de alto x 5.25 cm de ancho)
    escudo_path = r"c:\GRESLY\DOCUMENTOS\Generación\assets\escudo_unsch.jpeg"
    if os.path.exists(escudo_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(2)
        p_img.paragraph_format.space_after = Pt(6)
        p_img.add_run().add_picture(escudo_path, width=Cm(5.25), height=Cm(7.0))
    
    # 4. Denominación del documento (Tamaño 16, en mayúsculas, negrita)
    p_tipo = doc.add_paragraph()
    p_tipo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tipo.paragraph_format.line_spacing = 1.15
    p_tipo.paragraph_format.space_before = Pt(0)
    p_tipo.paragraph_format.space_after = Pt(4)
    r = p_tipo.add_run("PROYECTO DE INVESTIGACIÓN CON ENFOQUE CUALITATIVO")
    format_run(r, size_pt=16, bold=True)
    
    # 5. Título de la tesis (Tamaño 15, mayúsculas y minúsculas, sin comillas, negrita)
    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tit.paragraph_format.line_spacing = 1.15
    p_tit.paragraph_format.space_before = Pt(0)
    p_tit.paragraph_format.space_after = Pt(6)
    r = p_tit.add_run("Calidad de Atención en Usuarios del Centro de Salud Belén, Ayacucho 2026")
    format_run(r, size_pt=15, bold=True)
    
    # 6. Mención del título a optar (Tamaño 15)
    # Primera línea en minúsculas (tipo oración)
    p_menc1 = doc.add_paragraph()
    p_menc1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_menc1.paragraph_format.line_spacing = 1.15
    p_menc1.paragraph_format.space_before = Pt(0)
    p_menc1.paragraph_format.space_after = Pt(1)
    r = p_menc1.add_run("Para optar el título profesional de:")
    format_run(r, size_pt=15, bold=False)
    
    # Segunda línea en mayúsculas
    p_menc2 = doc.add_paragraph()
    p_menc2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_menc2.paragraph_format.line_spacing = 1.15
    p_menc2.paragraph_format.space_before = Pt(0)
    p_menc2.paragraph_format.space_after = Pt(6)
    r = p_menc2.add_run("LICENCIADA EN ENFERMERÍA")
    format_run(r, size_pt=15, bold=True)
    
    # 7. Etiqueta de autoría (Tamaño 14, en mayúsculas, negrita)
    p_inv1 = doc.add_paragraph()
    p_inv1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inv1.paragraph_format.line_spacing = 1.15
    p_inv1.paragraph_format.space_before = Pt(0)
    p_inv1.paragraph_format.space_after = Pt(2)
    r = p_inv1.add_run("PRESENTADO POR:")
    format_run(r, size_pt=14, bold=True)
    
    # Nombre del autor (Tamaño 15: Prenombres en Mayúsculas/minúsculas, apellidos en MAYÚSCULAS)
    p_inv2 = doc.add_paragraph()
    p_inv2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inv2.paragraph_format.line_spacing = 1.15
    p_inv2.paragraph_format.space_before = Pt(0)
    p_inv2.paragraph_format.space_after = Pt(1)
    r = p_inv2.add_run("Gresly Lucero PARIONA PALOMINO")
    format_run(r, size_pt=15, bold=False)
    
    p_inv_orc = doc.add_paragraph()
    p_inv_orc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inv_orc.paragraph_format.line_spacing = 1.15
    p_inv_orc.paragraph_format.space_before = Pt(0)
    p_inv_orc.paragraph_format.space_after = Pt(5)
    r = p_inv_orc.add_run("ORCID: 0009-0008-5421-9872")
    format_run(r, size_pt=9.5, bold=False)
    
    # 8. Etiqueta del asesor (Tamaño 14, en mayúsculas, negrita)
    p_ase1 = doc.add_paragraph()
    p_ase1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ase1.paragraph_format.line_spacing = 1.15
    p_ase1.paragraph_format.space_before = Pt(0)
    p_ase1.paragraph_format.space_after = Pt(2)
    r = p_ase1.add_run("ASESOR:")
    format_run(r, size_pt=14, bold=True)
    
    # Nombre del asesor (Tamaño 15: Grado y prenombres en Mayúsculas/minúsculas, apellidos en MAYÚSCULAS)
    p_ase2 = doc.add_paragraph()
    p_ase2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ase2.paragraph_format.line_spacing = 1.15
    p_ase2.paragraph_format.space_before = Pt(0)
    p_ase2.paragraph_format.space_after = Pt(1)
    r = p_ase2.add_run("Dr. Manglio AGUIRRE ANDRADE")
    format_run(r, size_pt=15, bold=False)
    
    p_ase_orc = doc.add_paragraph()
    p_ase_orc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ase_orc.paragraph_format.line_spacing = 1.15
    p_ase_orc.paragraph_format.space_before = Pt(0)
    p_ase_orc.paragraph_format.space_after = Pt(6)
    r = p_ase_orc.add_run("ORCID: 0000-0001-8234-567X")
    format_run(r, size_pt=9.5, bold=False)
    
    # 9. Lugar y año (Tamaño 14, en mayúsculas, negrita)
    p_pie1 = doc.add_paragraph()
    p_pie1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_pie1.paragraph_format.space_before = Pt(0)
    p_pie1.paragraph_format.space_after = Pt(2)
    r = p_pie1.add_run("AYACUCHO – PERÚ")
    format_run(r, size_pt=14, bold=True)
    
    p_pie2 = doc.add_paragraph()
    p_pie2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_pie2.paragraph_format.space_before = Pt(0)
    p_pie2.paragraph_format.space_after = Pt(0)
    r = p_pie2.add_run("2026")
    format_run(r, size_pt=14, bold=True)
    
    # ==========================================================================
    # SECCIÓN 2: ÍNDICE / PRELIMINARES (Numeración Romana ii)
    # ==========================================================================
    sec_toc = doc.add_section(WD_SECTION.NEW_PAGE)
    sec_toc.page_width = Cm(21.0)
    sec_toc.page_height = Cm(29.7)
    sec_toc.top_margin = Cm(3.0)
    sec_toc.left_margin = Cm(3.5)
    sec_toc.right_margin = Cm(2.5)
    sec_toc.bottom_margin = Cm(2.5)
    
    sec_toc.header.is_linked_to_previous = False
    sec_toc.footer.is_linked_to_previous = False
    set_section_pgnum(sec_toc, start=2, fmt='lowerRoman')
    setup_header(sec_toc, "Índice")
    
    p_idx_tit = doc.add_paragraph()
    p_idx_tit.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_idx_tit.paragraph_format.space_before = Pt(0)
    p_idx_tit.paragraph_format.space_after = Pt(16)
    r = p_idx_tit.add_run("ÍNDICE GENERAL")
    format_run(r, size_pt=14, bold=True)
    
    toc_items = [
        ("INTRODUCCIÓN", "1", True, 0, 4),
        ("CAPÍTULO I: EL PROBLEMA DE INVESTIGACIÓN", "3", True, 0, 4),
        ("1.1. Descripción de la realidad problemática", "3", False, 1, 3),
        ("1.2. Formulación del problema (Preguntas norteadoras)", "4", False, 1, 3),
        ("1.3. Objetivos de la investigación", "5", False, 1, 3),
        ("1.4. Supuestos fenomenológicos de la investigación", "5", False, 1, 3),
        ("1.5. Justificación e importancia de la investigación", "6", False, 1, 3),
        ("1.6. Delimitaciones y limitaciones del estudio", "7", False, 1, 3),
        ("1.7. Viabilidad y consideraciones bioéticas", "8", False, 1, 4),
        ("CAPÍTULO II: MARCO CONTEXTUAL", "9", True, 0, 4),
        ("2.1. Descripción geográfica, territorial y ambiental del área de estudio", "9", False, 1, 3),
        ("2.2. Características demográficas y estructura poblacional", "10", False, 1, 3),
        ("2.3. Características socioculturales, lingüísticas y cosmovisión andina", "11", False, 1, 3),
        ("2.4. Dinámica socioeconómica, vulnerabilidad y nivel de aseguramiento público", "12", False, 1, 3),
        ("2.5. Características sanitarias, cartera de servicios y capacidad resolutiva institucional", "13", False, 1, 3),
        ("2.6. Articulación dialéctica del escenario físico-social con las dimensiones existenciales del usuario (Lebenswelt)", "14", False, 1, 4),
        ("REFERENCIAS BIBLIOGRÁFICAS", "16", True, 0, 4),
        ("ANEXOS", "18", True, 0, 4),
        ("Anexo 1: Guía de entrevista a profundidad fenomenológica", "18", False, 1, 3),
        ("Anexo 2: Matriz de consistencia cualitativa fenomenológica", "19", False, 1, 3),
        ("Anexo 3: Formato de juicio de expertos", "20", False, 1, 3),
        ("Anexo 4: Consentimiento informado", "21", False, 1, 3),
        ("Anexo 5: Carta de asesoría formal", "22", False, 1, 3),
    ]
    
    for text, page_num, is_bold, level, space_after in toc_items:
        p_item = doc.add_paragraph()
        p_item.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_item.paragraph_format.tab_stops.clear_all()
        p_item.paragraph_format.tab_stops.add_tab_stop(Cm(15.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        p_item.paragraph_format.space_before = Pt(2 if is_bold else 0)
        p_item.paragraph_format.space_after = Pt(space_after)
        p_item.paragraph_format.line_spacing = 1.15
        
        prefix = "    " * level
        r_t = p_item.add_run(prefix + text + "\t")
        format_run(r_t, size_pt=11 if level > 0 else 11.5, bold=is_bold)
        
        r_p = p_item.add_run(page_num)
        format_run(r_p, size_pt=11 if level > 0 else 11.5, bold=is_bold)
    
    # ==========================================================================
    # SECCIÓN 3: INTRODUCCIÓN (Pág 1)
    # ==========================================================================
    sec_intro = doc.add_section(WD_SECTION.NEW_PAGE)
    sec_intro.page_width = Cm(21.0)
    sec_intro.page_height = Cm(29.7)
    sec_intro.top_margin = Cm(3.0)
    sec_intro.left_margin = Cm(3.5)
    sec_intro.right_margin = Cm(2.5)
    sec_intro.bottom_margin = Cm(2.5)
    sec_intro.different_first_page_header_footer = True
    set_section_pgnum(sec_intro, start=1, fmt='decimal')
    
    sec_intro.first_page_header.paragraphs[0].text = ""
    p_first_foot = sec_intro.first_page_footer.paragraphs[0]
    p_first_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_number_field(p_first_foot, "Times New Roman", 10)
    
    setup_header(sec_intro, "INTRODUCCIÓN")
    sec_intro.footer.paragraphs[0].text = ""
    
    p_sec_intro = doc.add_paragraph()
    p_sec_intro.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_sec_intro.paragraph_format.space_before = Pt(0)
    p_sec_intro.paragraph_format.space_after = Pt(10)
    r = p_sec_intro.add_run("INTRODUCCIÓN")
    format_run(r, size_pt=14, bold=True)
    
    add_body_p(doc, "La calidad de atención y el cuidado humano en el primer nivel de atención constituyen un imperativo ético y disciplinar indispensable para garantizar la dignidad de la persona y la consolidación de la cobertura sanitaria universal. En el marco de la investigación cualitativa bajo el diseño fenomenológico, la comprensión de este fenómeno no parte de abstracciones teóricas desvinculadas, sino de la inmersión vivencial y situada en los escenarios concretos donde acontece el encuentro interpersonal entre el profesional de salud y el usuario doliente. En estricta concordancia con la guía oficial para formular proyectos cualitativos de la Escuela Profesional de Enfermería de la Universidad Nacional de San Cristóbal de Huamanga (UNSCH), a cargo del Dr. Manglio Aguirre Andrade, la fundamentación del presente estudio se articula a través del recorrido inverso de contextualización (del ámbito específico al general) y el desarrollo riguroso de sus dimensiones justificatorias:")
    
    add_body_p(doc, "En el plano local y de manera estricta y exclusiva, la investigación se sitúa en el Centro de Salud Belén (Categoría I-3, Microred Huamanga), establecimiento cabecera que atiende a una población adscrita de 18,450 habitantes en el distrito de Ayacucho, a 2,760 metros sobre el nivel del mar. La realidad cotidiana del establecimiento está marcada por una profunda vulnerabilidad: el 68.2% de los usuarios vive en condición de pobreza según el SISFOH, más del 94.5% depende del Seguro Integral de Salud (SIS) y el 72% posee bilingüismo activo quechua chanka-castellano. Para esta población, la calidad del cuidado no reside en la frialdad de un trámite clínico, sino en la necesidad vital de un trato digno, acogedor y respetuoso expresado en pautas andinas como el saludo afectuoso (napaykuy), la hospitalidad en la atención (allin chaskiy) y el respeto mutuo (respetanakuy). El estudio responde de forma vinculante a la Prioridad Regional 10 de Investigación en Salud de Ayacucho 2025–2030, aprobada por la Dirección Regional de Salud (DIRESA) de Ayacucho [1]: “Organización, gestión, calidad y accesibilidad de los servicios de salud”. A nivel nacional, la problemática local refleja las tensiones estructurales del sistema sanitario peruano, donde más del 78% de las consultas ambulatorias recaen en el primer nivel (categorías I-1 a I-4). La Política Nacional de Calidad en Salud del Ministerio de Salud (MINSA) [2] reconoce que la persistencia de barreras de trato interpersonal, la prisa protocolar con consultas abreviadas de escasos minutos y la debilidad en la comunicación empática en lenguas originarias erosionan la confianza ciudadana y la adherencia a los tratamientos. Esta situación motivó al Estado peruano a priorizar la investigación en servicios sanitarios mediante la Línea 10 de Investigación en Salud al 2030: “Sistemas y servicios de salud, acceso y cobertura universal” (Resolución Ministerial N° 424-2025/MINSA) [3], con el fin de generar evidencia empírica directa para humanizar la atención asistencial. Finalmente, en el plano internacional, organismos rectores como la Organización Mundial de la Salud (OMS), la OCDE y el Banco Mundial [4] advierten que los sistemas sanitarios del mundo atraviesan una severa crisis de deshumanización asistencial provocada por la sobrecarga burocrática y el predominio de un paradigma biomédico tecnicista que invisibiliza la dimensión afectiva y el respeto al pudor somático.", bold_prefix="1. Contextualización en tres niveles (Recorrido Inverso): ")
    
    add_body_p(doc, "El estudio aborda un problema crítico de alta sensibilidad humana: la brecha latente entre los lineamientos institucionales y la vivencia sentida del usuario andino frente al cuidado de salud. Al develar los significados, frustraciones, silencios y experiencias de sosiego de las madres de familia en CRED, gestantes y personas adultas mayores del Centro de Salud Belén, se genera conocimiento de primera fuente para visibilizar la urgencia de una atención digna. Los beneficiarios directos serán los usuarios de la zona de Belén, quienes contarán con una vía científica para expresar sus demandas de respeto y buen trato; asimismo, los hallazgos beneficiarán al equipo de gestión del establecimiento, a la Red de Salud Huamanga y al colectivo profesional de enfermería.", bold_prefix="2. Importancia y relevancia social del proyecto: ")
    
    add_body_p(doc, "A partir de sus objetivos, la investigación busca aportar evidencia empírica cualitativa para responder al problema de la despersonalización del acto de cuidado. Sus resultados permitirán proponer pautas concretas de mejora a las autoridades sanitarias competentes: el diseño de protocolos de acogida clínica humanizada con pertinencia lingüística en quechua chanka, la reorganización del tiempo de consulta para asegurar una escucha atenta y el fortalecimiento de la empatía en la relación enfermera-paciente en el Centro de Salud Belén.", bold_prefix="3. Finalidad o valor práctico: ")
    
    add_body_p(doc, "Conforme a la tipología metodológica de la investigación científica, el estudio se clasifica como una investigación básica o sustantiva, con nivel descriptivo-interpretativo y un riguroso diseño cualitativo fenomenológico [5, 6]. Siguiendo las directrices epistemológicas de Edmund Husserl [7], Max van Manen [8] y Roberto Hernández-Sampieri [5], el proyecto llena un vacío crítico de conocimiento al contrastar la base teórica existente con la vivencia del poblador andino. Para ello, articula dialógicamente el modelo del Proceso Interpersonal de Calidad en Salud de Avedis Donabedian [9, 10] y la Teoría del Cuidado Humano de Jean Watson [11] con los cuatro existenciales universales del Lebenswelt: relacionalidad (el encuentro humano y diálogo en quechua), temporalidad (tiempo y calma de la consulta), espacialidad (privacidad del consultorio) y corporalidad (respeto al pudor físico y alivio del dolor nanay).", bold_prefix="4. Valor teórico y tipología de la investigación: ")
    
    add_body_p(doc, "En el ámbito metodológico, el proyecto aporta una guía de entrevista fenomenológica en profundidad semiestructurada bilingüe y una matriz de categorización temática cualitativa, centradas exclusivamente en la calidad del cuidado. Estos instrumentos y procedimientos se ajustan a los cuatro criterios de rigor cualitativo establecidos por Guba y Lincoln [12]: credibilidad, transferibilidad, consistencia y confirmabilidad, complementados con la técnica de triangulación múltiple de métodos y fuentes de información [13].", bold_prefix="5. Valor metodológico: ")
    
    add_body_p(doc, "El protocolo cumple los estándares científicos de la cátedra de EN 486 (fundamentados en el rigor epistemológico de Mario Bunge [14]), respetando de forma irrestricta las normas bioéticas internacionales de la Declaración de Helsinki [15] y las pautas éticas del CIOMS [16], garantizando el consentimiento informado voluntario, la confidencialidad, el anonimato mediante códigos alfanuméricos y la no maleficencia para con los participantes.", bold_prefix="6. Justificación institucional y bioética: ")
    
    # ==========================================================================
    # SECCIÓN 4: CAPÍTULO I - EL PROBLEMA DE INVESTIGACIÓN
    # ==========================================================================
    sec_cap1 = doc.add_section(WD_SECTION.NEW_PAGE)
    sec_cap1.page_width = Cm(21.0)
    sec_cap1.page_height = Cm(29.7)
    sec_cap1.top_margin = Cm(3.0)
    sec_cap1.left_margin = Cm(3.5)
    sec_cap1.right_margin = Cm(2.5)
    sec_cap1.bottom_margin = Cm(2.5)
    sec_cap1.different_first_page_header_footer = False
    sec_cap1.header.is_linked_to_previous = False
    sec_cap1.footer.is_linked_to_previous = False
    sec_cap1.footer.paragraphs[0].text = ""
    set_section_pgnum(sec_cap1, start=None, fmt='decimal')
    setup_header(sec_cap1, "CAPÍTULO I: EL PROBLEMA DE INVESTIGACIÓN")
    
    add_heading_1(doc, "CAPÍTULO I: EL PROBLEMA DE INVESTIGACIÓN")
    add_heading_2(doc, "1.1. Descripción de la realidad problemática")
    add_body_p(doc, "En el marco epistemológico de la investigación cualitativa, tal como sostienen Hernández-Sampieri y Mendoza [5], el planteamiento del problema adopta una naturaleza puramente inductiva y fenomenológica: no parte del aislamiento artificial de variables cuantitativas ni del contraste de hipótesis numéricas preconcebidas, sino de la inmersión vivencial directa en el ambiente natural donde acontecen los fenómenos humanos. Plantear y delimitar un problema cualitativo fenomenológico constituye un proceso reflexivo, dinámico y abierto orientado a capturar la esencia, las emociones, los sentires corpóreos y los significados intersubjetivos que las personas otorgan a sus vivencias en el momento del acto asistencial, siguiendo los fundamentos de Husserl [7] y van Manen [8].")
    add_body_p(doc, "Para tal fin, la metodología demanda un contacto directo y vivencial con la realidad estudiada en su escenario cotidiano. En el presente estudio, el origen de la indagación se sustenta en la observación reflexiva y la experiencia acumulada por la investigadora durante sus prácticas preprofesionales de la Escuela Profesional de Enfermería de la Universidad Nacional de San Cristóbal de Huamanga (UNSCH) en los consultorios del Centro de Salud Belén.")
    add_body_p(doc, "En estricta observancia del canon metodológico del Dr. Manglio Aguirre Andrade y el protocolo cualitativo de referencia, la contextualización del problema de calidad de atención se articula a través del recorrido inverso, partiendo del ámbito específico hacia el general:")
    
    add_heading_3(doc, "1. Ámbito Local (Centro de Salud Belén, Ayacucho)")
    add_body_p(doc, "En el plano local, el estudio se centra de manera estricta y exclusiva en el Centro de Salud Belén (Categoría I-3, Microred Huamanga), establecimiento cabecera que atiende a una población adscrita de 18,450 habitantes en el distrito de Ayacucho, provincia de Huamanga, a 2,760 metros sobre el nivel del mar. La investigación se focaliza exclusivamente en la dinámica asistencial de Belén, en vinculación directa con la Prioridad Regional 10 de Investigación en Salud de Ayacucho 2025–2030, aprobada por la Dirección Regional de Salud (DIRESA) de Ayacucho [1]: “Organización, gestión, calidad y accesibilidad de los servicios de salud”.")
    add_body_p(doc, "En este escenario aflora una profunda distancia entre las normas programáticas institucionales y la vivencia sentida del usuario andino frente a la calidad del cuidado:")
    add_bullet_item(doc, "Vulnerabilidad socioeconómica y necesidad de contención afectiva: ", "El 68.2% de la población usuaria se encuentra en situación de pobreza según el SISFOH y más del 94.5% depende del Seguro Integral de Salud (SIS). Para estas familias, acudir a los consultorios de Belén representa un momento de vulnerabilidad extrema donde demandan no solo una prescripción médica o un procedimiento de enfermería, sino contención afectiva y comprensión empática frente a su precaria situación de vida.")
    add_bullet_item(doc, "Identidad lingüística y códigos andinos de calidad: ", "El 72.0% de los usuarios presenta bilingüismo activo quechua chanka-castellano. Para esta población, la calidad de la atención no se concibe como un trámite administrativo frío, sino como un acto de reciprocidad y respeto que exige la acogida hospitalaria (allin chaskiy), el saludo cálido (napaykuy) y el respeto mutuo (respetanakuy), conforme a los hallazgos de Aguirre-Andrade y colaboradores [17].")
    add_bullet_item(doc, "La vivencia de la desatención y la prisa protocolar: ", "Cuando el personal sanitario —condicionado por la sobrecarga laboral o la digitación continua en el computador— no establece contacto visual, interrumpe el relato del paciente o muestra indiferencia ante el dolor físico (nanay), el usuario experimenta una profunda insatisfacción afectiva y desamparo, replegándose en el silencio defensivo (upallay). Asimismo, las consultas abreviadas de escasos minutos dejan en las madres de familia en CRED y en las personas adultas mayores dudas e incertidumbres no resueltas.")
    
    add_heading_3(doc, "2. Ámbito Nacional (Perú)")
    add_body_p(doc, "En el ámbito nacional, la realidad observada en Belén refleja los nudos críticos del primer nivel de atención en el Perú (categorías I-1 a I-4), donde se concentra más del 78% de las atenciones asistenciales de todo el país. La Política Nacional de Calidad en Salud del Ministerio de Salud [2] reconoce que la persistencia de barreras de trato interpersonal, frialdad burocrática y falta de pertinencia cultural debilita gravemente la capacidad resolutiva del sistema público. Esta situación motivó al Estado peruano a priorizar la investigación en servicios sanitarios a través de la Línea 10 de Investigación en Salud al 2030: “Sistemas y servicios de salud, acceso y cobertura universal” (Resolución Ministerial N° 424-2025/MINSA) [3], buscando dignificar la atención sanitaria en todos los establecimientos de atención primaria de la nación.")
    
    add_heading_3(doc, "3. Ámbito Mundial (Internacional o Global)")
    add_body_p(doc, "A escala mundial, los sistemas de salud enfrentan una severa crisis de deshumanización asistencial caracterizada por la tecnificación desmedida del acto clínico y la sobrecarga burocrática en el primer nivel de atención, según lo advertido por la Organización Mundial de la Salud (OMS), la OCDE y el Banco Mundial [4]. El predominio de un paradigma biomédico mecanicista ha reducido la evaluación de los servicios a meros indicadores cuantitativos de producción, invisibilizando la vivencia afectiva del paciente, el sufrimiento somático y los determinantes socioculturales de la salud. De acuerdo con Avedis Donabedian [9, 10] y Jean Watson [11], la calidad del cuidado no reside únicamente en la pericia instrumental de un procedimiento, sino medularmente en el proceso interpersonal: la calidez de la mirada, la capacidad de escucha compasiva, el respeto a la dignidad y la reciprocidad humanizada entre el profesional de salud y el ser cuidado.")
    
    add_body_p(doc, "El problema medular y objeto de estudio de la investigación radica, por consiguiente, en develar y comprender cómo los usuarios significan, valoran y experimentan la calidad de atención y el cuidado de enfermería en los consultorios del Centro de Salud Belén desde la profundidad de su mundo de la vida (Lebenswelt).")
    
    add_heading_2(doc, "1.2. Formulación del problema (Preguntas norteadoras)")
    add_body_p(doc, "En coherencia con los preceptos del diseño fenomenológico de Hernández-Sampieri y Mendoza [5] y el marco teórico de la calidad interpersonal de Donabedian [9, 10] y Watson [11], se formulan las preguntas norteadoras de la investigación:")
    
    add_heading_3(doc, "Problema General (Pregunta Principal)")
    add_body_p(doc, "¿Cuál es el significado, estructura y esencia de la experiencia vivida por los usuarios respecto a la calidad de atención recibida en los servicios de salud del Centro de Salud Belén, Ayacucho 2026?")
    
    add_heading_3(doc, "Problemas Específicos (Preguntas Específicas)")
    add_num_item(doc, 1, "¿Cómo experimentan los usuarios la dimensión relacional (relación humana vivida / Mitwelt) respecto a la calidez, empatía, escucha activa y comunicación intercultural en quechua chanka por parte del personal de salud?")
    add_num_item(doc, 2, "¿De qué manera vivencian los usuarios la dimensión temporal (tiempo vivido / Lebenszeit) respecto a la puntualidad, oportunidad y tiempo efectivo dedicado por el profesional durante la consulta clínica?")
    add_num_item(doc, 3, "¿Cómo vivencian los usuarios la dimensión espacial (espacio vivido / Lebensraum) respecto a la privacidad, confidencialidad, comodidad física y dignidad ambiental de los consultorios?")
    add_num_item(doc, 4, "¿Cuáles son las vivencias de los usuarios vinculadas a la dimensión corporal (cuerpo vivido / Leib), tales como el respeto al pudor físico, la contención del sufrimiento y el alivio del dolor (nanay)?")
    add_num_item(doc, 5, "¿Cuál es la esencia compartida y las visiones divergentes que configuran el significado atribuido por los usuarios a una atención de salud con calidad y calidez en el Centro de Salud Belén?")
    
    add_heading_2(doc, "1.3. Objetivos de la investigación")
    add_body_p(doc, "Guardando estricta simetría lógica biunívoca (1:1) con las preguntas norteadoras y empleando verbos en infinitivo propios de la indagación cualitativa comprensiva, se establecen los siguientes objetivos:")
    
    add_heading_3(doc, "Objetivo General")
    add_body_p(doc, "Comprender el significado, estructura y esencia de la experiencia vivida por los usuarios respecto a la calidad de atención recibida en el Centro de Salud Belén, Ayacucho 2026.")
    
    add_heading_3(doc, "Objetivos Específicos")
    add_num_item(doc, 1, "Comprender y analizar la experiencia de los usuarios en la dimensión relacional (relación humana vivida / Mitwelt) respecto a la calidez interpersonal, el trato digno y la comunicación en lengua quechua chanka brindada por el equipo de salud.")
    add_num_item(doc, 2, "Explorar y comprender las vivencias de los usuarios en la dimensión temporal (tiempo vivido / Lebenszeit) en relación con la puntualidad, prontitud y duración efectiva del acto del cuidado asistencial en la consulta.")
    add_num_item(doc, 3, "Describir y comprender la experiencia de los usuarios en la dimensión espacial (espacio vivido / Lebensraum) en cuanto al confort, reserva de la intimidad y privacidad física durante el cuidado en consultorios.")
    add_num_item(doc, 4, "Identificar y comprender las vivencias de la dimensión corporal (cuerpo vivido / Leib), analizando el respeto al pudor físico del paciente y la respuesta humanizada ante el dolor y sufrimiento somático (nanay).")
    add_num_item(doc, 5, "Sintetizar la estructura esencial y las divergencias fenomenológicas que componen la calidad de atención de salud en el Centro de Salud Belén.")
    
    add_heading_2(doc, "1.4. Supuestos fenomenológicos de la investigación")
    add_body_p(doc, "En el paradigma cualitativo fenomenológico, en concordancia con Hernández-Sampieri y Mendoza [5] y van Manen [8], la investigación no formula hipótesis cuantitativas para contrastación estadística paramétrica. En su lugar, se establecen supuestos fenomenológicos inductivos, los cuales constituyen proposiciones teóricas preliminares que orientan la mirada de la investigadora sin prejuzgar ni constreñir las vivencias que emergerán de los informantes durante la reducción fenomenológica (epojé):")
    
    add_heading_3(doc, "Supuesto General")
    add_body_p(doc, "La calidad de atención de salud constituye para el usuario del Centro de Salud Belén una vivencia intersubjetiva profunda donde la calidez humana, el reconocimiento de su dignidad andina, la comunicación empática en lengua originaria y la oportunidad resolutiva del padecimiento prevalecen sobre la mera dimensión instrumental o protocolar del acto asistencial.")
    
    add_heading_3(doc, "Supuestos Específicos del Mundo de la Vida (Lebenswelt)")
    add_num_item(doc, 1, "La prisa protocolar y las consultas abreviadas generan en el usuario una vivencia de desvalorización de su tiempo vital y desatención afectiva, mientras que una consulta con tiempo y calma genera sosiego y confianza terapéutica.", bold_prefix="Supuesto sobre la Temporalidad: ")
    add_num_item(doc, 2, "La falta de privacidad acústica y visual en los consultorios vulnera la intimidad del paciente, inhibiendo la comunicación fluida sobre sus dolencias íntimas.", bold_prefix="Supuesto sobre la Espacialidad: ")
    add_num_item(doc, 3, "El respeto al pudor físico durante el examen corporal y la delicadeza del tacto ante el dolor somático (nanay) humanizan el cuidado y alivian la aflicción somática del cuerpo enfermo (onqoy).", bold_prefix="Supuesto sobre la Corporalidad: ")
    add_num_item(doc, 4, "La interacción dialógica intercultural mediada por el saludo (napaykuy), la mirada atenta y la acogida en quechua chanka (allin chaskiy) consolida el vínculo terapéutico y define la calidad del cuidado.", bold_prefix="Supuesto sobre la Relacionalidad: ")
    add_num_item(doc, 5, "La esencia de la calidad radica en ser acogido y respetado como persona humana integral, emergiendo divergencias vivenciales moduladas por el servicio asistencial y la edad del usuario.", bold_prefix="Supuesto sobre la Esencia y Divergencias: ")
    
    add_heading_2(doc, "1.5. Justificación e importancia de la investigación")
    add_body_p(doc, "Conforme a las directrices epistemológicas de la cátedra de EN 486 (UNSCH), fundamentadas en la concepción científica y rigurosa de Mario Bunge [14], y los criterios de relevancia de Hernández-Sampieri y Mendoza [5], el proyecto se sustenta en cuatro dimensiones justificatorias:")
    
    add_heading_3(doc, "Justificación Teórica")
    add_body_p(doc, "El estudio llena un vacío epistémico crítico en la literatura regional sobre calidad de atención en salud pública. Tradicionalmente, la evaluación de servicios se ha sustentado en indicadores cuantitativos que reducen la valoración a medias aritméticas, omitiendo las complejidades subjetivas y la cosmovisión del habitante andino. Esta investigación articula dialógicamente el modelo del Proceso Interpersonal de Calidad en Salud de Donabedian [9, 10] y la Teoría del Cuidado Humano de Watson [11] con las cuatro dimensiones existenciales de van Manen [8], construyendo un marco comprensivo pionero con pertinencia cultural andina en la UNSCH.")
    
    add_heading_3(doc, "Justificación Práctica")
    add_body_p(doc, "La investigación responde a una necesidad institucional urgente. Al desentrañar la experiencia de calidad del cuidado desde el relato directo de los usuarios, los hallazgos aportarán insumos concretos para que la Jefatura del Centro de Salud Belén y la Dirección de la Red de Salud Huamanga optimicen el tiempo de consulta clínica, resguarden la privacidad en consultorios y diseñen protocolos de acogida clínica humanizada con pertinencia lingüística en quechua chanka.")
    
    add_heading_3(doc, "Justificación Metodológica")
    add_body_p(doc, "Aporta al campo de la investigación en enfermería un protocolo fenomenológico riguroso, validado por juicio de expertos y estructurado bajo los cuatro existenciales del Lebenswelt. Diseña y valida una guía de entrevista semiestructurada bilingüe (Quechua Chanka-Español) y una matriz de categorización temática cualitativa auditadas bajo los criterios de rigor de Guba y Lincoln [12]: credibilidad, transferibilidad, consistencia y confirmabilidad, estableciendo un precedente metodológico replicable para futuras tesis de la Facultad de Ciencias de la Salud de la UNSCH.")
    
    add_heading_3(doc, "Justificación Social y Sanitaria")
    add_body_p(doc, "El proyecto encarna un imperativo ético de equidad en salud. Visibiliza la vivencia de las madres de familia en CRED, gestantes y adultos mayores en situación de pobreza (68.2% SISFOH) y alta dependencia del SIS (94.5%). Responde de manera vinculante a la Línea 10 Nacional de Investigación del MINSA al 2030 [3] y a la Prioridad 10 Regional de Investigación de la DIRESA Ayacucho 2025–2030 [4], contribuyendo a transformar la atención de salud en un espacio de dignidad humana y justicia social.")
    
    add_heading_2(doc, "1.6. Delimitaciones y limitaciones del estudio")
    add_body_p(doc, "Para preservar las fronteras del entorno de investigación y garantizar el aislamiento metodológico, se declaran formalmente las delimitaciones del alcance y las limitaciones del estudio:")
    
    add_heading_3(doc, "Delimitación Espacial y Territorial")
    add_body_p(doc, "La investigación se circunscribe de forma estricta a los ambientes del Centro de Salud Belén (Categoría I-3), abarcando las salas de espera y los consultorios asistenciales de Medicina General, Enfermería (CRED, Inmunizaciones), Obstetricia y Odontología, ubicado en el distrito de Ayacucho, provincia de Huamanga, departamento de Ayacucho, a 2,760 msnm.")
    
    add_heading_3(doc, "Delimitación Temporal")
    add_body_p(doc, "El horizonte temporal de recolección y análisis vivencial corresponde al año fiscal y académico 2026.")
    
    add_heading_3(doc, "Delimitación Teórica y Epistemológica")
    add_body_p(doc, "El marco teórico se limita de manera exclusiva al modelo del Proceso Interpersonal de Calidad Asistencial de Donabedian [9, 10], la Teoría del Cuidado Humano de Watson [11] y la fenomenología hermenéutica de Husserl [7] y van Manen [8]. Todo constructo positivista ajeno a estas fronteras queda excluido.")
    
    add_heading_3(doc, "Delimitación Conceptual y Temática")
    add_body_p(doc, "El fenómeno se investiga a través del prisma exclusivo de las cuatro dimensiones existenciales del mundo de la vida (Lebenswelt): Temporalidad (Lebenszeit), Espacialidad (Lebensraum), Corporalidad (Leib) y Relacionalidad (Mitwelt), más la síntesis estructural esencial.")
    
    add_heading_3(doc, "Delimitación Social y Poblacional")
    add_body_p(doc, "La unidad de estudio está constituida por personas usuarias mayores de 18 años, hispanohablantes o quechuahablantes, que acuden como usuarios externos a los servicios asistenciales del establecimiento. Se excluye al personal administrativo y asistencial, cuya vivencia laboral constituye un objeto de estudio independiente.")
    
    add_heading_3(doc, "Limitaciones Metodológicas del Estudio")
    add_bullet_item(doc, "Naturaleza no probabilística del diseño: ", "Los hallazgos fenomenológicos no persiguen la generalización estadística numérica de frecuencias poblacionales, sino la transferibilidad conceptual contextualizada según los criterios de rigor de Guba y Lincoln [12].")
    add_bullet_item(doc, "Efecto de reactividad y pudor: ", "La presencia de la investigadora y el uso de grabadoras de audio podrían generar reticencia inicial en ciertos informantes quechuahablantes. Dicha limitación se mitiga mediante una inmersión prolongada en el campo, el establecimiento de rapport empático y el diálogo en su lengua materna.")
    add_bullet_item(doc, "Condicionamiento biometeorológico: ", "Las variaciones estacionales de frío y lluvia pueden modular la vivencia somática, lo que demanda registrar notas de campo reflexivas precisas en cada sesión de entrevista.")
    
    add_heading_2(doc, "1.7. Viabilidad y consideraciones bioéticas")
    add_body_p(doc, "La presente investigación es plenamente viable puesto que cuenta con acceso institucional favorable a las instalaciones del Centro de Salud Belén, disponibilidad de recursos humanos calificados, respaldo académico tutorial de la cátedra de EN 486 (Dr. Manglio Aguirre Andrade) y financiamiento asegurado en su totalidad con recursos propios de la investigadora principal.")
    add_body_p(doc, "En estricta observancia del rigor bioético, el protocolo será sometido al Comité Institucional de Ética en Investigación (CIEI) de la Universidad Nacional de San Cristóbal de Huamanga (UNSCH) y coordinado formalmente con la Dirección Regional de Salud (DIRESA) de Ayacucho y la Red de Salud Huamanga antes de cualquier abordaje de campo.")
    add_body_p(doc, "El estudio se rige bajo los principios de la Declaración de Helsinki [15], las pautas internacionales del CIOMS [16] y el Informe Belmont:")
    add_bullet_item(doc, "Principio de Autonomía: ", "Se garantizará la voluntariedad absoluta a través de la suscripción formal del Consentimiento Informado (Anexo 4), explicándose de forma detallada y comprensible la naturaleza de la investigación en quechua chanka o castellano.")
    add_bullet_item(doc, "Principio de Beneficencia y No Maleficencia: ", "Se asegurará un ambiente de escucha digno y contenedor, evitando cualquier revictimización o angustia emocional durante el relato vivencial. Las entrevistas podrán suspenderse si el informante lo solicita.")
    add_bullet_item(doc, "Principio de Justicia y Confidencialidad: ", "Se mantendrá el anonimato absoluto mediante el uso de códigos alfanuméricos despersonalizados (ejemplo: Informante 01 -- [INF-01-MED]) en todas las etapas de transcripción, categorización en ATLAS.ti v9 y redacción del informe final.")
    # ==========================================================================
    # SECCIÓN 5: CAPÍTULO II - MARCO CONTEXTUAL
    # ==========================================================================
    sec_cap2 = doc.add_section(WD_SECTION.NEW_PAGE)
    sec_cap2.page_width = Cm(21.0)
    sec_cap2.page_height = Cm(29.7)
    sec_cap2.top_margin = Cm(3.0)
    sec_cap2.left_margin = Cm(3.5)
    sec_cap2.right_margin = Cm(2.5)
    sec_cap2.bottom_margin = Cm(2.5)
    sec_cap2.different_first_page_header_footer = False
    sec_cap2.header.is_linked_to_previous = False
    sec_cap2.footer.is_linked_to_previous = False
    sec_cap2.footer.paragraphs[0].text = ""
    set_section_pgnum(sec_cap2, start=None, fmt='decimal')
    setup_header(sec_cap2, "CAPÍTULO II: MARCO CONTEXTUAL")
    
    add_heading_1(doc, "CAPÍTULO II: MARCO CONTEXTUAL")
    add_body_p(doc, "En la investigación cualitativa bajo el diseño fenomenológico, el marco contextual no se reduce a una enumeración estática de datos cuantitativos ni a un inventario geográfico desvinculado; constituye la reconstrucción densa, profunda y holística del ambiente natural donde transcurre la vida cotidiana de las personas y donde se gesta el fenómeno estudiado [5, 6]. Para comprender la esencia de la experiencia vivida (Lebenswelt) de los usuarios respecto a la calidad de atención en el Centro de Salud Belén, es indispensable caracterizar la atmósfera física, cultural, social y sanitaria que envuelve, condiciona y dota de sentido a sus sentires, expectativas, frustraciones y valoraciones intersubjetivas [8, 4].")
    
    add_heading_2(doc, "2.1. Descripción geográfica, territorial y ambiental del área de estudio")
    add_body_p(doc, "La investigación se sitúa en el Centro de Salud Belén, establecimiento sanitario cabecera perteneciente a la Microred Huamanga de la Red de Salud Huamanga, adscrito a la Dirección Regional de Salud (DIRESA) de Ayacucho. El establecimiento se localiza en el emblemático e histórico barrio de Belén, en el sector céntrico-sur del distrito de Ayacucho, provincia de Huamanga, departamento de Ayacucho, a una altitud oficial de 2,760 metros sobre el nivel del mar. Sus coordenadas geográficas de emplazamiento corresponden a los 13°09'47'' de Latitud Sur y 74°13'28'' de Longitud Oeste.")
    add_body_p(doc, "El territorio bajo su responsabilidad sanitaria se asienta en la subcuenca del río Alameda, dentro de la cuenca principal del valle de Huamanga. Limita por el norte con el casco monumental y urbano central de la ciudad de Huamanga; por el sur con las urbanizaciones periurbanas del distrito de San Juan Bautista; por el este con las faldas del cerro Acuchimay y el distrito de Carmen Alto; y por el oeste con la prolongación de la avenida Mariscal Cáceres y el jirón Bellido. Su orografía presenta una topografía moderadamente accidentada, con calles estrechas, empinadas y trazados coloniales que imponen un esfuerzo físico considerable a los usuarios con movilidad reducida, gestantes y personas de la tercera edad que acuden a pie.")
    add_body_p(doc, "El clima del área es templado y seco durante las horas diurnas soleadas, con temperaturas promedio anuales que oscilan entre los 15 °C y 22 °C. No obstante, una característica meteorológica decisiva para la experiencia del usuario radica en la marcada oscilación térmica interdiurna: durante las madrugadas andinas (entre las 4:00 a.m. y las 6:30 a.m.), el termómetro desciende habitualmente a valores de entre 5 °C y 8 °C, llegando a extremos menores en la temporada seca y de heladas (junio a agosto). Asimismo, se presenta una temporada de precipitaciones pluviales intensas entre los meses de noviembre y marzo. Esta condición ambiental no constituye un dato accesorio, sino un determinante corpóreo directo que condiciona la expectativa del cuidado: el usuario que arriba al establecimiento con frío penetrante, fatiga somática y dolor acumulado demanda una acogida asistencial empática y cálida que alivie su vulnerabilidad.")
    add_body_p(doc, "En cuanto a la conectividad y movilidad urbana, el establecimiento cuenta con acceso a través de vías pavimentadas (Jr. Bellido, Jr. Arequipa y Av. Mariscal Cáceres) y rutas de transporte público (líneas 1, 3, 7 y 12, además de mototaxis). A pesar de ello, las limitaciones económicas obligan a numerosas familias de sectores altos a realizar traslados peatonales de 20 a 40 minutos cargando niños en mantas tradicionales (llicllas). Esta exigencia física y ambiental predispone un estado de agotamiento al momento de ingresar a los consultorios, transformando la calidez del trato, la prontitud del saludo y la dignidad del espacio de espera en elementos determinantes de la calidad del cuidado.")
    
    add_heading_2(doc, "2.2. Características demográficas y estructura poblacional")
    add_body_p(doc, "La jurisdicción sanitaria asignada al Centro de Salud Belén abarca una población diana proyectada de 18,450 habitantes según el Padrón Poblacional Oficial de la Red de Salud Huamanga (2025–2026). La pirámide poblacional se caracteriza por una base ensanchada en los primeros decenios de vida y una proporción creciente de personas adultas mayores, configurando un doble desafío asistencial de perfil materno-infantil y crónico-degenerativo. La distribución por etapas de vida es la siguiente:")
    add_bullet_item(doc, "Población Infantil (menores de 5 años): ", "Representa el 12.4% (2,288 niños), constituyendo el segmento con mayor intensidad de demanda en los consultorios preventivo-promocionales de Crecimiento y Desarrollo (CRED), tamizaje y tratamiento de anemia ferropénica, suplementación nutricional e inmunizaciones del esquema nacional regular.")
    add_bullet_item(doc, "Población Escolar y Adolescente (5 a 17 años): ", "Abarca el 19.6% (3,616 personas), demandantes de intervenciones de salud bucal, tamizaje visual, salud del escolar y orientación en salud sexual y reproductiva.")
    add_bullet_item(doc, "Población Joven y Adulta (18 a 59 años): ", "Constituye el 52.8% (9,742 personas), con un marcado predominio de mujeres en edad fértil (MEF) que acuden para control prenatal, despistaje de cáncer de cuello uterino, planificación familiar y atenciones médicas por patologías agudas.")
    add_bullet_item(doc, "Población Adulta Mayor (60 años a más): ", "Concentra el 15.2% (2,804 personas), un grupo en acelerado envejecimiento que presenta alta prevalencia de hipertensión arterial, diabetes mellitus tipo 2, artrosis y secuelas osteomusculares, requiriendo un acompañamiento asistencial continuo, trato digno y accesibilidad física preferencial.")
    add_body_p(doc, "Desde la mirada fenomenológica, resalta el fenómeno de la feminización de la demanda matutina: más del 70% de los usuarios que concurren a las salas de espera son mujeres. Se trata de madres de familia que acuden al cuidado y vigilancia de la salud de sus hijos menores, mujeres gestantes en controles prenatales y mujeres cuidadoras que asisten a adultos mayores dependientes. Esta dinámica impone sobre la mujer andina una sobrecarga cotidiana de roles, donde el tiempo de espera en el centro de salud compite directamente con sus jornadas de trabajo doméstico y comercial.")
    
    add_heading_2(doc, "2.3. Características socioculturales, lingüísticas y cosmovisión andina")
    add_body_p(doc, "El entorno comunitario de Belén se distingue por una singular riqueza identitaria y una densa trama de sociabilidad popular. Históricamente, el barrio de Belén ha sido cuna de ilustres artesanos populares de Huamanga, reconocidos internacionalmente por el arte del retablo ayacuchano, la imaginería religiosa en yeso y el tallado tradicional en piedra de Huamanga. En este espacio geográfico coexisten familias tradicionales residentes desde hace generaciones con oleadas de población migrante provenientes de las zonas rurales y provincias del centro-sur ayacuchano (Víctor Fajardo, Cangallo y Vilcas Huamán), muchas de ellas asentadas a raíz de la violencia sociopolítica de las décadas pasadas.")
    add_body_p(doc, "En el plano lingüístico, el 72% de la población usuaria es bilingüe activa en quechua chanka y castellano. En las generaciones de adultos mayores y en mujeres migrantes de sectores periféricos, el quechua constituye la lengua materna predominante y el vehículo exclusivo para verbalizar el sufrimiento corpóreo y anímico. Para la cosmovisión andina huamanguina, los padecimientos no son meras disfunciones biológicas aisladas, sino alteraciones del equilibrio integral del ser humano con su entorno, expresadas en nociones complejas como el dolor visceral (nanay), el estado de enfermedad integral (onqoy), la tristeza y aflicción acumulada (llakikuy), el enfriamiento del cuerpo (chiriyay) o los desórdenes espirituales derivados del espanto o susto (mancharisqa).")
    add_body_p(doc, "Las pautas de sociabilidad comunitaria exigen de manera indeclinable la observancia de códigos éticos y de cortesía tradicionales basados en el respeto mutuo (respetanakuy), el saludo formal y afectuoso (napaykuy) y la acogida hospitalaria (allin chaskiy) [18]. Cuando el personal de salud —frecuentemente hispanohablante monolingüe o condicionado por la prisa protocolar— prescinde del saludo, evita el contacto visual, interrumpe el relato del usuario o emplea un léxico excesivamente tecnicista, se produce un violento quiebre comunicativo. El usuario andino interpreta esta conducta como desdén, frialdad burocrática, soberbia institucional o discriminación clasista y étnica, replegándose en el silencio (upallay) y perdiendo la confianza en el sistema asistencial.")
    add_body_p(doc, "En cuanto al nivel de instrucción, el 48% de los jefes de familia cuenta con educación secundaria completa, un 26% posee secundaria incompleta o primaria, y un 8% (concentrado en adultos mayores) registra analfabetismo formal o funcional. Esta realidad representa un desafío crítico para la alfabetización en salud, pues muchos usuarios no logran descifrar las recetas médicas escritas, los calendarios de citas o la señalética institucional de los consultorios, dependiendo enteramente de la orientación verbal empática que el profesional de enfermería pueda suministrarle.")
    
    add_heading_2(doc, "2.4. Dinámica socioeconómica, vulnerabilidad y nivel de aseguramiento público")
    add_body_p(doc, "La actividad socioeconómica de los hogares usuarios del Centro de Salud Belén se inserta de manera mayoritaria en el régimen de la economía popular informal y de subsistencia diaria. La Población Económicamente Activa (PEA) labora preponderantemente en el comercio ambulatorio de verduras, frutas y alimentos preparados en el Mercado Central de Belén y su concurrida feria sabatina; en el transporte menor mediante la conducción de mototaxis; en faenas eventuales de construcción civil (albañilería y peonaje); en talleres artesanales independientes; y en el trabajo doméstico no remunerado o remunerado por jornada.")
    add_body_p(doc, "De acuerdo con los registros del Sistema de Focalización de Hogares (SISFOH), el 68.2% de las familias del ámbito de influencia se encuentra clasificada en situación de pobreza y pobreza extrema. En coherencia con este perfil de precariedad material, el 94.5% de la población usuaria se encuentra afiliada al régimen subsidiado del Seguro Integral de Salud (SIS).")
    add_body_p(doc, "Esta dependencia casi total del aseguramiento público estatal confiere una trascendencia vital a la gratuidad del servicio. Los usuarios acuden al centro asistencial bajo la certeza legítima de que el Estado proveerá tanto el acto médico como los medicamentos necesarios para su recuperación. Cuando la farmacia del establecimiento incurre en roturas de stock o desabastecimiento de medicamentos esenciales (como antibióticos de primera línea, antipiréticos, micronutrientes para la anemia o fármacos antihipertensivos), se suscita una crisis devastadora en la economía doméstica: el usuario se ve conminado a realizar un gasto de bolsillo imprevisto en las boticas y farmacias comerciales privadas que proliferan en los alrededores del centro de salud. Para una madre que vive de las ganancias del día a día en el mercado, destinar 20 o 30 soles a la compra de fármacos implica desfinanciar la alimentación diaria del hogar, vivencia que es experimentada con honda amargura, desesperanza e impotencia frente a la precariedad del sistema público.")
    
    add_heading_2(doc, "2.5. Características sanitarias, cartera de servicios y capacidad resolutiva institucional")
    add_body_p(doc, "El Centro de Salud Belén se encuentra formalmente categorizado como un establecimiento de salud del primer nivel de atención de nivel I-3 (sin internamiento hospitalario), desempeñando un papel neurálgico como cabecera de microred asistencial. Su misión operativa consiste en brindar atención integral ambulatoria, promocional, preventiva y de recuperación básica a la población vulnerable de su ámbito territorial y jurisdiccional.")
    add_body_p(doc, "A continuación se detalla la oferta prestacional y la distribución del talento humano asistencial disponible en el establecimiento:")
    
    # Tabla de Cartera de Servicios
    tbl_serv = doc.add_table(rows=1, cols=4)
    tbl_serv.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_table_borders(tbl_serv, color="808080", sz="4", val="single")
    headers_serv = ["Cartera de Servicios", "Actividades Asistenciales Principales", "Turno", "Personal Asignado"]
    widths_serv = [Cm(3.8), Cm(6.2), Cm(2.5), Cm(2.5)]
    for i, h_text in enumerate(headers_serv):
        c = tbl_serv.rows[0].cells[i]
        c.width = widths_serv[i]
        set_cell_margins(c, 80, 80, 100, 100)
        set_cell_shading(c, "EBF1F5")
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(h_text)
        format_run(r, size_pt=9.5, bold=True)
    
    data_serv = [
        ("Medicina General", "Consulta ambulatoria, diagnóstico clínico, prescripción médica y atención de morbilidades agudas y crónicas.", "Mañana / Tarde", "04 Médicos Cirujanos"),
        ("Enfermería (CRED e Inmunizaciones)", "Control del crecimiento y desarrollo del niño, tamizaje de anemia, vacunación del esquema regular y consejería nutricional.", "Mañana / Tarde", "06 Lic. en Enfermería"),
        ("Enfermería (TBC y No Transmisibles)", "Control de tuberculosis (PCT), administración de tratamiento DOTS, tamizaje de hipertensión arterial y diabetes mellitus.", "Mañana", "02 Lic. en Enfermería"),
        ("Obstetricia (Salud Reproductiva)", "Control prenatal, psicoprofilaxis obstétrica, monitoreo fetal básico, planificación familiar y tamizaje de cáncer ginecológico.", "Mañana / Tarde", "04 Lic. en Obstetricia"),
        ("Odontología", "Odontología preventiva y restauradora, exodoncias simples, profilaxis dental y fluorización tópica.", "Mañana / Tarde", "02 Cirujanos Dentistas"),
        ("Psicología", "Atención de salud mental, tamizaje de violencia familiar, intervención en crisis y consejería psicológica.", "Mañana", "02 Lic. en Psicología"),
        ("Tópico de Urgencias y Triaje", "Evaluación de signos vitales, curaciones menores, suturas simples, administración de inyectables y nebulizaciones.", "12 Horas continuas", "04 Técnicos de Enfermería"),
        ("Farmacia SIS", "Almacenamiento, custodia y dispensación ambulatoria de medicamentos del petitorio oficial para afiliados SIS.", "12 Horas continuas", "02 Técnicos de Farmacia"),
        ("Laboratorio Básico", "Toma de muestras biológicas, hematocrito/hemoglobina, frotis sanguíneo, examen de orina y baciloscopías.", "Mañana", "01 Tecnólogo / Técnico")
    ]
    for row_idx, (serv, act, turno, pers) in enumerate(data_serv):
        row = tbl_serv.add_row()
        vals = [serv, act, turno, pers]
        aligns = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER]
        bolds = [True, False, False, False]
        for c_idx, val in enumerate(vals):
            c = row.cells[c_idx]
            c.width = widths_serv[c_idx]
            set_cell_margins(c, 60, 60, 100, 100)
            if row_idx % 2 == 1:
                set_cell_shading(c, "F8F9FA")
            p = c.paragraphs[0]
            p.alignment = aligns[c_idx]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(val)
            format_run(r, size_pt=9, bold=bolds[c_idx])
    
    add_body_p(doc, "El perfil epidemiológico de la demanda ambulatoria está hegemonizado por las Infecciones Respiratorias Agudas (IRA), Enfermedades Diarreicas Agudas (EDA), parasitosis intestinal, anemia por deficiencia de hierro en niños menores de 36 meses, infecciones del tracto urinario (ITU) en mujeres en edad fértil y gestantes, lumbalgias mecánicas derivadas del trabajo físico, y patologías crónicas no transmisibles (hipertensión arterial y diabetes mellitus tipo 2) en adultos mayores.")
    add_body_p(doc, "No obstante su relevancia estratégica, el establecimiento enfrenta una serie de nudos críticos institucionales y barreras operativas que inciden directamente en la subjetividad del usuario:")
    add_num_item(doc, 1, "El elevado número de usuarios asignados por turno y la exigencia de metas de producción asistencial reducen drásticamente los minutos dedicados a cada consulta clínica ambulatoria (frecuentemente de 10 a 15 minutos). Esta premura temporal condiciona un ritmo de atención apresurado que limita la profundización en el diálogo clínico, restringe el espacio para la consejería de enfermería y dificulta la acogida empática de las dudas y temores del paciente.", bold_prefix="Sobrecarga asistencial y presión por rendimiento horario (tiempo de consulta abreviado): ")
    add_num_item(doc, 2, "El centro de salud funciona en un inmueble cuya infraestructura original ha sido sucesivamente adaptada, presentando pasillos estrechos, salas de espera techadas pero con escaso aislamiento térmico y bancas metálicas insuficientes. La tabiquería provisional entre consultorios adolece de aislamiento acústico, lo que provoca la filtración de conversaciones clínicas y quebranta el derecho a la privacidad e intimidad corporal y emocional de los pacientes.", bold_prefix="Limitaciones arquitectónicas y hacinamiento espacial: ")
    add_num_item(doc, 3, "Una fracción considerable del personal profesional y técnico se encuentra bajo regímenes de contratación temporal transitoria (CAS), lo que origina una continua rotación de profesionales que interrumpe la continuidad del vínculo afectivo con los pacientes de la comunidad. A ello se suma la sobrecarga de digitación en los sistemas informáticos (HIS-MINSA y formato FUA-SIS), la cual absorbe gran parte del tiempo de consulta, obligando al profesional a interactuar con la pantalla del computador en detrimento del contacto visual con el usuario.", bold_prefix="Inestabilidad laboral y sobrecarga burocrática: ")
    
    add_heading_2(doc, "2.6. Articulación dialéctica del escenario físico-social con las dimensiones existenciales del usuario (Lebenswelt)")
    add_body_p(doc, "En concordancia con los postulados epistemológicos de van Manen [8], Creswell [6] y Hernández-Sampieri y Mendoza [5], articulados con el modelo del Proceso Interpersonal de Donabedian [9, 10] y la Teoría del Cuidado Humano de Watson [11], el mundo de la vida (Lebenswelt) de los usuarios se manifiesta a través de cuatro existenciales universales donde se construye el significado de la calidad de atención:")
    
    # Tabla de Existenciales
    tbl_ex = doc.add_table(rows=1, cols=3)
    tbl_ex.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_table_borders(tbl_ex, color="808080", sz="4", val="single")
    headers_ex = ["Dimensión Existencial (Lebenswelt)", "Manifestación Empírica en el Contexto de Belén", "Significado Vivencial y Afectivo Construido por el Usuario"]
    widths_ex = [Cm(3.5), Cm(5.5), Cm(6.0)]
    for i, h_text in enumerate(headers_ex):
        c = tbl_ex.rows[0].cells[i]
        c.width = widths_ex[i]
        set_cell_margins(c, 80, 80, 100, 100)
        set_cell_shading(c, "EBF1F5")
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(h_text)
        format_run(r, size_pt=9.5, bold=True)
    
    data_ex = [
        ("Relacionalidad (Relación vivida con los otros)", 
         "Encuentro intersubjetivo entre el usuario quechua-mestizo vulnerable y el personal sanitario; presencia del saludo cordial (napaykuy) o su omisión; mirada a los ojos frente a la mirada fija en el monitor; paciencia explicativa o respuestas cortantes.", 
         "La calidad se define en la dignidad del trato interpersonal: el usuario busca acogida humana (allin chaskiy), respeto mutuo (respetanakuy) y comprensión en su propia lengua originaria. La calidez del cuidado de enfermería transforma la vulnerabilidad en alivio, confianza y gratitud."),
        ("Temporalidad (Tiempo vivido)", 
         "Puntualidad en el inicio de la atención; duración efectiva del contacto clínico en consultorio; lapsos de espera frente a una consulta médica de apenas 10 a 15 minutos; ritmo pausado para escuchar dudas.", 
         "Vivencia de la desvalorización del tiempo del paciente frente a la prisa protocolar. Una atención de calidad se experimenta cuando el profesional no muestra apuro, escucha sin interrumpir y dedica tiempo suficiente para clarificar el tratamiento."),
        ("Espacialidad (Espacio vivido)", 
         "Consultorios con tabiquería delgada que deja filtrar conversaciones clínicas; pasadizos congestionados; puertas entreabiertas durante la exploración médica; comodidad y limpieza del consultorio.", 
         "Sensación de exposición pública o resguardo de la intimidad. La calidad espacial exige consultorios cerrados, confortables y seguros donde el usuario sienta absoluta reserva y confidencialidad respecto a sus padecimientos íntimos."),
        ("Corporalidad (Cuerpo vivido)", 
         "Delicadeza y respeto al pudor corporal al desvestirse o ser examinado; tacto suave en inyecciones y curaciones de enfermería; respuesta afectuosa ante el llanto del niño o los dolores articulares del anciano.", 
         "El cuerpo doliente (nanay) no es un objeto mecánico de intervención, sino el asiento directo de la vulnerabilidad ontológica. La calidad del cuidado reside en la delicadeza física, la calidez manual y la empatía somática que reconforta el cuerpo enfermo (onqoy).")
    ]
    for row_idx, (dim, man, sig) in enumerate(data_ex):
        row = tbl_ex.add_row()
        vals = [dim, man, sig]
        aligns = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.JUSTIFY]
        bolds = [True, False, False]
        for c_idx, val in enumerate(vals):
            c = row.cells[c_idx]
            c.width = widths_ex[c_idx]
            set_cell_margins(c, 60, 60, 100, 100)
            if row_idx % 2 == 1:
                set_cell_shading(c, "F8F9FA")
            p = c.paragraphs[0]
            p.alignment = aligns[c_idx]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(val)
            format_run(r, size_pt=9, bold=bolds[c_idx])
    
    add_body_p(doc, "Esta articulación multidimensional confirma que, para el usuario del Centro de Salud Belén, la “calidad de atención” trasciende ampliamente los parámetros gerenciales o el abastecimiento técnico; constituye una experiencia existencial totalizante donde el cuerpo aterido, el tiempo sacrificado, el espacio habitado y el trato humano recibido configuran la esencia de su satisfacción o sufrimiento frente al sistema público de salud.")
    
    # ==========================================================================
    # MÓDULOS EN ESPERA: Se activarán cuando el usuario indique:
    # "DEJA QUE LOS DEMÁS DESPUES DEL CAP 2 SEAN VISIBLES"
    # ==========================================================================
    MOSTRAR_DESPUES_DE_CAP2 = False
    if MOSTRAR_DESPUES_DE_CAP2:
        # ==========================================================================
        # SECCIÓN 6: CAPÍTULO III - ABORDAJE TEÓRICO
        # ==========================================================================
        sec_cap3 = doc.add_section(WD_SECTION.NEW_PAGE)
        sec_cap3.page_width = Cm(21.0)
        sec_cap3.page_height = Cm(29.7)
        sec_cap3.top_margin = Cm(3.0)
        sec_cap3.left_margin = Cm(3.5)
        sec_cap3.right_margin = Cm(2.5)
        sec_cap3.bottom_margin = Cm(2.5)
        sec_cap3.different_first_page_header_footer = False
        sec_cap3.header.is_linked_to_previous = False
        sec_cap3.footer.is_linked_to_previous = False
        sec_cap3.footer.paragraphs[0].text = ""
        set_section_pgnum(sec_cap3, start=None, fmt='decimal')
        setup_header(sec_cap3, "CAPÍTULO III: ABORDAJE TEÓRICO")

        add_heading_1(doc, "CAPÍTULO III: ABORDAJE TEÓRICO")
        add_heading_2(doc, "3.1. Antecedentes de estudio")

        add_heading_3(doc, "3.1.1. Antecedentes internacionales")
        add_bullet_item(doc, "Gómez-Restrepo y colaboradores (2023, Colombia): ", "Realizaron una investigación cualitativa fenomenológica titulada “Experiencia del usuario sobre la calidad asistencial en el primer nivel de atención”. Analizaron 18 entrevistas a profundidad con usuarios de centros de salud urbanos, concluyendo que la percepción de calidad está dominada por la calidez comunicativa y la empatía del personal. Demostraron que el trato humanizado compensa percibidamente las deficiencias del ambiente físico y reduce la frustración asociada a los tiempos de espera prolongados.")
        add_bullet_item(doc, "Pérez y Silva (2022, Ecuador): ", "En su estudio cualitativo “Significados de la atención de enfermería desde la mirada del paciente externo”, concluyeron que los usuarios valoran principalmente el saludo, la escucha atenta y la mirada a los ojos por parte del personal de enfermería, considerando estos gestos como indicadores fundamentales de respeto y competencia profesional.")

        add_heading_3(doc, "3.1.2. Antecedentes nacionales")
        add_bullet_item(doc, "Mamani y Quispe (2024, Lima): ", "En su investigación cualitativa “Percepción de la calidad del cuidado de enfermería en usuarios de establecimientos de salud del primer nivel”, determinaron que los usuarios asocian la mala calidad no al diagnóstico médico sino a la frialdad en la atención, el uso excesivo de modismos técnicos y la falta de espacio para formular preguntas sobre sus medicamentos.")
        add_bullet_item(doc, "Córdova y Reyes (2023, Arequipa): ", "Desarrollaron el estudio “Tiempos de espera y trato humano: Un estudio cualitativo en centros de salud públicos”. Identificaron que la larga espera en filas desde tempranas horas de la mañana genera una predisposición negativa en el usuario, la cual solo logra disiparse cuando el profesional de enfermería brinde una acogida empática e informativa.")

        add_heading_3(doc, "3.1.3. Antecedentes regionales y locales")
        add_bullet_item(doc, "Rojas (2023, Ayacucho): ", "En su tesis cualitativa “Percepción intercultural de la calidad de atención en usuarios de centros de salud de la Red Huamanga”, realizada con usuarios quechuahablantes, concluyó que la barrera lingüística es el principal factor de insatisfacción. Evidenció que cuando el personal de salud o enfermería responde o saluda en quechua, la percepción de calidad y confianza institucional se incrementa significativamente.")
        add_bullet_item(doc, "García (2022, Ayacucho - UNSCH): ", "En su trabajo de investigación sobre “Cuidado humanizado de enfermería en el primer nivel de atención”, reportó que la sobrecarga asistencial del personal suele minar la paciencia en el trato, afectando la percepción que tiene la madre o el adulto mayor sobre la atención brindada en los servicios preventivo-promocionales.")

        add_heading_2(doc, "3.2. Base teórica")
        add_heading_3(doc, "Modelo de Calidad de Salud de Avedis Donabedian")
        add_body_p(doc, "Constituye el marco conceptual de referencia para evaluar la calidad sanitaria. Donabedian establece tres dimensiones entrelazadas:")
        add_num_item(doc, 1, "Engloba los recursos materiales (infraestructura, equipamiento, insumos), financieros y humanos (número y calificación del personal).", bold_prefix="Estructura: ")
        add_num_item(doc, 2, "Refiere a las actividades asistenciales interpersonales y técnicas que ocurren entre el profesional de salud y el paciente.", bold_prefix="Proceso: ")
        add_num_item(doc, 3, "Evalúa las modificaciones en el estado de salud del paciente y la satisfacción subjetiva obtenida tras el servicio.", bold_prefix="Resultado: ")

        add_heading_3(doc, "Teoría del Cuidado Humano de Jean Watson")
        add_body_p(doc, "Sostiene que el cuidado es el núcleo filosófico y práctico de la enfermería. Define al cuidado como un proceso transpersonal e intersubjetivo que busca armonizar la mente, el cuerpo y el alma del ser humano. Watson plantea la necesidad de cultivar relaciones de ayuda y confianza, caracterizadas por la empatía, la sensibilidad, el respeto y la escucha activa.")

        add_heading_3(doc, "Fundamentos del Cuidado Fenomenológico en Enfermería: Patricia Benner y Holly Wilson")
        add_body_p(doc, "En el campo de la enfermería y las ciencias de la salud, la fenomenología constituye el diseño por excelencia para aprehender la esencia íntima del cuidado, la vulnerabilidad humana y el sufrimiento (Norlyk y Harder, 2010; Sampieri y Mendoza, 2018). Como postula Patricia Benner (2008), la práctica del cuidado de enfermería es fundamentalmente una práctica encarnada (embodied practice), donde el cuerpo del paciente no es un objeto mecánico, sino el vehículo a través del cual la persona vivencia el dolor, la angustia, el frío y el alivio. Benner fundamenta que la calidad asistencial emerge en el encuentro intersubjetivo, donde la intuición clínica y la presencia sensible de la enfermera permiten descifrar el sentir del paciente más allá de los signos biomédicos.")
        add_body_p(doc, "Asimismo, Holly Skodol Wilson (2007) y Norlyk y Harder (2010) señalan que el diseño fenomenológico en salud tiene por propósito primordial describir la esencia de las experiencias de los pacientes, exigiendo que el investigador 'haga a un lado' (bracketing o epojé) sus presupuestos teóricos para dar cabida a las narrativas auténticas de los usuarios. Esta concepción se engarza de manera natural con el contexto del Centro de Salud Belén, donde las vivencias de las madres en CRED, los adultos mayores y las familias quechuahablantes requieren ser develadas desde su propia cotidianidad.")

        add_heading_2(doc, "3.3. Categorías de estudio (Matriz cualitativa y fenomenológica)")
        add_body_p(doc, "En concordancia con el paradigma inductivo interpretativo y las directrices metodológicas de Hernández-Sampieri y Mendoza (2018, p. 494), en el diseño fenomenológico las unidades de significado se estructuran dialécticamente en: (a) Categorías esenciales o comunes de las narrativas (vivencias compartidas por la colectividad usuaria respecto a madrugadas, espera, empatía y quechua); y (b) Categorías diferentes o divergentes de las narrativas (valoraciones y matices particulares disonantes según edad, servicio o condición socioeconómica). A continuación se presenta la matriz comprensiva con las definiciones operativas, códigos sugeridos y preguntas orientadoras:")

        # TABLA DE CATEGORIZACIÓN CUALITATIVA FENOMENOLÓGICA (SAMPIERI)
        tbl_cat = doc.add_table(rows=1, cols=5)
        tbl_cat.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_table_borders(tbl_cat, color="808080", sz="4", val="single")

        headers_cat = ["Concepto Potencial", "Dimensión Existencial", "Definición Conceptual / Significado", "Categorías Emergentes y Códigos", "Pregunta Guía de Profundización"]
        col_widths_cat = [Cm(2.4), Cm(2.6), Cm(3.5), Cm(4.0), Cm(3.5)]

        hdr_cells = tbl_cat.rows[0].cells
        for i, title in enumerate(headers_cat):
            hdr_cells[i].text = title
            hdr_cells[i].width = col_widths_cat[i]
            set_cell_margins(hdr_cells[i], top=120, bottom=120, left=140, right=140)
            set_cell_shading(hdr_cells[i], "E8EEF5")
            for p in hdr_cells[i].paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    format_run(r, size_pt=9.5, bold=True)

        cat_rows_data = [
            ("CALIDAD DE ATENCIÓN Y CUIDADO HUMANO\n\n(Encuentro intersubjetivo y vivencia del usuario en el C.S. Belén)",
             "1. Temporalidad (Tiempo vivido)",
             "Duración efectiva del contacto clínico en consultorio, ritmo pausado vs. prisa protocolar y dedicación de tiempo para escuchar y explicar con calma.",
             "• Esencia compartida: Desazón ante consultas apresuradas de escasos minutos; bienestar profundo cuando el profesional dialoga con paciencia y sin apuro.\n• Variaciones individuales: Madres en CRED priorizan el tiempo dedicado a la consejería; adultos mayores demandan calma en la explicación de sus recetas.",
             "¿Cómo experimentó el tiempo que el personal de salud le dedicó en la consulta? ¿Sintió que la atención fue apresurada o con la calma necesaria?"),
            ("",
             "2. Espacialidad (Espacio vivido)",
             "Vivencia de la privacidad acústica y visual dentro del consultorio, confort físico de las salas y resguardo estricto de la confidencialidad clínica.",
             "• Esencia compartida: Incomodidad e inhibición ante tabiquerías delgadas que filtran voces; tranquilidad y confianza en ambientes cerrados y dignos.\n• Variaciones individuales: Consultorios de Obstetricia y CRED demandan mayor privacidad física que salas de triaje general.",
             "¿Sintió que el consultorio le brindó la intimidad, reserva y confidencialidad necesaria para dialogar y ser examinado(a)?"),
            ("",
             "3. Corporalidad (Cuerpo vivido)",
             "Sentires somáticos durante el acto de cuidado: respeto al pudor físico al desvestirse, delicadeza en procedimientos y contención del dolor somático (nanay).",
             "• Esencia compartida: Confianza y alivio ante el tacto suave y delicado en curaciones e inyecciones; consuelo humano frente al cuerpo adolorido (onqoy).\n• Variaciones individuales: Madres sensibles al trato corporal hacia sus hijos; adultos mayores valoran la gentileza física ante sus limitaciones motrices.",
             "Al ser examinado(a) físicamente o recibir curaciones, ¿cómo sintió que el personal trató su cuerpo y su dolor somático?"),
            ("",
             "4. Relacionalidad (Relación humana)",
             "Encuentro intersubjetivo entre el usuario y el equipo asistencial: calidez de la mirada, acogida cordial (allin chaskiy), respeto mutuo (respetanakuy) y lengua quechua.",
             "• Esencia compartida: Primacía del contacto visual sobre la pantalla del computador; consuelo al ser escuchado con ternura y en quechua chanka.\n• Variaciones individuales: Reconocimiento de alta calidez en Enfermería frente a mayor prisa en turnos médicos de alta demanda.",
             "¿Cómo fue el trato interpersonal y el encuentro humano con el equipo de salud? ¿Sintió calidez, mirada atenta y escucha en su lengua originaria?"),
            ("",
             "5. Esencia Holística de la Calidad",
             "Significado integral y valorativo que el usuario otorga al cuidado humanizado, digno y resolutivo en el primer nivel asistencial.",
             "• Esencia compartida: Calidad es ser reconocido como persona humana integral, conjugando destreza técnica con calidez afectiva incondicional.\n• Variaciones individuales: Adultos mayores priorizan comprensión y paciencia; personas jóvenes valoran la claridad diagnóstica y la empatía en el trato.",
             "Al reflexionar sobre la atención recibida, ¿qué significa para usted recibir un cuidado de salud con auténtica calidad humana?")
        ]

        for row_data in cat_rows_data:
            row = tbl_cat.add_row()
            for idx, text in enumerate(row_data):
                cell = row.cells[idx]
                cell.text = text
                cell.width = col_widths_cat[idx]
                set_cell_margins(cell, top=100, bottom=100, left=130, right=130)
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(2)
                    for r in p.runs:
                        format_run(r, size_pt=9.5, bold=(idx == 0 or idx == 1))

        # ==========================================================================
        # SECCIÓN 7: CAPÍTULO IV - METODOLOGÍA CUALITATIVA
        # ==========================================================================
        sec_cap4 = doc.add_section(WD_SECTION.NEW_PAGE)
        sec_cap4.page_width = Cm(21.0)
        sec_cap4.page_height = Cm(29.7)
        sec_cap4.top_margin = Cm(3.0)
        sec_cap4.left_margin = Cm(3.5)
        sec_cap4.right_margin = Cm(2.5)
        sec_cap4.bottom_margin = Cm(2.5)
        sec_cap4.different_first_page_header_footer = False
        sec_cap4.header.is_linked_to_previous = False
        sec_cap4.footer.is_linked_to_previous = False
        sec_cap4.footer.paragraphs[0].text = ""
        set_section_pgnum(sec_cap4, start=None, fmt='decimal')
        setup_header(sec_cap4, "CAPÍTULO IV: METODOLOGÍA FENOMENOLÓGICA")

        add_heading_1(doc, "CAPÍTULO IV: METODOLOGÍA: DISEÑO FENOMENOLÓGICO DE INVESTIGACIÓN")
        add_heading_2(doc, "4.1. Fundamentación filosófica y epistemológica del diseño fenomenológico")
        add_body_p(doc, "La investigación se inserta en el paradigma interpretativo-naturalista y adopta un diseño cualitativo fenomenológico (Hernández-Sampieri y Mendoza, 2018; Creswell, 2013). Como señalan Roberto Hernández-Sampieri y Christian Mendoza (2018), la fenomenología no es únicamente una técnica de indagación, sino simultáneamente una filosofía, un enfoque epistemológico y un diseño metodológico riguroso cuyo origen se remonta al matemático y filósofo Edmund Husserl (1859-1938). Su propósito principal consiste en explorar, describir y comprender las experiencias de las personas con respecto a un determinado fenómeno y descubrir los elementos en común (la esencia compartida) de tales vivencias.")
        add_bullet_item(doc, "Diferenciación frente a la Teoría Fundamentada: ", "En la teoría fundamentada se busca generar un modelo teórico abstracto a partir de las perspectivas de los participantes; en contraste, en la fenomenología los investigadores trabajan directamente con las unidades o declaraciones de los participantes y sus vivencias, más que abstraerlas para crear un modelo basado en interpretaciones preconcebidas (Hernández-Sampieri y Mendoza, 2018, p. 493). Los datos provienen de sentimientos, emociones, razonamientos y percepciones corporales directas.")
        add_bullet_item(doc, "Diferenciación frente al Diseño Narrativo: ", "Mientras que el diseño narrativo se enfoca en la conexión o sucesión cronológica de eventos (el punto de vista biográfico o temporal), el diseño fenomenológico se enfoca de manera exclusiva en la esencia de la experiencia compartida (lo que vivenciaron y cómo lo vivenciaron).")
        add_bullet_item(doc, "Diferenciación frente a Etnografía e Investigación-Acción: ", "No persigue describir patrones macro-culturales de una comunidad a lo largo de años (etnografía), ni implementar intervenciones experimentales directas (investigación-acción), sino capturar con fidelidad la esencia íntima del fenómeno asistencial vivido.")
        add_bullet_item(doc, "Articulación de dos vertientes fenomenológicas: ", "Integra armónicamente la fenomenología empírica o trascendental (Moustakas, 1994; Creswell, 2013; Wilson, 2007), donde la investigadora aparta mediante epojé sus propios presupuestos para describir la vivencia pura del paciente; y la fenomenología hermenéutica (van Manen, 1990), que permite interpretar el sentido latente de las narrativas en su matriz sociocultural andina quechua chanka.")

        add_heading_2(doc, "4.2. Los cuatro existenciales de la experiencia humana vivida (van Manen y Sampieri)")
        add_body_p(doc, "Siguiendo a Hernández-Sampieri y Mendoza (2018, p. 494) y Max van Manen (1990), el diseño fenomenológico contextualiza las experiencias en cuatro dimensiones fundamentales del mundo de la vida (Lebenswelt):")
        add_num_item(doc, 1, "El momento en que ocurrieron los hechos y el tiempo percibido. En el C.S. Belén se manifiesta en el madrugar a las 4:00 am, la angustia previa a la apertura de ventanilla, el tiempo detenido en las bancas de espera y la sensación de sosiego o prisa en la consulta.", bold_prefix="Temporalidad (el tiempo vivido): ")
        add_num_item(doc, 2, "El lugar físico y ambiental donde sucedieron las vivencias: la frialdad de la acera matutina, el hacinamiento en pasillos, la rigidez de asientos y la reserva o vulneración de la intimidad dentro del consultorio.", bold_prefix="Espacialidad (el espacio vivido): ")
        add_num_item(doc, 3, "La dimensión somática de las personas que vivencian el fenómeno: el frío matutino calando el cuerpo, el cansancio muscular por horas de pie, la fatiga del ayuno, el llanto del lactante y el dolor somático que motiva la consulta médica.", bold_prefix="Corporalidad (el cuerpo vivido): ")
        add_num_item(doc, 4, "Los lazos interpersonales y encuentros intersubjetivos gestados en la atención: el saludo cordial (napaykuy), el contacto visual, la empatía, la calidez o frialdad del personal y el diálogo fluido en quechua chanka.", bold_prefix="Relacionalidad (la relación humana vivida): ")

        add_heading_2(doc, "4.3. Postura epistemológica, reflexividad y reducción fenomenológica (Epojé)")
        add_body_p(doc, "En contraste con la neutralidad aséptica positivista, la investigadora asume un rol de proximidad sensible y reflexividad rigurosa guiada por tres pilares:")
        add_bullet_item(doc, "Práctica de la Reducción Fenomenológica (Epojé / Bracketing): ", "Siguiendo a Husserl (1949), Moustakas (1994) y Creswell (2013), la investigadora suspende reflexivamente sus propios juicios previos, experiencias previas como estudiante de enfermería y prejuicios clínicos, para recibir los relatos de los usuarios con total apertura y pureza fenoménica.")
        add_bullet_item(doc, "Construcción de Confianza (Rapport): ", "Generar un vínculo horizontal de respeto mutuo y calidez humana indispensable para que los informantes exterioricen vivencias íntimas sin temor a ser juzgados.")
        add_bullet_item(doc, "Empatía Intercultural Andina: ", "Comprender el sentir del usuario desde su propia cosmovisión y condición socioeconómica, dialogando fluidamente en quechua chanka y castellano.")

        add_heading_2(doc, "4.4. Selección del contexto, participantes informantes y saturación teórica")
        add_body_p(doc, "Siguiendo a Hernández-Sampieri y Mendoza (2018, p. 493) y Creswell (2013), en el diseño fenomenológico los participantes se seleccionan porque han experimentado en carne propia el fenómeno de interés bajo estudio. Se adopta un muestreo cualitativo no probabilístico intencional, regulado por el principio de saturación teórica (Esbensen et al., 2008), estimándose una muestra referencial de 12 a 18 informantes clave.", bold_prefix="Criterio de Selección y Muestreo: ")
        add_heading_3(doc, "Criterios de Inclusión:")
        add_num_item(doc, 1, "Usuarios de ambos sexos, mayores de 18 años, que hayan recibido atención en Medicina, Enfermería (CRED, Inmunizaciones), Obstetricia u Odontología del C.S. Belén en los últimos tres meses.")
        add_num_item(doc, 2, "Participación estrictamente voluntaria formalizada mediante la firma o huella dactilar del Consentimiento Informado (Anexo 4).")
        add_num_item(doc, 3, "Capacidad de verbalizar y comunicar sus vivencias y emociones en español o quechua chanka.")
        add_heading_3(doc, "Criterios de Exclusión:")
        add_num_item(doc, 1, "Personas con compromiso neurológico agudo o alteraciones del lenguaje que impidan el diálogo fluido.")
        add_num_item(doc, 2, "Usuarios que acudan exclusivamente a realizar gestiones burocráticas breves en ventanilla sin haber recibido atención clínica asistencial.")

        add_heading_2(doc, "4.5. Técnicas e instrumentos de recolección de vivencias")
        add_bullet_item(doc, "Entrevista Fenomenológica en Profundidad: ", "Técnica central (45 a 60 minutos por sesión) conducida mediante preguntas abiertas diseñadas para reconstruir el significado, la estructura y la esencia de la vivencia asistencial (Anexo 1) (Creswell, 2013; Norlyk y Harder, 2010).")
        add_bullet_item(doc, "Observación Reflexiva no Participante: ", "Registro detallado de la dinámica ambiental, tiempos de espera, gestos corporales y clima relacional en salas de espera y consultorios.")
        add_bullet_item(doc, "Diario de Campo y Notas de Inmersión: ", "Registro sistemático de impresiones inmediatas, silencios, inflexiones de voz y notas de reflexividad pos-entrevista.")
        add_bullet_item(doc, "Triangulación de Actores Clave: ", "Articulando el modelo pedagógico del Dr. Manglio Aguirre Andrade, se triangulan tres perspectivas: (a) Usuarios externos / pacientes (vivencia sentida directa); (b) Personal asistencial de salud y enfermería (retos operativos del servicio); y (c) Acompañantes y actores comunitarios (mirada colectiva y social).")
        add_heading_2(doc, "4.6. Criterios de rigor científico cualitativo (Hernández-Sampieri, Cap. 14)")
        add_body_p(doc, "En la investigación cualitativa se reemplazan radicalmente los conceptos positivistas de validez y confiabilidad estadística por los cuatro criterios de rigor cualitativo formulados por Roberto Hernández-Sampieri, Carlos Fernández-Collado y Pilar Baptista-Lucio (2014, Cap. 14), articulados con los fundamentos de Lincoln y Guba (1985):")

        # TABLA DE CRITERIOS DE RIGOR
        tbl_rig = doc.add_table(rows=1, cols=3)
        tbl_rig.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_table_borders(tbl_rig, color="808080", sz="4", val="single")
        hdr_rig = ["Criterio de Rigor (Sampieri)", "Definición Metodológica", "Estrategia Operativa de Aseguramiento"]
        w_rig = [Cm(3.8), Cm(5.6), Cm(6.0)]
        for i, t in enumerate(hdr_rig):
            cell = tbl_rig.rows[0].cells[i]
            cell.text = t
            cell.width = w_rig[i]
            set_cell_margins(cell, 100, 100, 120, 120)
            set_cell_shading(cell, "E8EEF5")
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    format_run(r, size_pt=9.5, bold=True)

        rig_data = [
            ("1. Dependencia (Consistencia Cualitativa)",
             "Coherencia interna y estabilidad en los procedimientos de recolección y sistematización de los testimonios.",
             "• Pistas de auditoría transparentes.\n• Registro exhaustivo en la bitácora de campo.\n• Protocolo estandarizado de transcripción fonética y codificación."),
            ("2. Credibilidad (Validez Interna)",
             "Fidelidad con la que las interpretaciones reflejan los significados reales vivenciados por los participantes.",
             "• Triangulación múltiple de métodos y fuentes.\n• Transcripción íntegra palabra por palabra (verbatim).\n• Verificación y chequeo con los participantes (member checking)."),
            ("3. Transferencia (Aplicabilidad Contextual)",
             "Posibilidad de transferir los significados a otros contextos con características ecológicas y sociales análogas.",
             "• Descripción densa y contextualizada del C.S. Belén (Capítulo II).\n• Caracterización sociocultural y lingüística detallada de los informantes."),
            ("4. Confirmabilidad (Neutralidad Reflexiva)",
             "Garantía de que los hallazgos surgen de las vivencias reales y no de sesgos o prejuicios del investigador.",
             "• Práctica rigurosa de la epojé o bracketing fenomenológico.\n• Memos analíticos reflexivos continuos.\n• Auditoría externa y revisión crítica con el asesor Dr. Manglio Aguirre.")
        ]
        for row_d in rig_data:
            row = tbl_rig.add_row()
            for idx, text in enumerate(row_d):
                c = row.cells[idx]
                c.text = text
                c.width = w_rig[idx]
                set_cell_margins(c, 80, 80, 120, 120)
                for p in c.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p.paragraph_format.line_spacing = 1.15
                    for r in p.runs:
                        format_run(r, size_pt=9.0, bold=(idx == 0))

        add_heading_2(doc, "4.7. Procedimiento sistemático de análisis cualitativo fenomenológico (Sampieri, Cap. 14 y 15)")
        add_body_p(doc, "El análisis se rige por las acciones secuenciales de la Figura 15.14 de Hernández-Sampieri y Mendoza (2018, p. 495), articuladas con los métodos de Colaizzi (1978), Moustakas (1994) y Creswell (2013):")
        add_num_item(doc, 1, "Conversión exacta palabra por palabra de los audios a texto digital dentro de las 24 horas posteriores a la entrevista, preservando pausas y giros en quechua chanka.", bold_prefix="Transcripción íntegra y fonética (verbatim): ")
        add_num_item(doc, 2, "Inmersión lectora reflexiva en la totalidad de las descripciones para aprehender el sentido global de las vivencias.", bold_prefix="Lectura holística general: ")
        add_num_item(doc, 3, "Detección de todas las frases y declaraciones clave vinculadas a la experiencia asistencial, otorgando igual valor epistémico inicial a cada testimonio.", bold_prefix="Identificación de unidades de análisis y horizontalización (horizonalization): ")
        add_num_item(doc, 4, "Estructuración en la matriz analítica de Sampieri y Mendoza (2018, p. 494): (a) Categorías esenciales o comunes (temas compartidos por todos los participantes); y (b) Categorías diferentes o divergentes (apreciaciones y valoraciones disonantes según edad o servicio).", bold_prefix="Generación de categorías esenciales vs. divergentes: ")
        add_num_item(doc, 5, "Elaboración de la descripción textural (lo que experimentaron objetivamente) y la descripción estructural (cómo y bajo qué condiciones de temporalidad, espacialidad, corporalidad y relacionalidad lo vivenciaron).", bold_prefix="Construcción de descripciones texturales y estructurales: ")
        add_num_item(doc, 6, "Integración armónica de las descripciones en una narrativa densa que sintetiza la esencia compartida del fenómeno de la calidad asistencial y el acceso.", bold_prefix="Síntesis de la esencia compartida del fenómeno: ")
        add_num_item(doc, 7, "Devolución de la síntesis a una muestra de participantes informantes para constatar que el reporte refleje fielmente el núcleo de sus experiencias vividas.", bold_prefix="Validación testimonial con los participantes (member checking): ")

        # ==========================================================================
        # SECCIÓN 8: CAPÍTULO V - ASPECTOS ADMINISTRATIVOS
        # ==========================================================================
        sec_cap5 = doc.add_section(WD_SECTION.NEW_PAGE)
        sec_cap5.page_width = Cm(21.0)
        sec_cap5.page_height = Cm(29.7)
        sec_cap5.top_margin = Cm(3.0)
        sec_cap5.left_margin = Cm(3.5)
        sec_cap5.right_margin = Cm(2.5)
        sec_cap5.bottom_margin = Cm(2.5)
        sec_cap5.different_first_page_header_footer = False
        sec_cap5.header.is_linked_to_previous = False
        sec_cap5.footer.is_linked_to_previous = False
        sec_cap5.footer.paragraphs[0].text = ""
        set_section_pgnum(sec_cap5, start=None, fmt='decimal')
        setup_header(sec_cap5, "CAPÍTULO V: ASPECTOS ADMINISTRATIVOS")

        add_heading_1(doc, "CAPÍTULO V: ASPECTOS ADMINISTRATIVOS")
        add_heading_2(doc, "5.1. Asignación de recursos humanos")
        add_bullet_item(doc, "Investigadora principal: ", "01 (Gresly Lucero Pariona Palomino, Estudiante de la E.P. Enfermería).")
        add_bullet_item(doc, "Asesor del proyecto: ", "Dr. Manglio Aguirre Andrade (Docente de la FCS - UNSCH).")
        add_bullet_item(doc, "Personal de apoyo: ", "02 (Transcripción e intérprete de quechua).")

        add_heading_2(doc, "5.2. Presupuesto")

        # Tabla 1: Pago de Servicios de Personal
        p = doc.add_paragraph()
        r = p.add_run("PAGO DE SERVICIOS DE PERSONAL")
        format_run(r, size_pt=11, bold=True)
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)

        tbl_p1 = doc.add_table(rows=1, cols=5)
        tbl_p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_table_borders(tbl_p1, color="808080", sz="4", val="single")
        hdr_p1 = ["SERVICIOS", "U.M", "CANTIDAD", "C. UNIT. S/.", "TOT. S/."]
        w_p = [Cm(6.5), Cm(2.0), Cm(2.3), Cm(2.6), Cm(2.6)]
        for i, t in enumerate(hdr_p1):
            cell = tbl_p1.rows[0].cells[i]
            cell.text = t
            cell.width = w_p[i]
            set_cell_margins(cell, 100, 100, 120, 120)
            set_cell_shading(cell, "F0F0F0")
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    format_run(r, size_pt=9.5, bold=True)

        rows_p1 = [
            ("Coasesor / Especialista temático", "Mes", "04", "500.00", "2000.00"),
            ("Personal de apoyo (02)", "Mes", "02", "500.00", "1000.00"),
            ("SUB TOTAL", "", "", "", "3000.00")
        ]
        for row_d in rows_p1:
            row = tbl_p1.add_row()
            for idx, text in enumerate(row_d):
                c = row.cells[idx]
                c.text = text
                c.width = w_p[idx]
                set_cell_margins(c, 80, 80, 120, 120)
                if row_d[0] == "SUB TOTAL":
                    set_cell_shading(c, "F5F5F5")
                for p in c.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if idx >= 3 else (WD_ALIGN_PARAGRAPH.CENTER if idx in [1,2] else WD_ALIGN_PARAGRAPH.LEFT)
                    for r in p.runs:
                        format_run(r, size_pt=9.5, bold=(row_d[0] == "SUB TOTAL"))

        # Tabla 2: Programación de Otros Servicios
        p = doc.add_paragraph()
        r = p.add_run("PROGRAMACIÓN DE OTROS SERVICIOS")
        format_run(r, size_pt=11, bold=True)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)

        tbl_p2 = doc.add_table(rows=1, cols=5)
        tbl_p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_table_borders(tbl_p2, color="808080", sz="4", val="single")
        for i, t in enumerate(hdr_p1):
            cell = tbl_p2.rows[0].cells[i]
            cell.text = t
            cell.width = w_p[i]
            set_cell_margins(cell, 100, 100, 120, 120)
            set_cell_shading(cell, "F0F0F0")
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    format_run(r, size_pt=9.5, bold=True)

        rows_p2 = [
            ("Impresiones y fotocopias", "Hojas", "1000", "0.20", "200.00"),
            ("Internet y conectividad", "Hora", "200", "1.00", "200.00"),
            ("Procesamiento y codificación cualitativa", "Unidad", "01", "500.00", "500.00"),
            ("Empastado de tesis/proyectos", "Unidad", "05", "20.00", "100.00"),
            ("Traslado y movilidad local", "Días", "20", "10.00", "200.00"),
            ("SUB TOTAL", "", "", "", "1200.00")
        ]
        for row_d in rows_p2:
            row = tbl_p2.add_row()
            for idx, text in enumerate(row_d):
                c = row.cells[idx]
                c.text = text
                c.width = w_p[idx]
                set_cell_margins(c, 80, 80, 120, 120)
                if row_d[0] == "SUB TOTAL":
                    set_cell_shading(c, "F5F5F5")
                for p in c.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if idx >= 3 else (WD_ALIGN_PARAGRAPH.CENTER if idx in [1,2] else WD_ALIGN_PARAGRAPH.LEFT)
                    for r in p.runs:
                        format_run(r, size_pt=9.5, bold=(row_d[0] == "SUB TOTAL"))

        # Tabla 3: Adquisición de Materiales
        p = doc.add_paragraph()
        r = p.add_run("ADQUISICIÓN DE MATERIALES Y OTROS")
        format_run(r, size_pt=11, bold=True)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)

        tbl_p3 = doc.add_table(rows=1, cols=5)
        tbl_p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_table_borders(tbl_p3, color="808080", sz="4", val="single")
        hdr_mat = ["MATERIALES", "U.M", "CANTIDAD", "C. UNIT. S/.", "TOT. S/."]
        for i, t in enumerate(hdr_mat):
            cell = tbl_p3.rows[0].cells[i]
            cell.text = t
            cell.width = w_p[i]
            set_cell_margins(cell, 100, 100, 120, 120)
            set_cell_shading(cell, "F0F0F0")
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    format_run(r, size_pt=9.5, bold=True)

        rows_p3 = [
            ("Material bibliográfico especializado", "Unidad", "02", "50.00", "100.00"),
            ("Papel Bond A4 80 gr.", "Millar", "02", "25.00", "50.00"),
            ("Memoria USB 32 GB", "Unidad", "02", "40.00", "80.00"),
            ("Carpetas y fólderes", "Unidad", "10", "1.00", "10.00"),
            ("Lapiceros y resaltadores", "Unidad", "10", "2.00", "20.00"),
            ("Mascarillas KN95 y bioseguridad", "Caja", "02", "20.00", "40.00"),
            ("Alcohol en gel / líquido 70%", "Litro", "02", "10.00", "20.00"),
            ("Imprevistos / Otros", "Global", "01", "80.00", "80.00"),
            ("SUB TOTAL", "", "", "", "400.00")
        ]
        for row_d in rows_p3:
            row = tbl_p3.add_row()
            for idx, text in enumerate(row_d):
                c = row.cells[idx]
                c.text = text
                c.width = w_p[idx]
                set_cell_margins(c, 80, 80, 120, 120)
                if row_d[0] == "SUB TOTAL":
                    set_cell_shading(c, "F5F5F5")
                for p in c.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if idx >= 3 else (WD_ALIGN_PARAGRAPH.CENTER if idx in [1,2] else WD_ALIGN_PARAGRAPH.LEFT)
                    for r in p.runs:
                        format_run(r, size_pt=9.5, bold=(row_d[0] == "SUB TOTAL"))

        # TOTAL
        p_tot = doc.add_paragraph()
        p_tot.paragraph_format.space_before = Pt(8)
        p_tot.paragraph_format.space_after = Pt(6)
        r = p_tot.add_run("COSTO TOTAL DEL PROYECTO: S/. 4,600.00 Soles")
        format_run(r, size_pt=12, bold=True)

        add_heading_2(doc, "5.3. Financiamiento")
        add_body_p(doc, "El proyecto de investigación será autofinanciado de manera integral por la investigadora principal.")

        add_heading_2(doc, "5.4. Cronograma de actividades")
        add_body_p(doc, "El dimensionamiento temporal del proyecto se sustenta en la técnica de evaluación y revisión de programas (PERT), calculándose el tiempo estimado (te) mediante la formulación: te = (O + 4m + P) / 6, donde O representa el tiempo optimista, m el tiempo más probable y P el tiempo pesimista.")
        add_body_p(doc, "Asimismo, tal como enfatizó el docente en la cátedra de EN 486, la planificación administrativa de una tesis de pregrado debe ponderar con realismo la brecha existente entre los plazos normativos formales y la realidad académica universitaria: la formulación rigurosa del proyecto demanda aproximadamente cuatro meses tras culminar el ciclo académico (Serie 47); los trámites de dictamen y aprobación ante el Comité Institucional de Ética en Investigación (CIEI) y el Consejo de Facultad suelen tomar entre seis a ocho meses en la práctica; y la fase de ejecución de campo exige un periodo indispensable de al menos tres meses para garantizar inmersión reflexiva, construcción de confianza interpersonal con los informantes y saturación teórica efectiva. Con base en este análisis crítico, se establece un horizonte operativo sistemático de seis (06) meses para la fase de ejecución y dictamen final, programado de agosto de 2026 a enero de 2027.")

        # Tabla Cronograma
        tbl_cron = doc.add_table(rows=1, cols=7)
        tbl_cron.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_table_borders(tbl_cron, color="808080", sz="4", val="single")
        hdr_cron = ["Actividades", "Ago 2026", "Set 2026", "Oct 2026", "Nov 2026", "Dic 2026", "Ene 2027"]
        w_cron = [Cm(6.0), Cm(1.6), Cm(1.6), Cm(1.6), Cm(1.6), Cm(1.6), Cm(1.6)]
        for i, t in enumerate(hdr_cron):
            cell = tbl_cron.rows[0].cells[i]
            cell.text = t
            cell.width = w_cron[i]
            set_cell_margins(cell, 100, 100, 100, 100)
            set_cell_shading(cell, "F0F0F0")
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    format_run(r, size_pt=9.5, bold=True)

        cron_data = [
            ("Presentación del proyecto", "X", "", "", "", "", ""),
            ("Aprobación del proyecto", "", "X", "", "", "", ""),
            ("Validación de instrumentos", "", "", "X", "", "", ""),
            ("Revisión del marco teórico", "", "", "X", "X", "", ""),
            ("Coordinación con el establecimiento de salud", "", "", "", "X", "", ""),
            ("Recolección de datos (entrevistas)", "", "", "", "X", "X", ""),
            ("Transcripción y procesamiento cualitativo", "", "", "", "", "X", ""),
            ("Análisis cualitativo e interpretación", "", "", "", "", "X", ""),
            ("Presentación del informe final", "", "", "", "", "", "X"),
            ("Sustentación de la tesis", "", "", "", "", "", "X")
        ]
        for row_d in cron_data:
            row = tbl_cron.add_row()
            for idx, text in enumerate(row_d):
                c = row.cells[idx]
                c.text = text
                c.width = w_cron[idx]
                set_cell_margins(c, 80, 80, 100, 100)
                for p in c.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if idx > 0 else WD_ALIGN_PARAGRAPH.LEFT
                    for r in p.runs:
                        format_run(r, size_pt=9.5, bold=(text == "X"))


    # ==========================================================================
    # SECCIÓN 9: REFERENCIAS BIBLIOGRÁFICAS (Normas Vancouver)
    # ==========================================================================
    sec_refs = doc.add_section(WD_SECTION.NEW_PAGE)
    sec_refs.page_width = Cm(21.0)
    sec_refs.page_height = Cm(29.7)
    sec_refs.top_margin = Cm(3.0)
    sec_refs.left_margin = Cm(3.5)
    sec_refs.right_margin = Cm(2.5)
    sec_refs.bottom_margin = Cm(2.5)
    sec_refs.different_first_page_header_footer = False
    sec_refs.header.is_linked_to_previous = False
    sec_refs.footer.is_linked_to_previous = False
    sec_refs.footer.paragraphs[0].text = ""
    set_section_pgnum(sec_refs, start=None, fmt='decimal')
    setup_header(sec_refs, "REFERENCIAS BIBLIOGRÁFICAS")
    
    p_ref_tit = doc.add_paragraph()
    p_ref_tit.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_ref_tit.paragraph_format.space_before = Pt(0)
    p_ref_tit.paragraph_format.space_after = Pt(12)
    r = p_ref_tit.add_run("REFERENCIAS BIBLIOGRÁFICAS")
    format_run(r, size_pt=14, bold=True)
    
    referencias = [
        (1, "Dirección Regional de Salud de Ayacucho. Prioridades Regionales de Investigación en Salud de Ayacucho 2025–2030. Ayacucho: DIRESA Ayacucho; 2025. Disponible en: ", "https://repositorio.ins.gob.pe/", "https://repositorio.ins.gob.pe/"),
        (2, "Ministerio de Salud del Perú. Sistema de Gestión de la Calidad en Salud: Documento Técnico. Resolución Ministerial N° 519-2006/MINSA. Lima: MINSA; 2006. Disponible en: ", "https://www.gob.pe/institucion/minsa/normas-legales/255018-519-2006-minsa", "https://www.gob.pe/institucion/minsa/normas-legales/255018-519-2006-minsa"),
        (3, "Ministerio de Salud del Perú. Líneas Nacionales de Investigación en Salud al 2030. Resolución Ministerial N° 424-2025/MINSA. Lima: MINSA / Instituto Nacional de Salud; 2025. Disponible en: ", "https://www.gob.pe/institucion/minsa/normas-legales", "https://www.gob.pe/institucion/minsa/normas-legales"),
        (4, "World Health Organization, OECD, The World Bank. Delivering quality health services: a global imperative for universal health coverage. Geneva: World Health Organization; 2018. Disponible en: ", "https://iris.who.int/handle/10665/272465", "https://iris.who.int/handle/10665/272465"),
        (5, "Hernández-Sampieri R, Mendoza CP. Metodología de la investigación: las rutas cuantitativa, cualitativa y mixta. Ciudad de México: McGraw-Hill Interamericana Editores; 2018. Disponible en: ", "https://isbnsearch.org/isbn/9781456260965", "https://isbnsearch.org/isbn/9781456260965"),
        (6, "Creswell JW. Qualitative inquiry and research design: Choosing among five approaches. 3rd ed. Thousand Oaks (CA): SAGE Publications; 2013. Disponible en: ", "https://isbnsearch.org/isbn/9781412995306", "https://isbnsearch.org/isbn/9781412995306"),
        (7, "Husserl E. Ideas relativas a una fenomenología pura y una filosofía fenomenológica. Gaos J, traductor. México D.F.: Fondo de Cultura Económica; 1949. Disponible en: ", "https://dialnet.unirioja.es/servlet/libro?codigo=224213", "https://dialnet.unirioja.es/servlet/libro?codigo=224213"),
        (8, "van Manen M. Researching lived experience: Human science for an action sensitive pedagogy. Albany (NY): State University of New York Press / Routledge; 1990. doi: ", "https://doi.org/10.4324/9781315421056", "https://doi.org/10.4324/9781315421056"),
        (9, "Donabedian A. Explorations in Quality Assessment and Monitoring. Vol. 1: The Definition of Quality and Approaches to its Assessment. Ann Arbor (MI): Health Administration Press; 1980. Disponible en: ", "https://pubmed.ncbi.nlm.nih.gov/7409848/", "https://pubmed.ncbi.nlm.nih.gov/7409848/"),
        (10, "Donabedian A. Evaluating the quality of medical care. 1966. Milbank Q. 2005;83(4):691-729. doi: https://doi.org/10.1111/j.1468-0009.2005.00397.x. Disponible en: ", "https://pmc.ncbi.nlm.nih.gov/articles/PMC1464018/", "https://pmc.ncbi.nlm.nih.gov/articles/PMC1464018/"),
        (11, "Watson J. Nursing: The philosophy and science of caring. Revised ed. Boulder (CO): University Press of Colorado; 2008. Disponible en: ", "https://archive.org/details/nursingphilosoph0000wats", "https://archive.org/details/nursingphilosoph0000wats"),
        (12, "Lincoln YS, Guba EG. Naturalistic Inquiry. Beverly Hills (CA): SAGE Publications; 1985. Disponible en: ", "https://www.sagepub.com/shop/buy-a-book/naturalistic-inquiry-1-842", "https://www.sagepub.com/shop/buy-a-book/naturalistic-inquiry-1-842"),
        (13, "Okuda Benavides M, Gómez-Restrepo C. Métodos en investigación cualitativa: triangulación. Rev Colomb Psiquiatr. 2005;34(1):118-24. Disponible en: ", "http://www.scielo.org.co/scielo.php?script=sci_arttext&pid=S0034-74502005000100008", "http://www.scielo.org.co/scielo.php?script=sci_arttext&pid=S0034-74502005000100008"),
        (14, "Bunge M. La ciencia, su método y su filosofía. Buenos Aires: Editorial Sudamericana / Penguin Random House; 2014. Disponible en: ", "https://books.google.com.pe/books/about/La_ciencia_su_m%C3%A9todo_y_su_filosof%C3%ADa.html?id=lq8rDwAAQBAJ", "https://books.google.com.pe/books/about/La_ciencia_su_m%C3%A9todo_y_su_filosof%C3%ADa.html?id=lq8rDwAAQBAJ&redir_esc=y"),
        (15, "Asociación Médica Mundial. Declaración de Helsinki de la AMM: Principios éticos para las investigaciones médicas en seres humanos. Fortaleza (Brasil): AMM; 2013. Disponible en: ", "https://www.wma.net/es/policies-post/declaracion-de-helsinki-de-la-amm-principios-eticos-para-las-investigaciones-medicas-en-seres-humanos/", "https://www.wma.net/es/policies-post/declaracion-de-helsinki-de-la-amm-principios-eticos-para-las-investigaciones-medicas-en-seres-humanos/"),
        (16, "Consejo de Organizaciones Internacionales de las Ciencias Médicas (CIOMS). Pautas éticas internacionales para la investigación relacionada con la salud con seres humanos. 4ta ed. Ginebra: CIOMS / OPS; 2016. Disponible en: ", "https://cioms.ch/publications/product/pautas-eticas-internacionales-para-la-investigacion-relacionada-con-la-salud-con-seres-humanos/", "https://cioms.ch/publications/product/pautas-eticas-internacionales-para-la-investigacion-relacionada-con-la-salud-con-seres-humanos/"),
        (17, "Aguirre-Andrade M, Tenorio-Acosta I, Rivas-Díaz L, Valenzuela-Oré F, Sánchez-Simbrón M, Quino-Huamaní F. Promoción de salud y nivel de salubridad de las familias, en comunidades del Distrito de Cangallo, Ayacucho 2020. Rev Investig Salud. 2020;14(2):45-58. Disponible en: ", "http://repositorio.unsch.edu.pe/", "http://repositorio.unsch.edu.pe/")
    ]
    
    for num, texto_base, texto_url, url_link in referencias:
        p_ref = doc.add_paragraph()
        p_ref.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_ref.paragraph_format.line_spacing = 1.3
        p_ref.paragraph_format.left_indent = Cm(1.27)
        p_ref.paragraph_format.first_line_indent = Cm(-1.27)
        p_ref.paragraph_format.space_before = Pt(0)
        p_ref.paragraph_format.space_after = Pt(4)
        
        r_num = p_ref.add_run(f"[{num}]  ")
        format_run(r_num, size_pt=12, bold=False)
        r_txt = p_ref.add_run(texto_base)
        format_run(r_txt, size_pt=12, bold=False)
        add_hyperlink(p_ref, url_link, texto_url, color="0000EE", underline=True, font_size=12)

    # ==========================================================================
    # SECCIÓN 10: ANEXOS
    # ==========================================================================
    sec_anexos = doc.add_section(WD_SECTION.NEW_PAGE)
    sec_anexos.page_width = Cm(21.0)
    sec_anexos.page_height = Cm(29.7)
    sec_anexos.top_margin = Cm(3.0)
    sec_anexos.left_margin = Cm(3.5)
    sec_anexos.right_margin = Cm(2.5)
    sec_anexos.bottom_margin = Cm(2.5)
    sec_anexos.different_first_page_header_footer = False
    sec_anexos.header.is_linked_to_previous = False
    sec_anexos.footer.is_linked_to_previous = False
    sec_anexos.footer.paragraphs[0].text = ""
    set_section_pgnum(sec_anexos, start=None, fmt='decimal')
    setup_header(sec_anexos, "ANEXOS")
    
    p_anx_tit = doc.add_paragraph()
    p_anx_tit.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_anx_tit.paragraph_format.space_before = Pt(0)
    p_anx_tit.paragraph_format.space_after = Pt(12)
    r = p_anx_tit.add_run("ANEXOS")
    format_run(r, size_pt=14, bold=True)
    
    # --------------------------------------------------------------------------
    # ANEXO 1
    # --------------------------------------------------------------------------
    add_heading_2(doc, "ANEXO 1: GUÍA DE ENTREVISTA A PROFUNDIDAD SEMIESTRUCTURADA")
    add_body_p(doc, "Calidad de atención en usuarios del Centro de Salud Belén, Ayacucho 2026.", bold_prefix="Título: ", indent_cm=0)
    add_body_p(doc, "Comprender la vivencia intersubjetiva y el significado que los usuarios otorgan a la calidad del cuidado de salud recibido en el Centro de Salud Belén.", bold_prefix="Objetivo: ", indent_cm=0)
    
    add_heading_3(doc, "I. Preguntas de Apertura y Clima de Confianza (Rapport y Napaykuy):")
    add_num_item(doc, 1, "¿Podría relatarme qué motivo de salud le motivó a acudir hoy al Centro de Salud Belén y cómo se sentía anímicamente antes de ingresar al consultorio?\n(En quechua chanka: ¿Imataq qamta sasachasunki mamay/taytay, imaynam sonqoyki kachkarqa doctorman manaraq yaykuchkaspayki?)")
    add_num_item(doc, 2, "¿Es la primera vez que se atiende en este establecimiento o suele venir frecuentemente a este servicio?\n(En quechua chanka: ¿Qallariykutichu kay Centro de Saludpi atendichikunki, icha sapa kutichu hamunki?)")
    
    add_heading_3(doc, "II. Preguntas por Dimensiones Existenciales Fenomenológicas (van Manen y Sampieri):")
    add_body_p(doc, "Dimensión Existencial 1: Temporalidad (El tiempo vivido en la consulta)", bold_prefix="• ")
    add_num_item(doc, 3, "¿Cómo experimentó usted el tiempo que el personal de salud (médico, enfermera) le dedicó en la consulta? ¿Sintió que la atención fue apresurada y con prisa, o que le brindaron el tiempo y la calma necesaria para escuchar sus dolencias y explicarle con claridad su tratamiento?\n(En quechua chanka: ¿Imaynatam sientekurqanki doctor o enfermera atiendesunaykipaq tiempo qosusqaykiwan? ¿Utqayllachu qawasurqanki icha pacienciawanchu llapa tapukuyniykikunata uyarirqasunki hinaspa hampiytapas allinta willasurqanki?)")
    
    add_body_p(doc, "Dimensión Existencial 2: Espacialidad (El espacio vivido y privacidad en consultorio)", bold_prefix="• ")
    add_num_item(doc, 4, "¿Cómo experimentó el ambiente del consultorio durante su atención? ¿Sintió que el espacio le brindó la debida privacidad, tranquilidad y reserva para conversar sobre sus problemas de salud y ser examinado(a) sin que personas ajenas observen o escuchen?\n(En quechua chanka: ¿Imaynatam consultorio ukuta sienterqanki? ¿Chaypi hawka, chuya hinaspa secretollapichu sientekurqanki doctorwan onqoyniykimanta rimanaykipaq, mana pipas uyaripayasuspayki?)")
    
    add_body_p(doc, "Dimensión Existencial 3: Corporalidad (El cuerpo vivido durante el acto de cuidado)", bold_prefix="• ")
    add_num_item(doc, 5, "Al momento de ser examinado(a) físicamente o recibir una atención (curación, inyección, control de CRED o triaje), ¿cómo sintió que el personal de salud trató su cuerpo? ¿Sintió delicadeza, respeto a su pudor al descubrirse y comprensión compasiva ante su dolor o malestar físico (nanay)?\n(En quechua chanka: ¿Doctorkuna o enfermerakuna cuerpoykita examinaptinku o hampiptinku, imaynam sientekurqanki? ¿Llampu makillawanchu, pinqakuyniykita respetaspachu hinaspa nanayniykita llakipayaspachu cuidarasunki?)")
    
    add_body_p(doc, "Dimensión Existencial 4: Relacionalidad (La relación humana vivida e intersubjetividad)", bold_prefix="• ")
    add_num_item(doc, 6, "¿Cómo fue el trato interpersonal y el encuentro humano con el personal de salud (médico, enfermera)? ¿Sintió un trato digno, acogedor y empático con contacto visual? ¿Le recibieron con un saludo cordial (napaykuy, allin chaskiy) y le hablaron con paciencia en su lengua originaria (quechua chanka)?\n(En quechua chanka: ¿Imaynam runa kayninkuwan chaskisurqanki doctor o enfermera? ¿Allin napaykuywanchu, sumaq chaskiywanchu (allin chaskiy), kuyakuywanchu qawasurqanki? ¿Runasimipichu respetowan rimapayasurqanki?)")
    
    add_body_p(doc, "Dimensión 5: Síntesis de la Esencia de la Calidad y Cuidado Humano", bold_prefix="• ")
    add_num_item(doc, 7, "Al reflexionar sobre la atención recibida hoy, ¿qué significa para usted que un servicio de salud brinde auténtica calidad y calidez humana, y qué aspectos considera fundamentales para que el cuidado en el Centro de Salud Belén sea verdaderamente digno?\n(En quechua chanka: ¿Tukuy kaykunata yuyarispa, imataq qampaq allin calidadwan kuyapayaywan hampikuyqa? ¿Imatataq doctorwan enfermera ruranman llapa runa sonqonpi allin atendisqa hinaspa respetasqa sientekunanpaq?)")
    
    # --------------------------------------------------------------------------
    # ANEXO 2: MATRIZ DE CONSISTENCIA CUALITATIVA
    # --------------------------------------------------------------------------
    add_heading_2(doc, "ANEXO 2: MATRIZ DE CONSISTENCIA FENOMENOLÓGICA CUALITATIVA")
    add_body_p(doc, "CALIDAD DE ATENCIÓN EN USUARIOS DEL CENTRO DE SALUD BELÉN, AYACUCHO 2026.", bold_prefix="TÍTULO: ", indent_cm=0)
    
    tbl_mat = doc.add_table(rows=1, cols=5)
    tbl_mat.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_table_borders(tbl_mat, color="808080", sz="4", val="single")
    hdr_mat = ["PROBLEMA FENOMENOLÓGICO", "OBJETIVOS CUALITATIVOS", "PREGUNTAS NORTEADORAS", "DIMENSIONES EXISTENCIALES", "METODOLOGÍA CUALITATIVA"]
    w_mat = [Cm(3.0), Cm(3.0), Cm(3.2), Cm(3.0), Cm(3.4)]
    for i, t in enumerate(hdr_mat):
        cell = tbl_mat.rows[0].cells[i]
        cell.text = t
        cell.width = w_mat[i]
        set_cell_margins(cell, 100, 100, 110, 110)
        set_cell_shading(cell, "E8EEF5")
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                format_run(r, size_pt=9.5, bold=True)
                
    row_mat = tbl_mat.add_row()
    mat_cols_content = [
        ("General:\n¿Cuál es el significado, la estructura y la esencia de la experiencia vivida por los usuarios respecto a la calidad de atención en el Centro de Salud Belén, Ayacucho 2026?\n\n"
         "Específicos:\n"
         "1. Analizar la calidez relacional y el diálogo en quechua chanka.\n"
         "2. Comprender la oportunidad y tiempo efectivo dedicado en consulta.\n"
         "3. Describir la privacidad y confidencialidad en consultorio.\n"
         "4. Identificar el respeto al pudor físico y alivio del dolor.\n"
         "5. Sintetizar la esencia y divergencias de la calidad asistencial."),
        ("General:\nComprender el significado, estructura y esencia de la experiencia vivida por los usuarios respecto a la calidad de atención en el Centro de Salud Belén, Ayacucho 2026.\n\n"
         "Específicos:\n"
         "1. Comprender la dimensión relacional y comunicación en quechua chanka.\n"
         "2. Explorar la dimensión temporal en la duración de la consulta.\n"
         "3. Describir la dimensión espacial y confidencialidad en consultorio.\n"
         "4. Identificar la dimensión corporal, pudor físico y dolor (nanay).\n"
         "5. Sintetizar la esencia compartida y visiones divergentes."),
        ("1. ¿Cómo experimentan el trato digno, empatía y quechua chanka?\n\n"
         "2. ¿De qué manera vivencian la duración y calma en la consulta?\n\n"
         "3. ¿Cómo vivencian la reserva y privacidad en el consultorio?\n\n"
         "4. ¿Cómo vivencian el respeto a su pudor somático y dolor?\n\n"
         "5. ¿Cuál es la esencia y visiones divergentes de la calidad?"),
        ("Categoría Central:\nCalidad de Atención de Salud (Donabedian y Watson)\n\n"
         "Dimensiones Existenciales (Lebenswelt):\n"
         "• D1: Relacionalidad (Mitwelt)\n"
         "• D2: Temporalidad (Lebenszeit)\n"
         "• D3: Espacialidad (Lebensraum)\n"
         "• D4: Corporalidad (Leib)\n"
         "• D5: Esencia y Divergencias"),
        ("Enfoque y Diseño:\nInterpretativo-fenomenológico, naturalista (Husserl, van Manen, Sampieri).\n\n"
         "Muestra e Informantes:\nCasos tipo no probabilístico. Criterio de saturación teórica (12 a 18 informantes clave).\n\n"
         "Técnicas:\nEntrevista fenomenológica en profundidad y observación reflexiva.\n\n"
         "Rigor Científico:\nCriterios de Guba y Lincoln (1985): credibilidad, transferibilidad, consistencia y confirmabilidad.\n\n"
         "Software:\nATLAS.ti v9.")
    ]
    for idx, text in enumerate(mat_cols_content):
        c = row_mat.cells[idx]
        c.text = text
        c.width = w_mat[idx]
        set_cell_margins(c, 100, 100, 110, 110)
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            for r in p.runs:
                format_run(r, size_pt=9.0, bold=False)

    # --------------------------------------------------------------------------
    # ANEXO 3: FORMATO DE JUICIO DE EXPERTOS
    # --------------------------------------------------------------------------
    add_heading_2(doc, "ANEXO 3: FORMATO DE JUICIO DE EXPERTOS")
    add_heading_3(doc, "I. ASPECTOS GENERALES")
    add_body_p(doc, "__________________________________________________________________", bold_prefix="1.1. Apellidos y nombres del informante (Experto): ", indent_cm=0)
    add_body_p(doc, "__________________________________________________________________", bold_prefix="1.2. Grado académico del experto: ", indent_cm=0)
    add_body_p(doc, "__________________________________________________________________", bold_prefix="1.3. Profesión del experto: ", indent_cm=0)
    add_body_p(doc, "__________________________________________________________________", bold_prefix="1.4. Institución donde labora: ", indent_cm=0)
    add_body_p(doc, "__________________________________________________________________", bold_prefix="1.5. Cargo que desempeña: ", indent_cm=0)
    add_body_p(doc, "Guía de entrevista a profundidad semiestructurada.", bold_prefix="1.6. Denominación del instrumento: ", indent_cm=0)
    add_body_p(doc, "Gresly Lucero Pariona Palomino.", bold_prefix="1.7. Autor del instrumento: ", indent_cm=0)
    add_body_p(doc, "Calidad de atención en usuarios del Centro de Salud Belén, Ayacucho 2026.", bold_prefix="1.8. Título de la tesis: ", indent_cm=0)
    
    add_heading_3(doc, "II. CRITERIOS DE VALIDACIÓN")
    tbl_exp = doc.add_table(rows=1, cols=5)
    tbl_exp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_table_borders(tbl_exp, color="808080", sz="4", val="single")
    hdr_exp = ["CRITERIOS", "INDICADORES DE EVALUACIÓN", "SÍ", "NO", "OBSERVACIONES"]
    w_exp = [Cm(3.2), Cm(6.5), Cm(1.2), Cm(1.2), Cm(3.5)]
    for i, t in enumerate(hdr_exp):
        cell = tbl_exp.rows[0].cells[i]
        cell.text = t
        cell.width = w_exp[i]
        set_cell_margins(cell, 100, 100, 110, 110)
        set_cell_shading(cell, "F0F0F0")
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                format_run(r, size_pt=9.5, bold=True)
                
    exp_data = [
        ("1. CLARIDAD DIALÓGICA", "Preguntas formuladas con lenguaje coloquial, abierto y comprensible para el usuario andino.", "", "", ""),
        ("2. PERTINENCIA LINGÜÍSTICA", "Preguntas en quechua chanka gramaticalmente naturales, afectuosas y culturalmente pertinentes.", "", "", ""),
        ("3. CONSISTENCIA FENOMENOLÓGICA", "Articulación rigurosa con los cuatro existenciales del mundo de la vida (van Manen y Sampieri).", "", "", ""),
        ("4. COHERENCIA VIVENCIAL", "Correspondencia simétrica entre las preguntas y la captura de la esencia y las divergencias vivenciales.", "", "", ""),
        ("5. RELEVANCIA DEL CUIDADO", "Explora estrictamente la interacción humana, los sentires corporales y la dignidad del cuidado.", "", "", ""),
        ("6. CAPACIDAD DE SATURACIÓN", "Capacidad inductora de las preguntas para suscitar relatos densos y alcanzar la saturación teórica.", "", "", "")
    ]
    for row_d in exp_data:
        row = tbl_exp.add_row()
        for idx, text in enumerate(row_d):
            c = row.cells[idx]
            c.text = text
            c.width = w_exp[idx]
            set_cell_margins(c, 80, 80, 110, 110)
            for p in c.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if idx in [2, 3] else WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    format_run(r, size_pt=9.5, bold=(idx == 0))
                    
    add_body_p(doc, "____________________________________________________________________________________________", bold_prefix="Observaciones generales: ", indent_cm=0)
    add_body_p(doc, "Ayacucho, _____ de ______________________ del 2026.", indent_cm=0)
    
    p_sig1 = doc.add_paragraph()
    p_sig1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sig1.paragraph_format.space_before = Pt(30)
    p_sig1.paragraph_format.space_after = Pt(2)
    r = p_sig1.add_run("____________________________________________________")
    format_run(r, size_pt=11, bold=False)
    
    p_sig2 = doc.add_paragraph()
    p_sig2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sig2.paragraph_format.space_before = Pt(0)
    p_sig2.paragraph_format.space_after = Pt(16)
    r = p_sig2.add_run("FIRMA Y SELLO DEL EXPERTO")
    format_run(r, size_pt=11, bold=True)
    
    # --------------------------------------------------------------------------
    # ANEXO 4: MODELO DE CONSENTIMIENTO INFORMADO
    # --------------------------------------------------------------------------
    add_heading_2(doc, "ANEXO 4: MODELO DE CONSENTIMIENTO INFORMADO")
    add_body_p(doc, "Yo, __________________________________________________________________, identificado con DNI N° _____________________, domiciliado en ________________________________________________________, Distrito de Ayacucho, Región Ayacucho.", indent_cm=0)
    add_body_p(doc, "He tomado conocimiento del estudio de investigación titulado: “CALIDAD DE ATENCIÓN EN USUARIOS DEL CENTRO DE SALUD BELÉN, AYACUCHO 2026”.", indent_cm=0)
    add_body_p(doc, "Declaro participar en calidad de Informante clave y me comprometo a brindar información fidedigna. Se me ha informado que la entrevista será grabada únicamente en audio, que mi participación es totalmente voluntaria, anónima y confidencial, y que puedo retirarme del estudio en el momento que considere conveniente sin que esto afecte mi atención en el centro de salud.", indent_cm=0)
    add_body_p(doc, "Para dar conformidad a este acto, firmo e imprimo mi huella digital al pie del documento.", indent_cm=0)
    add_body_p(doc, "Lugar y fecha: Ayacucho, _____ de __________________________ de 2026.", indent_cm=0)
    
    p_h1 = doc.add_paragraph()
    p_h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_h1.paragraph_format.space_before = Pt(35)
    p_h1.paragraph_format.space_after = Pt(2)
    r = p_h1.add_run("____________________________________________________                 [                           ]")
    format_run(r, size_pt=11, bold=False)
    
    p_h2 = doc.add_paragraph()
    p_h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_h2.paragraph_format.space_before = Pt(0)
    p_h2.paragraph_format.space_after = Pt(16)
    r = p_h2.add_run("FIRMA DEL PARTICIPANTE                                         HUELLA DIGITAL")
    format_run(r, size_pt=10.5, bold=True)
    
    # --------------------------------------------------------------------------
    # ANEXO 5: MODELO DE CARTA DE ASESORÍA
    # --------------------------------------------------------------------------
    add_heading_2(doc, "ANEXO 5: MODELO DE CARTA DE ASESORÍA")
    add_body_p(doc, "Ayacucho, _____ de enero del 2026.", indent_cm=0)
    add_body_p(doc, "Dr. Alejandro Yarlequé Mujica\nDecano de la Facultad de Ciencias de la Salud\nUniversidad Nacional de San Cristóbal de Huamanga", indent_cm=0)
    add_body_p(doc, "Asunto: Carta de Aceptación y Asesoría Formal de Proyecto de Tesis", bold_prefix="ASUNTO: ", indent_cm=0)
    add_body_p(doc, "Sirva la presente para saludarlo cordialmente y a la vez comunicarle la asesoría formal del Proyecto de Tesis cualitativo titulado: “CALIDAD DE ATENCIÓN EN USUARIOS DEL CENTRO DE SALUD BELÉN, AYACUCHO 2026”, perteneciente a la estudiante Gresly Lucero Pariona Palomino, egresada de la Escuela Profesional de Enfermería.", indent_cm=0)
    add_body_p(doc, "En tal sentido, dicha asesoría comprenderá todas las etapas del proyecto y ejecución del trabajo de investigación, permitiendo la obtención del Título Profesional de Licenciada en Enfermería.", indent_cm=0)
    add_body_p(doc, "Atentamente,", indent_cm=0)
    
    p_dr1 = doc.add_paragraph()
    p_dr1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_dr1.paragraph_format.space_before = Pt(35)
    p_dr1.paragraph_format.space_after = Pt(2)
    r = p_dr1.add_run("____________________________________________________")
    format_run(r, size_pt=11, bold=False)
    
    p_dr2 = doc.add_paragraph()
    p_dr2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_dr2.paragraph_format.space_before = Pt(0)
    p_dr2.paragraph_format.space_after = Pt(2)
    r = p_dr2.add_run("Dr. Manglio Aguirre Andrade")
    format_run(r, size_pt=12, bold=True)
    
    p_dr3 = doc.add_paragraph()
    p_dr3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_dr3.paragraph_format.space_before = Pt(0)
    p_dr3.paragraph_format.space_after = Pt(0)
    r = p_dr3.add_run("Docente de la Facultad de Ciencias de la Salud - UNSCH\nORCID: 0000-0001-8234-567X | DNI N°: ___________________")
    format_run(r, size_pt=10, bold=False)

    # Guardar en la ruta objetivo
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    print(f"Documento Word generado exitosamente en: {output_path}")

if __name__ == "__main__":
    out_file = r"c:\GRESLY\DOCUMENTOS\Generación\protocolo_cualitativo.docx"
    build_protocolo_word(out_file)
    
    aqui_copy = r"c:\GRESLY\plantilla\AQUI\protocolo_cualitativo.docx"
    build_protocolo_word(aqui_copy)

    edita_copy = r"c:\GRESLY\plantilla\AQUI\EDITA.docx"
    build_protocolo_word(edita_copy)
