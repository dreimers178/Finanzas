"""Utilidades para generar documentos .docx con formato APA 7 (python-docx >= 1.2)."""
import re
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_COLOR_INDEX, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor

FONT = "Times New Roman"
AUTHOR = "Revisión metodológica"
INITIALS = "RM"

TOKEN = re.compile(r"(\*\*.+?\*\*|\*.+?\*|==.+?==)")


def new_document():
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(27.94), Cm(21.59)  # carta
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Cm(2.54))
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    pf = st.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    # Encabezados APA 7
    specs = {
        1: dict(align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, italic=False, indent=0),
        2: dict(align=WD_ALIGN_PARAGRAPH.LEFT, bold=True, italic=False, indent=0),
        3: dict(align=WD_ALIGN_PARAGRAPH.LEFT, bold=True, italic=True, indent=0),
        4: dict(align=WD_ALIGN_PARAGRAPH.LEFT, bold=True, italic=False, indent=1.27),
    }
    for lvl, s in specs.items():
        hs = doc.styles[f"Heading {lvl}"]
        hs.font.name = FONT
        hs.font.size = Pt(12)
        hs.font.bold = s["bold"]
        hs.font.italic = s["italic"]
        hs.font.color.rgb = RGBColor(0, 0, 0)
        rpr = hs.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.append(rf)
        for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rf.set(qn(a), FONT)
        for th in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            if rf.get(qn(th)) is not None:
                del rf.attrib[qn(th)]
        hpf = hs.paragraph_format
        hpf.alignment = s["align"]
        hpf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
        hpf.space_before = Pt(0)
        hpf.space_after = Pt(0)
        hpf.first_line_indent = Cm(s["indent"])
        hpf.keep_with_next = True
    add_page_number(sec)
    return doc


def add_page_number(section):
    p = section.header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _field(p.add_run(), "PAGE")


def _field(run, code):
    for tag, text in (("begin", None), ("instr", code), ("separate", None), ("end", None)):
        if tag == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = code
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tag)
        run._r.append(el)


def add_toc(doc):
    p = doc.add_paragraph()
    _field(p.add_run(), 'TOC \\o "1-3" \\h \\z \\u')
    q = doc.add_paragraph()
    r = q.add_run("(Clic derecho → «Actualizar campo» para generar la tabla de contenido con la paginación final.)")
    r.italic = True
    r.font.size = Pt(10)


def runs(p, text, size=None):
    """Agrega texto con marcas **negrita**, *cursiva* y ==resaltado (verificar)==."""
    out = []
    for part in TOKEN.split(text):
        if not part:
            continue
        if part.startswith("**"):
            r = p.add_run(part[2:-2]); r.bold = True
        elif part.startswith("=="):
            r = p.add_run(part[2:-2]); r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        elif part.startswith("*"):
            r = p.add_run(part[1:-1]); r.italic = True
        else:
            r = p.add_run(part)
        if size:
            r.font.size = Pt(size)
        out.append(r)
    return out


def comment(doc, rs, text):
    if rs and text:
        doc.add_comment(rs, text=text, author=AUTHOR, initials=INITIALS)


def para(doc, text, c=None, indent=True, align=None, size=None, spacing=None):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Cm(1.27)
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "justify":
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if spacing == 1:
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    rs = runs(p, text, size)
    comment(doc, rs, c)
    return p


def heading(doc, text, level, c=None):
    h = doc.add_heading(level=level)
    rs = runs(h, text)
    comment(doc, rs, c)
    return h


def heading5(doc, text, c=None):
    """Nivel 4 APA en línea: sangría, negrita, termina en punto; el texto sigue en el mismo párrafo."""
    pass


def bullets(doc, items, c=None, numbered=False):
    style = "List Number" if numbered else "List Bullet"
    first = None
    for i, it in enumerate(items):
        p = doc.add_paragraph(style=style)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
        rs = runs(p, it)
        if i == 0:
            first = rs
    comment(doc, first, c)


def ref(doc, text, c=None):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(1.27)
    pf.first_line_indent = Cm(-1.27)
    rs = runs(p, text)
    comment(doc, rs, c)
    return p


def page_break(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def _borders(cell, top=None, bottom=None):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for edge, val in (("top", top), ("bottom", bottom), ("left", None), ("right", None)):
        el = OxmlElement(f"w:{edge}")
        if val:
            el.set(qn("w:val"), "single"); el.set(qn("w:sz"), str(val)); el.set(qn("w:color"), "000000")
        else:
            el.set(qn("w:val"), "nil")
        b.append(el)
    tcPr.append(b)


def table(doc, number, title, header, rows, note=None, widths=None, size=10, c=None):
    """Tabla con estilo APA 7: número en negrita, título en cursiva, solo líneas horizontales."""
    p = doc.add_paragraph(); p.paragraph_format.keep_with_next = True
    rs = runs(p, f"**Tabla {number}**"); comment(doc, rs, c)
    p = doc.add_paragraph(); p.paragraph_format.keep_with_next = True
    r = p.add_run(title); r.italic = True
    t = doc.add_table(rows=1 + len(rows), cols=len(header))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(header):
        cell = t.rows[0].cells[j]
        cell.text = ""
        cp = cell.paragraphs[0]
        cp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        runs(cp, f"**{h}**", size)
        _borders(cell, top=8, bottom=8)
    for i, row in enumerate(rows, start=1):
        last = i == len(rows)
        for j, val in enumerate(row):
            cell = t.rows[i].cells[j]
            cell.text = ""
            lines = val.split("\n") if isinstance(val, str) else [str(val)]
            for k, line in enumerate(lines):
                cp = cell.paragraphs[0] if k == 0 else cell.add_paragraph()
                cp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                runs(cp, line, size)
            _borders(cell, bottom=8 if last else None)
    if widths:
        t.autofit = False
        for j, col in enumerate(t.columns):
            col.width = Cm(widths[j])
        grid = t._tbl.tblGrid
        for j, gc in enumerate(grid.findall(qn("w:gridCol"))):
            gc.set(qn("w:w"), str(int(Cm(widths[j]).twips)))
        for row in t.rows:
            for j, w in enumerate(widths):
                row.cells[j].width = Cm(w)
    if note:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.paragraph_format.space_after = Pt(12)
        rs = runs(p, f"*Nota.* {note}", 10)
    else:
        doc.add_paragraph()
    return t
