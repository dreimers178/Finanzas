"""Convierte .docx a PDF con LibreOffice actualizando la tabla de contenido.

Uso: /usr/bin/python3 a_pdf.py entrada.docx salida.pdf [--comentarios]
"""
import subprocess, sys, time, pathlib
import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    p = PropertyValue(); p.Name = name; p.Value = value
    return p


src = pathlib.Path(sys.argv[1]).resolve()
dst = pathlib.Path(sys.argv[2]).resolve()
with_comments = "--comentarios" in sys.argv

proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore",
                         "--accept=socket,host=localhost,port=2002;urp;"])
try:
    ctx = None
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(60):
        try:
            ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
            break
        except Exception:
            time.sleep(0.5)
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.loadComponentFromURL(src.as_uri(), "_blank", 0, (prop("Hidden", True),))
    # Sustituye el campo TOC de python-docx (que LibreOffice no regenera) por un índice nativo.
    search = doc.createSearchDescriptor()
    search.SearchString = "Actualizar campo"
    found = doc.findFirst(search)
    if found is not None:
        text = found.getText()
        cur = text.createTextCursorByRange(found.getStart())
        cur.gotoStartOfParagraph(False)
        cur.gotoEndOfParagraph(True)
        cur.setString("")
        toc = doc.createInstance("com.sun.star.text.ContentIndex")
        toc.CreateFromOutline = True
        toc.Level = 3
        toc.Title = ""
        text.insertTextContent(cur, toc, False)
    for _ in range(2):  # dos pasadas: la TOC cambia la paginación
        idx = doc.getDocumentIndexes()
        for i in range(idx.getCount()):
            idx.getByIndex(i).update()
        doc.refresh()
    filt = uno.Any("[]com.sun.star.beans.PropertyValue", tuple([
        prop("ExportNotes", with_comments),
        prop("ExportNotesInMargin", with_comments),
        prop("UseTaggedPDF", True),
    ]))
    doc.storeToURL(dst.as_uri(), (prop("FilterName", "writer_pdf_Export"), prop("FilterData", filt)))
    doc.close(True)
    print("OK", dst)
finally:
    try:
        desktop.terminate()
    except Exception:
        pass
    proc.wait(timeout=30)
