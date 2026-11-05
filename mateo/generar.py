"""Hojas para imprimir del libro «Mi familia, mi lugar especial» de Mateo (tamaño carta)."""
import re, html

def trazo(texto):
    """Texto para repasar: contorno gris; mayúsculas y vocales con acento en rojo."""
    out = []
    for ch in texto:
        e = html.escape(ch)
        if ch.isupper() or ch in "áéíóúÁÉÍÓÚ":
            out.append(f'<span class="r">{e}</span>')
        else:
            out.append(e)
    return f'<span class="trace">{"".join(out)}</span>'

def hoja(titulo, instr, cuerpo, fondo):
    return f'''<section class="sheet"><div class="hdr"><b>{titulo}</b><span>{instr}</span></div>
<div class="sheet-body" style="--fondo:{fondo}">{cuerpo}</div></section>'''

CUT = '<div class="cut">✂ recortar por la línea punteada</div>'

hojas = []

# Portada
hojas.append(hoja("Portada", "Recorta el recuadro y pégalo en la cartulina. Agrega stickers alrededor.", f'''
<div class="cutbox cover">
  <div class="cv-title">Mi familia,<br><span>mi lugar especial</span></div>
  <img class="cv-logo" src="img/mateo_logo.jpg" alt="Mateo">
  <div class="cv-author">Autor: Mateo · 1.° Green</div>
</div>{CUT}''', "#e8f5e9"))

# Página 1
hojas.append(hoja("Página 1 · Así soy yo", "Fondo: hoja verde. Recorta la foto y la tira de texto; Mateo dibuja y repasa las letras.", f'''
<div class="row2">
  <div class="cutbox photo"><img src="img/estadio.jpg" alt="Mateo en el estadio"></div>
  <div class="cutbox drawbox"><div class="dlabel">Así me dibujo yo</div></div>
</div>
<div class="cutbox lines">
  <div class="tl">{trazo("Mi nombre es Mateo.")}</div>
  <div class="tl">{trazo("Tengo 6 años.")}</div>
  <div class="tl">{trazo("Me gusta el fútbol.")}</div>
</div>{CUT}''', "#e8f5e9"))

# Página 2
labels = [("Mamá", "Ana"), ("Papá", "José Luis"), ("Manito", "José Luis"), ("Yo", "Mateo")]
lab = "".join(f'<div class="cutbox tag"><span class="arrow">➜</span>{trazo(a)}<br><small>{trazo(b)}</small></div>' for a, b in labels)
hojas.append(hoja("Página 2 · Mi familia", "Fondo: hoja azul. Pega la foto y coloca cada etiqueta junto a la persona.", f'''
<div class="cutbox photo big"><img src="img/playa.jpg" alt="Familia en la playa"></div>
<div class="tags">{lab}</div>
<div class="cutbox lines one"><div class="tl">{trazo("Esta es mi familia.")}</div></div>{CUT}''', "#e3f2fd"))

# Página 3
acts = ["jugar juegos de mesa", "hacer la tarea con mis papás", "viajar"]
boxes = "".join(f'<div class="cutbox mini"><div class="mini-draw"></div><div class="mini-cap">{trazo(a)}</div></div>' for a in acts)
hojas.append(hoja("Página 3 · Mi familia hace", "Fondo: hoja roja. Mateo dibuja en el recuadro grande y en los tres pequeños.", f'''
<div class="cutbox lines"><div class="tl">{trazo("Me gusta jugar fútbol")}</div><div class="tl" style="border:0">{trazo("con mi familia.")}</div></div>
<div class="cutbox drawbox wide"><div class="dlabel">Mi dibujo: jugando fútbol con mi familia</div></div>
<div class="sub">También me gusta:</div>
<div class="row3">{boxes}</div>{CUT}''', "#ffebee"))

# Página 4
hojas.append(hoja("Página 4 · Lo que amo de mi familia", "Fondo: hoja verde o azul. Recorta el corazón; Mateo repasa las letras y lo colorea.", f'''
<div class="heart-wrap">
  <svg class="heart" viewBox="0 0 100 92" preserveAspectRatio="none"><path d="M50 90 C 18 66 2 50 2 28 C 2 12 14 2 28 2 C 38 2 46 8 50 16 C 54 8 62 2 72 2 C 86 2 98 12 98 28 C 98 50 82 66 50 90 Z" fill="#fff" stroke="#e53935" stroke-width="1.2" stroke-dasharray="2.5 1.8"/></svg>
  <div class="heart-text">
    <div class="tl">{trazo("Lo que más me gusta")}</div>
    <div class="tl">{trazo("de mi familia es que")}</div>
    <div class="tl">{trazo("es divertida, cariñosa")}</div>
    <div class="tl">{trazo("y le gusta el fútbol.")}</div>
  </div>
</div>
<div class="cutbox badge">🏆 {trazo("La mejor familia")}</div>{CUT}''', "#e3f2fd"))

# Stickers
hojas.append(hoja("Stickers de fútbol", "Recorta los que más te gusten para decorar la portada y las páginas.", '''
<div class="stickers"><img src="img/stickers.jpg" alt="Stickers"></div>''', "#fff"))

HTML = f'''<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Libro de Mateo</title><style>
@font-face{{font-family:"Nunito";src:url(fonts/Nunito.woff2) format("woff2");font-weight:200 1000}}
@font-face{{font-family:"Alfa";src:url(fonts/AlfaSlabOne.woff2) format("woff2")}}
@page{{size:letter;margin:0}}
*{{box-sizing:border-box}}
html,body{{margin:0;font-family:"Nunito",sans-serif;color:#222}}
.sheet{{width:21.59cm;height:27.94cm;padding:1cm 1.2cm;break-after:page;display:flex;flex-direction:column}}
.hdr{{border-bottom:2px solid #2e7d32;padding-bottom:.15cm;margin-bottom:.4cm;font-size:10.5pt;color:#555}}
.hdr b{{display:block;font-family:"Alfa";font-weight:400;font-size:17pt;color:#2e7d32}}
.sheet-body{{flex:1;display:flex;flex-direction:column;gap:.45cm;align-items:center}}
.cutbox{{border:2px dashed #999;border-radius:.3cm;padding:.35cm;background:#fff}}
.cut{{font-size:9pt;color:#999;margin-top:auto}}
.trace{{font-weight:900;color:transparent;-webkit-text-stroke:1.1px #9e9e9e;letter-spacing:.04em}}
.trace .r{{-webkit-text-stroke-color:#e53935}}
.tl{{font-size:34pt;line-height:1.45;border-bottom:2px solid #cfd8dc;white-space:nowrap}}
.lines{{width:100%;padding:.2cm .6cm}} .lines.one .tl{{border-bottom:0}}
/* portada */
.cover{{width:17.5cm;height:20cm;display:flex;flex-direction:column;align-items:center;justify-content:space-between;padding:1cm;
  border-color:#2e7d32;background:linear-gradient(#fff,#fff) padding-box}}
.cv-title{{font-family:"Alfa";font-size:36pt;color:#2e7d32;text-align:center;line-height:1.05}}
.cv-title span{{color:#1565c0;font-size:30pt}}
.cv-logo{{width:14cm}}
.cv-author{{font-size:20pt;font-weight:900;color:#c62828}}
/* fotos y dibujos */
.row2{{display:flex;gap:.5cm;width:100%;justify-content:center}}
.photo{{padding:.25cm}} .photo img{{display:block;height:11.5cm;width:auto;border-radius:.2cm}}
.photo.big img{{height:14.5cm}}
.drawbox{{width:8.6cm;height:12.05cm;position:relative}}
.drawbox.wide{{width:100%;height:7.6cm}}
.dlabel{{position:absolute;left:.35cm;top:.25cm;font-size:12pt;color:#aaa;font-weight:800}}
.tags{{display:grid;grid-template-columns:repeat(4,1fr);gap:.35cm;width:100%}}
.tag{{text-align:center;font-size:24pt;line-height:1.2;padding:.25cm .1cm}}
.tag small{{font-size:17pt}}
.tag .arrow{{display:block;color:#1565c0;font-size:18pt;transform:rotate(-90deg)}}
.sub{{align-self:flex-start;font-size:15pt;font-weight:900;color:#c62828}}
.row3{{display:grid;grid-template-columns:repeat(3,1fr);gap:.35cm;width:100%}}
.mini{{padding:.2cm}} .mini-draw{{height:4.6cm}} .mini-cap{{font-size:15pt;text-align:center;line-height:1.2}}
/* corazón */
.heart-wrap{{position:relative;width:18.5cm;height:16cm}}
.heart{{position:absolute;inset:0;width:100%;height:100%}}
.heart-text{{position:absolute;left:0;right:0;top:3.4cm;text-align:center}}
.heart-text .tl{{font-size:27pt;border:0;line-height:1.5}}
.badge{{font-size:26pt;padding:.3cm .7cm}}
.stickers{{flex:1;display:flex;justify-content:center}} .stickers img{{height:24cm}}
</style></head><body>{"".join(hojas)}</body></html>'''
open("libro.html", "w", encoding="utf-8").write(HTML)
print("OK")
