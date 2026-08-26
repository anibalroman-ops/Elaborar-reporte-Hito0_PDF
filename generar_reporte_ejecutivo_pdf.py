#!/usr/bin/env python3
"""Genera un PDF de revisión de CONTENIDO para el reporte ejecutivo institucional del Hito 0.

Este script es independiente del reporte científico completo y de su generador
(generar_reporte_hito0_editorial_v3.py / Reporte_Hito0_Editorial_V3.pdf), que no se
tocan. Produce una maquetación simple, a una columna, pensada solo para revisar el
texto del reporte ejecutivo antes de aplicarle la línea editorial definitiva que el
usuario entregará más adelante.

Uso:
    python3 generar_reporte_ejecutivo_pdf.py
    python3 generar_reporte_ejecutivo_pdf.py --input otro.md --output otro.pdf
"""
from __future__ import annotations

import argparse
import re
from pathlib import Path

from bs4 import BeautifulSoup, Tag
from markdown_it import MarkdownIt
from weasyprint import HTML

DEFAULT_INPUT = "Reporte_Hito0_Ejecutivo_Institucional_v1.md"
DEFAULT_OUTPUT = "Reporte_Hito0_Ejecutivo_Institucional_v1_BORRADOR.pdf"

CSS = """
@page {
    size: A4;
    margin: 20mm 22mm 22mm 22mm;
    @top-center {
        content: "Reporte Institucional del Hito 0 — Versión ejecutiva (borrador de contenido)";
        font-family: "Liberation Sans", sans-serif;
        font-size: 8pt;
        color: #8a94a3;
    }
    @bottom-center {
        content: counter(page);
        font-family: "Liberation Sans", sans-serif;
        font-size: 9pt;
        color: #6b7280;
    }
}
@page :first {
    @top-center { content: none; }
    @bottom-center { content: none; }
}

* { box-sizing: border-box; }

body {
    font-family: "Liberation Serif", "Times New Roman", serif;
    font-size: 11pt;
    line-height: 1.5;
    color: #1a1a1a;
}

/* ---------- Portada ---------- */
div.cover {
    page-break-after: always;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    text-align: center;
    height: 255mm;
}
div.cover h1 {
    font-family: "Liberation Sans", sans-serif;
    font-size: 27pt;
    font-weight: bold;
    color: #16233d;
    margin: 0 0 4mm 0;
    line-height: 1.25;
    max-width: 140mm;
}
div.cover h2 {
    font-family: "Liberation Sans", sans-serif;
    font-size: 15pt;
    font-weight: normal;
    color: #3c5c86;
    margin: 0 0 18mm 0;
    border: none;
    padding: 0;
}
div.cover p {
    font-family: "Liberation Sans", sans-serif;
    font-size: 12pt;
    color: #2a2a2a;
    margin: 1.5mm 0;
}
div.cover p strong {
    display: inline-block;
    padding: 2mm 5mm;
    border: 0.6pt solid #16233d;
    border-radius: 1mm;
    color: #16233d;
    margin-top: 10mm;
}
div.cover hr {
    display: none;
}

/* ---------- Encabezados ---------- */
h1 {
    font-family: "Liberation Sans", sans-serif;
    font-size: 19pt;
    color: #16233d;
    margin: 0 0 8mm 0;
    padding-bottom: 3mm;
    border-bottom: 1.2pt solid #16233d;
    page-break-after: avoid;
}
body > h1 {
    page-break-before: always;
}
h2 {
    font-family: "Liberation Sans", sans-serif;
    font-size: 14pt;
    color: #1f3a5f;
    margin: 9mm 0 3mm 0;
    page-break-after: avoid;
}
h3 {
    font-family: "Liberation Sans", sans-serif;
    font-size: 12pt;
    color: #2a4a72;
    margin: 6mm 0 2mm 0;
    page-break-after: avoid;
}
h3:first-child, h2:first-child { margin-top: 0; }

/* ---------- Texto ---------- */
p { margin: 0 0 3mm 0; text-align: justify; orphans: 3; widows: 3; }
strong { color: #111827; }
em { color: #33415c; }

ul, ol { margin: 0 0 3mm 0; padding-left: 6mm; }
li { margin: 0 0 1.2mm 0; }
li > ul, li > ol { margin-top: 1.2mm; }

/* Índice: sin viñetas, look de tabla de contenidos */
h2#índice + ul,
h2#indice + ul {
    list-style: none;
    padding-left: 0;
}
h2#índice ul, h2#indice ul,
h2#índice + ul ul, h2#indice + ul ul {
    list-style: none;
    padding-left: 6mm;
}
h2#índice li, h2#indice li,
h2#índice + ul li, h2#indice + ul li {
    margin-bottom: 0.8mm;
}
h2#índice + ul > li, h2#indice + ul > li {
    margin-top: 3mm;
    font-family: "Liberation Sans", sans-serif;
}
h2#índice + ul > li > ul li, h2#indice + ul > li > ul li {
    font-size: 9.6pt;
    color: #333;
}

blockquote {
    margin: 4mm 0 4mm 0;
    padding: 3mm 6mm;
    border-left: 2.2pt solid #1f3a5f;
    background: #f4f6f9;
    font-style: italic;
    color: #16233d;
    page-break-inside: avoid;
}
blockquote p { margin: 0; text-align: left; }

hr {
    border: none;
    border-top: 0.5pt solid #c9d1dc;
    margin: 6mm 0;
}

/* ---------- Tablas ---------- */
table {
    width: 100%;
    border-collapse: collapse;
    margin: 3mm 0 5mm 0;
    font-size: 8.6pt;
    line-height: 1.32;
}
th, td {
    border: 0.5pt solid #c9d1dc;
    padding: 1.6mm 2.2mm;
    text-align: left;
    vertical-align: top;
}
th {
    background: #e8edf4;
    font-family: "Liberation Sans", sans-serif;
    font-weight: bold;
    color: #16233d;
}
tr { page-break-inside: avoid; }
em, table em { color: inherit; }
table + p em, p > em:only-child { display: block; font-size: 9pt; color: #555; }

/* Glosario y notas finales */
h1:last-of-type + p em:only-child,
body > p:last-child em:only-child {
    display: block;
    text-align: center;
    color: #555;
    margin-top: 8mm;
}
"""


def strip_html_comments(text: str) -> str:
    return re.sub(r"<!--.*?-->", "", text, flags=re.DOTALL)


def markdown_to_html(md_text: str) -> str:
    md = MarkdownIt("commonmark", {"html": True, "breaks": False})
    md.enable("table")
    return md.render(md_text)


def wrap_cover(soup: BeautifulSoup) -> None:
    """Envuelve el título, subtítulo y metadatos iniciales (hasta el primer <hr>) en una portada."""
    hr = soup.find("hr")
    if hr is None:
        return
    nodes_before = []
    node = hr.previous_sibling
    while node is not None:
        nodes_before.append(node)
        node = node.previous_sibling
    nodes_before.reverse()

    cover_div = soup.new_tag("div")
    cover_div["class"] = "cover"
    hr.insert_before(cover_div)
    for n in nodes_before:
        cover_div.append(n.extract())
    cover_div.append(hr.extract())


def build_html(md_path: Path) -> str:
    raw = strip_html_comments(md_path.read_text(encoding="utf-8-sig"))
    body_html = markdown_to_html(raw)
    soup = BeautifulSoup(body_html, "html.parser")
    wrap_cover(soup)
    return (
        "<!DOCTYPE html>\n"
        '<html lang="es">\n<head>\n<meta charset="utf-8">\n'
        "<title>Reporte Institucional del Hito 0 — Versión ejecutiva (borrador)</title>\n"
        f"<style>{CSS}</style>\n</head>\n<body>\n{soup}\n</body>\n</html>"
    )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", default=DEFAULT_INPUT)
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    md_path = Path(args.input)
    out_path = Path(args.output)
    html_text = build_html(md_path)
    out_path.with_suffix(".html").write_text(html_text, encoding="utf-8")
    HTML(string=html_text, base_url=str(md_path.parent)).write_pdf(str(out_path))
    print(f"PDF generado: {out_path} ({out_path.stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
