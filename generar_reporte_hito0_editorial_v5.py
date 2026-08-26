#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Generador editorial del Diagnostico Hito 0 - Medición Atributos i+e.

Pipeline:
    Markdown -> HTML semántico -> CSS paginado -> PDF (WeasyPrint)

El diseño implementa y refina la especificación "Diseño maestro editorial — Diagnostico Hito 0, V4": A4, narrativa a dos columnas, portadillas de
Parte, tipografía serif/sans jerarquizada, figuras complejas y tablas anchas a ancho completo, tablas pequeñas en una columna,
paleta teal/naranja/grafito, captions editoriales, encabezados corridos,
folios, índice navegable, marcadores PDF, flujo de columnas estrictamente secuencial y referencias específicas al final de cada Parte.

Dependencias:
    pip install weasyprint markdown-it-py beautifulsoup4 pillow pymupdf

Uso recomendado (PowerShell / VS Code):
    python .\generar_reporte_hito0_editorial_v3.py \
      --md ".\Reporte_Hito0_Integrado_MEJORADO_v13_REORDEN_FIGURAS_8_13.md" \
      --figuras ".\01_figuras" \
      --salida ".\Reporte_Hito0_Editorial_V3.pdf"

También acepta un ZIP en --figuras.

Notas:
- No reescribe el contenido sustantivo del Markdown.
- El índice general y las portadillas son elementos editoriales derivados de
  la estructura del documento.
- Las expresiones LaTeX simples usadas en el reporte se convierten a símbolos
  Unicode/HTML para evitar mostrar barras invertidas en el PDF.
- No descarga ni distribuye fuentes. Usa Source Serif 4 / Noto Serif y Noto
  Sans si están instaladas, con fallbacks del sistema.
"""

from __future__ import annotations

import argparse
import html as html_lib
import os
import re
import shutil
import sys
import tempfile
import unicodedata
import zipfile
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

try:
    from bs4 import BeautifulSoup, NavigableString, Tag
except ImportError:
    print("ERROR: falta beautifulsoup4. Instálalo con: pip install beautifulsoup4")
    raise SystemExit(1)

try:
    from markdown_it import MarkdownIt
except ImportError:
    print("ERROR: falta markdown-it-py. Instálalo con: pip install markdown-it-py")
    raise SystemExit(1)

try:
    from PIL import Image as PILImage
except ImportError:
    print("ERROR: falta Pillow. Instálalo con: pip install pillow")
    raise SystemExit(1)

try:
    from weasyprint import HTML
except ImportError:
    print("ERROR: falta WeasyPrint. Instálalo con: pip install weasyprint")
    raise SystemExit(1)

try:
    import pymupdf as fitz  # PyMuPDF (API actual)
except ImportError:
    print("ERROR: falta PyMuPDF. Instálalo con: pip install pymupdf")
    raise SystemExit(1)


# =============================================================================
# CONFIGURACIÓN EDITORIAL
# =============================================================================

SUPPORTED_IMAGES = {".png", ".jpg", ".jpeg", ".webp", ".svg"}

# Figuras cuyo contenido (multipanel, ejes, texto denso) no se lee bien
# comprimido en el ancho de una sola columna narrativa (~85mm): se fuerzan
# a ancho completo ("hero") aunque su relación de aspecto no lo sugiera.
# figura-01 se incluye también para que deje de ocupar una página dedicada
# completa (con espacio en blanco alrededor) y en su lugar fluya dentro
# del cuerpo, igual que las figuras 2 y 3.
#
# Se identifican por el slug descriptivo del nombre de archivo (la parte
# después del número, p. ej. "recorrido-general" en "figura-01-recorrido-
# general.png") y NO por su número de figura: el número cambia cada vez
# que el Markdown se reordena o renumera (ya ha ocurrido varias veces en
# este documento), mientras que el slug identifica la figura de forma
# estable entre versiones.
FORCE_HERO_FIGURE_SLUGS = {
    "recorrido-general",
    "arquitectura-conceptual",
    "instrumentos",
    "distribucion-puntaje",
    "focalizacion-rasch",
    "mapa-persona-item-rasch",
    "informacion-sem-rasch",
    "estructura-empirica",
    "pcm-mapa-grupo-umbrales",
    "parametros-paso-pcm2",
    "funcionamiento-categorias-pcm",
    "triangulacion-juicio-experto",
    "trazabilidad-refinamiento",
}

# Tablas que, pese a ser "pequeñas" según la heurística de densidad, se
# leen mejor a ancho completo por el largo de su contenido textual. Se
# identifican por el texto del título (no por su número, por la misma
# razón que FORCE_HERO_FIGURE_SLUGS): basta con que el título de la tabla
# CONTENGA uno de estos textos.
FORCE_WIDE_TABLE_TITLES = {"estado actual del sistema"}

# Figuras que se compactan a un alto máximo fijo (mm), a pedido explícito,
# independientemente de si dejan o no espacio en blanco: se identifican
# por el slug del nombre de archivo, igual que FORCE_HERO_FIGURE_SLUGS.
COMPACT_FIGURE_MAX_HEIGHT_MM = {
    "parametros-paso-pcm2": 65.0,
    "funcionamiento-categorias-pcm": 68.0,
}


MASTER_CSS = r"""
:root {
  --teal: #00A499;
  --orange: #E77500;
  --graphite: #394049;
  --text: #23282b;
  --cream: #f6f4ef;
  --teal-dark: #0b6f69;
  --teal-deep: #0b5e59;
  --title-dark: #1f2c30;
  --secondary: #4e5659;
  --tertiary: #687174;
  --line: #ccd6d4;
  --line-soft: #dce3e1;
  --teal-bg: #f3f8f7;
  --warm-bg: #fbf5ef;
  --font-serif: "Source Serif 4", "Noto Serif", "DejaVu Serif", Georgia, serif;
  --font-sans: "Noto Sans", Arial, sans-serif;
}

@page {
  size: A4;
  margin: 18mm 16mm 19mm 16mm;

  @top-left {
    content: string(current-part);
    font-family: var(--font-sans);
    font-size: 6.6pt;
    color: var(--tertiary);
  }
  @top-right {
    content: string(current-section);
    font-family: var(--font-sans);
    font-size: 6.6pt;
    color: var(--tertiary);
  }
  @bottom-left {
    content: "Facultad de Ingeniería · Universidad de Santiago de Chile";
    font-family: var(--font-sans);
    font-size: 6.4pt;
    color: #7a8386;
  }
  @bottom-right { content: none; }
}

@page cover {
  size: A4;
  margin: 0;
  @top-left { content: none; }
  @top-right { content: none; }
  @bottom-left { content: none; }
  @bottom-right { content: none; }
}

@page part {
  size: A4;
  margin: 0;
  @top-left { content: none; }
  @top-right { content: none; }
  @bottom-left { content: none; }
  @bottom-right { content: none; }
}

@page landscape {
  size: A4 landscape;
  margin: 16mm 16mm 17mm 16mm;
  @top-left {
    content: string(current-part);
    font-family: var(--font-sans);
    font-size: 6.6pt;
    color: var(--tertiary);
  }
  @top-right {
    content: string(current-section);
    font-family: var(--font-sans);
    font-size: 6.6pt;
    color: var(--tertiary);
  }
  @bottom-left {
    content: "Facultad de Ingeniería · Universidad de Santiago de Chile";
    font-family: var(--font-sans);
    font-size: 6.4pt;
    color: #7a8386;
  }
  @bottom-right { content: none; }
}

@page figurepage {
  size: A4;
  margin: 10mm 11mm 12mm 11mm;
  @top-left { content: none; }
  @top-right { content: none; }
  @bottom-left { content: none; }
  @bottom-right { content: none; }
}

* { box-sizing: border-box; }

html { font-size: 9.25pt; }
body {
  margin: 0;
  color: var(--text);
  font-family: var(--font-serif);
  font-size: 9.25pt;
  line-height: 1.48;
  text-rendering: optimizeLegibility;
  hyphens: auto;
}

p {
  margin: 0 0 3.15mm;
  text-align: justify;
  /* orphans/widows bajos a propósito: con column-fill:balance, exigir un
     mínimo alto de líneas por fragmento fuerza a WeasyPrint a mover un
     párrafo COMPLETO a la columna siguiente en vez de partirlo, lo que
     deja huecos en blanco al final de la columna anterior. Con 2 se
     preserva el mínimo tipográfico razonable (nunca 1 sola línea huérfana)
     sin bloquear el reparto fino entre columnas. */
  orphans: 2;
  widows: 2;
}

strong { font-weight: 700; }
em { font-style: italic; }
a { color: var(--teal-deep); text-decoration: none; }
code {
  font-family: var(--font-sans);
  font-size: .92em;
  color: #33464a;
  overflow-wrap: anywhere;
}

/* ------------------------------------------------------------------------- */
/* PORTADA */
/* ------------------------------------------------------------------------- */
.cover {
  page: cover;
  break-after: page;
  height: 297mm;
  position: relative;
  background: var(--cream);
  overflow: hidden;
}
.cover::before {
  content: "";
  position: absolute;
  left: 0; right: 0; top: 0;
  height: 43%;
  background: var(--teal);
}
.cover::after {
  content: "";
  position: absolute;
  left: 22mm;
  bottom: 23mm;
  width: 27mm;
  height: 2.2mm;
  background: var(--orange);
}
.cover-inner {
  position: relative;
  z-index: 2;
  height: 100%;
  padding: 26mm 22mm 24mm;
  display: flex;
  flex-direction: column;
}
.cover-kicker {
  font-family: var(--font-sans);
  text-transform: uppercase;
  letter-spacing: .16em;
  font-size: 8.6pt;
  font-weight: 800;
  color: #eafffc;
  margin-bottom: 12mm;
}
.cover-title {
  max-width: 166mm;
  font-family: var(--font-sans);
  font-size: 25pt;
  line-height: 1.04;
  font-weight: 800;
  color: #fff;
  margin: 0;
  bookmark-level: none;
}
.cover-meta {
  margin-top: auto;
  max-width: 150mm;
  font-family: var(--font-sans);
  color: var(--graphite);
}
.cover-meta .faculty {
  font-size: 13pt;
  font-weight: 700;
  line-height: 1.25;
}
.cover-meta .institution {
  font-size: 9.6pt;
  margin-top: 2mm;
  color: var(--secondary);
}
.cover-meta .descriptor {
  margin-top: 10mm;
  font-size: 8.2pt;
  text-transform: uppercase;
  letter-spacing: .12em;
  color: var(--teal-dark);
  font-weight: 700;
}

/* ------------------------------------------------------------------------- */
/* RESUMEN EJECUTIVO */
/* ------------------------------------------------------------------------- */
.executive-summary {
  page: auto;
  break-before: page;
  break-after: page;
  padding-top: 2mm;
}
.executive-summary > h1 {
  font-family: var(--font-sans);
  color: var(--title-dark);
  font-size: 20pt;
  line-height: 1.05;
  margin: 0 0 6mm;
  padding-top: 4mm;
  border-top: 5px solid var(--teal);
  bookmark-level: 1;
}
.executive-summary .summary-body {
  column-count: 1;
}
.executive-summary .summary-body > p:first-of-type {
  font-size: 10.1pt;
  line-height: 1.52;
  color: var(--graphite);
  margin-bottom: 5mm;
}
.executive-summary strong { color: #1e5350; }

/* ------------------------------------------------------------------------- */
/* ÍNDICE */
/* ------------------------------------------------------------------------- */
.toc {
  break-before: page;
  break-after: page;
  font-family: var(--font-sans);
}
.toc h1 {
  font-size: 20pt;
  color: var(--title-dark);
  border-top: 5px solid var(--teal);
  padding-top: 4mm;
  margin: 0 0 8mm;
  bookmark-level: none;
}
.toc-list { column-count: 2; column-gap: 10mm; }
.toc-part,
.toc-sec {
  break-inside: avoid;
  margin: 0 0 2.6mm;
}
.toc-part { margin-top: 3.5mm; }
.toc-part a,
.toc-sec a {
  display: flex;
  gap: 2mm;
  align-items: baseline;
  color: var(--graphite);
}
.toc-part a { font-weight: 800; color: var(--teal-dark); }
.toc-sec a { font-size: 8.15pt; }
.toc-sec { padding-left: 5mm; }
.toc-num { color: var(--orange); font-weight: 800; min-width: 9mm; }
.toc-text { flex: 1; }
.toc-page::after { content: target-counter(attr(data-href), page); }
.toc-page {
  margin-left: auto;
  color: var(--tertiary);
  font-variant-numeric: tabular-nums;
}

/* ------------------------------------------------------------------------- */
/* PORTADILLAS DE PARTE */
/* ------------------------------------------------------------------------- */
.part-opener {
  page: part;
  break-before: page;
  break-after: page;
  height: 297mm;
  margin: 0;
  padding: 24mm 20mm 18mm;
  background: var(--cream);
  position: relative;
  overflow: hidden;
  font-family: var(--font-sans);
}
.part-opener::before {
  content: "";
  position: absolute;
  left: 0; top: 0; right: 0;
  height: 41%;
  background: var(--teal);
}
.part-opener .part-content { position: relative; z-index: 2; }
.part-opener .part-kicker {
  color: #fff;
  text-transform: uppercase;
  letter-spacing: .16em;
  font-size: 9.6pt;
  font-weight: 800;
  margin: 0 0 7mm;
}
.part-opener h1.part-heading {
  margin: 0;
  color: #fff;
  bookmark-level: 1;
  string-set: current-part content();
}
.part-opener .part-number {
  display: block;
  font-size: 47pt;
  font-weight: 800;
  line-height: .9;
  margin-bottom: 7mm;
}
.part-opener .part-name {
  display: block;
  max-width: 160mm;
  font-size: 23.5pt;
  line-height: 1.06;
  font-weight: 750;
}
.part-opener .part-lower {
  position: absolute;
  top: 130mm;
  left: 20mm;
  right: 20mm;
}
.part-description {
  max-width: 158mm;
  color: var(--graphite);
  font-size: 10.6pt;
  line-height: 1.46;
  margin-bottom: 8mm;
}
.part-description p { text-align: left; margin-bottom: 3mm; }
.part-description blockquote {
  margin: 0 0 5mm;
  padding: 3.5mm 4.5mm;
  background: rgba(255,255,255,.48);
  border-left: 3px solid var(--orange);
}
.part-mini-index {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 0 12mm;
  border-top: 2px solid var(--teal);
  padding-top: 4mm;
}
.part-mini-index a {
  display: grid;
  grid-template-columns: 14mm 1fr;
  gap: 2.5mm;
  align-items: start;
  min-height: 16mm;
  padding: 4mm 0 4.5mm;
  border-bottom: 1px solid #d9e1df;
  font-size: 9.8pt;
  line-height: 1.24;
  color: var(--graphite);
}
.part-mini-index .mini-num {
  color: var(--orange);
  font-weight: 850;
  font-size: 13pt;
  line-height: 1;
}
.part-mini-index .mini-label { font-weight: 650; }
.part-mini-index.index-few a {
  min-height: 25mm;
  padding-top: 6mm;
  font-size: 11.4pt;
  line-height: 1.25;
}
.part-mini-index.index-few .mini-num { font-size: 16pt; }
.part-mini-index.index-many a {
  min-height: 13mm;
  padding: 3mm 0;
  font-size: 8.8pt;
}
.part-mini-index.index-many .mini-num { font-size: 11.5pt; }

/* ------------------------------------------------------------------------- */
/* CUERPO EN DOS COLUMNAS */
/* ------------------------------------------------------------------------- */
/* Las secciones ya no fuerzan página nueva: el flujo de una Parte es continuo.
   Esto elimina grandes huecos al final de secciones cortas. */
.section-major {
  break-before: auto;
  margin: 0;
}
.section-major + .section-major { margin-top: 5mm; }

/* Flujo editorial en dos columnas. Los cambios de sección/subsección no
   crean un nuevo flujo: solo los elementos de ancho completo lo hacen.

   NOTA TÉCNICA (WeasyPrint, ago-2026): se usa deliberadamente
   ``column-fill: balance`` y NO ``column-fill: auto``. El motor de
   renderizado tiene un bug confirmado en la implementación de
   ``column-fill: auto`` para contenedores multicolumna con altura auto:
   cuando el contenido de un ``.flow-columns`` no fuerza el desbordamiento
   de una columna completa (lo habitual, dado que estos tramos suelen ser
   cortos entre figuras/tablas), TODO el contenido se coloca en la columna
   izquierda y la columna derecha queda completamente en blanco -incluso
   en tramos que continúan en una página siguiente-. Esto es exactamente
   el defecto de "columnas mal ordenadas / espacios en blanco" que motivó
   esta reescritura. ``column-fill: balance`` no sufre este problema: usa
   el alto de página disponible como referencia real y reparte el
   contenido correctamente entre ambas columnas, tanto en la primera
   página de un tramo como en sus continuaciones. El orden de lectura
   (columna izquierda completa -> columna derecha -> página siguiente)
   se preserva igual con column-count:2, ya que "balance" solo cambia
   CUÁNTO contenido entra en cada columna, no el orden en que se lee. */
.flow-columns {
  column-count: 2;
  column-gap: 9mm;
  column-fill: balance;
}
.flow-columns::after { content: ""; display: block; clear: both; }

h2, h3, h4 {
  font-family: var(--font-sans);
  break-after: avoid;
  color: var(--title-dark);
}
h2 {
  font-size: 13.3pt;
  line-height: 1.12;
  margin: 0 0 4mm;
  padding: 3.2mm 0 1.5mm;
  border-top: 4px solid var(--teal);
  border-bottom: 1px solid var(--line);
  string-set: current-section content();
  bookmark-level: 2;
}
h3 {
  font-size: 11pt;
  line-height: 1.15;
  margin: 5mm 0 2.5mm;
  color: #0e6f69;
  bookmark-level: 3;
}
h4 {
  font-size: 9.6pt;
  line-height: 1.2;
  margin: 4mm 0 2mm;
  color: var(--graphite);
  bookmark-level: 4;
}

ul, ol {
  margin: 2mm 0 4mm;
  padding-left: 5mm;
}
li {
  margin-bottom: 2.1mm;
  /* auto (no avoid) por la misma razón que orphans/widows arriba: con
     column-fill:balance, un ítem que no puede partirse fuerza a mover
     la lista completa a la columna siguiente y deja hueco en blanco. */
  break-inside: auto;
}
li::marker { color: var(--teal); }

.attrs {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 2.2mm 7mm;
  margin: 4mm 0 6mm;
  padding: 5mm 6mm 5mm 11mm;
  background: var(--cream);
  font-family: var(--font-sans);
}
.attrs li { margin: 0; padding-left: 1mm; }
.attrs li::marker { color: var(--orange); font-weight: 800; }

hr {
  border: 0;
  border-top: 1px solid var(--line);
  margin: 5mm 0;
}

/* ------------------------------------------------------------------------- */
/* FIGURAS */
/* ------------------------------------------------------------------------- */
.figure {
  margin: 6mm 0 7mm;
  break-inside: avoid;
  width: 100%;
}
.figure.hero { break-before: auto; }
.figure.landscape {
  page: landscape;
  break-before: page;
  break-after: page;
}
.figure.fullpage {
  page: figurepage;
  break-before: page;
  break-after: page;
  margin: 0;
  width: 100%;
}
.figure.fullpage .figure-frame {
  border: 0;
  padding: 0;
}
.figure.fullpage img {
  width: 100%;
  max-height: 250mm;
  object-fit: contain;
}
.figure.fullpage figcaption { margin-top: 3mm; }
.figure-frame {
  border-top: 2px solid var(--teal);
  border-bottom: 1px solid #d9e2e0;
  padding: 3mm 0;
  background: #fff;
  text-align: center;
}
.figure img {
  width: 100%;
  height: auto;
  max-height: 150mm;
  object-fit: contain;
  display: block;
  margin: 0 auto;
}
.figure.hero img { max-height: 165mm; }
.figure.landscape img { max-height: 155mm; }
.figure figcaption {
  font-family: var(--font-serif);
  font-size: 7.45pt;
  line-height: 1.35;
  color: var(--secondary);
  text-align: left;
  margin-top: 2mm;
  break-inside: avoid;
}
.figure figcaption strong {
  font-family: var(--font-sans);
  color: var(--teal-dark);
  font-weight: 800;
}

/* ------------------------------------------------------------------------- */
/* TABLAS */
/* ------------------------------------------------------------------------- */
.table-title {
  font-family: var(--font-serif);
  font-size: 7.7pt;
  line-height: 1.35;
  color: var(--secondary);
  margin: 5mm 0 2.3mm;
  break-after: avoid;
}
.table-title strong {
  font-family: var(--font-sans);
  color: var(--teal-dark);
  font-weight: 800;
}
.panel-label {
  font-family: var(--font-sans);
  font-size: 7.8pt;
  font-weight: 700;
  color: var(--graphite);
  margin: 3mm 0 1.5mm;
  break-after: avoid;
}
.table-group {
  margin: 4mm 0 5mm;
}
.table-group.small-table {
  margin: 3mm 0 4.5mm;
  break-inside: avoid;
}
.table-group.wide-table {
  break-inside: auto;
}
.table-block {
  margin: 0;
  break-inside: auto;
}
.table-group.small-table table {
  font-size: 7.75pt;
  line-height: 1.28;
}
.table-group.small-table .table-title {
  margin-top: 0;
  font-size: 7.8pt;
}
table {
  width: 100%;
  border-collapse: collapse;
  font-family: var(--font-sans);
  font-size: 7.35pt;
  line-height: 1.28;
}
thead { display: table-header-group; }
tr { break-inside: avoid; }
th {
  background: var(--graphite);
  color: #fff;
  text-align: left;
  padding: 2.3mm 2.5mm;
  font-weight: 700;
  vertical-align: bottom;
}
td {
  padding: 2mm 2.5mm;
  border-bottom: 1px solid var(--line-soft);
  vertical-align: top;
}
tbody tr:nth-child(even) td { background: #f6f8f7; }
table.matrix { font-size: 6.2pt; line-height: 1.22; }
table.medium { font-size: 6.8pt; line-height: 1.24; }
.table-note {
  font-size: 7.5pt;
  color: var(--tertiary);
  margin-top: -2mm;
  margin-bottom: 5mm;
}

/* ------------------------------------------------------------------------- */
/* CAJAS / CITAS / CADENAS */
/* ------------------------------------------------------------------------- */
blockquote {
  margin: 3.5mm 0 4.5mm;
  padding: 3.5mm 4.5mm;
  background: #f7f9f8;
  border-left: 3px solid var(--line);
  color: var(--graphite);
  break-inside: auto;
}
blockquote.note-box,
blockquote.keybox,
blockquote.question,
blockquote.chain { break-inside: avoid; }
blockquote p:last-child { margin-bottom: 0; }
.keybox {
  margin: 5mm 0 6mm;
  padding: 5mm 6mm;
  background: var(--teal-bg);
  border-left: 4px solid var(--teal);
  font-size: 9.45pt;
  line-height: 1.47;
  break-inside: avoid;
}
.note-box {
  margin: 5mm 0;
  padding: 4.5mm 5.5mm;
  background: var(--warm-bg);
  border-left: 3px solid var(--orange);
  break-inside: avoid;
}
.question {
  margin: 5mm 0 6mm;
  padding: 6mm 8mm;
  background: var(--graphite);
  color: #fff;
  font-family: var(--font-sans);
  font-size: 13pt;
  font-weight: 700;
  line-height: 1.35;
  text-align: center;
  break-inside: avoid;
}
.question p { text-align: center; margin: 0; }
.chain {
  margin: 5mm 0 6mm;
  padding: 4mm 5mm;
  border: 1px solid var(--line);
  background: #fbfcfb;
  text-align: center;
  font-family: var(--font-sans);
  font-weight: 700;
  color: #1d5652;
  font-size: 8.9pt;
  line-height: 1.6;
  break-inside: avoid;
}
.chain p { text-align: center; margin: 0; }

/* ------------------------------------------------------------------------- */
/* ECUACIONES */
/* ------------------------------------------------------------------------- */
.math-inline {
  font-family: var(--font-serif);
  font-style: italic;
  white-space: nowrap;
}
.equation {
  text-align: center;
  font-family: var(--font-serif);
  font-style: italic;
  font-size: 13pt;
  padding: 3mm 4mm;
  margin: 3mm 0 4mm;
  color: #1b4e4a;
  background: #f7faf9;
  border-top: 1px solid #d8e3e1;
  border-bottom: 1px solid #d8e3e1;
  break-inside: avoid;
}
.equation sub, .math-inline sub { font-size: .72em; }
.equation sup, .math-inline sup { font-size: .72em; }

/* ------------------------------------------------------------------------- */
/* REFERENCIAS */
/* ------------------------------------------------------------------------- */
.references-section { break-before: page; }
.references-section > h1 {
  font-family: var(--font-sans);
  font-size: 18.8pt;
  line-height: 1.05;
  color: var(--title-dark);
  margin: 0 0 6mm;
  padding-top: 3.2mm;
  border-top: 5px solid var(--teal);
  string-set: current-section content();
  bookmark-level: 1;
}
.references-columns {
  column-count: 2;
  column-gap: 9mm;
  font-size: 8.2pt;
  line-height: 1.36;
}
.references-columns p {
  margin: 0 0 2.6mm;
  padding-left: 5mm;
  text-indent: -5mm;
  text-align: left;
  overflow-wrap: anywhere;
  break-inside: avoid;
}

/* Referencias específicas de cada Parte. Forman parte del mismo flujo de dos
   columnas y, por tanto, respetan izquierda -> derecha -> página siguiente. */
.part-references-heading {
  font-family: var(--font-sans);
  font-size: 10.8pt;
  line-height: 1.15;
  color: var(--teal-dark);
  margin: 7mm 0 3mm;
  padding-top: 2.5mm;
  border-top: 1.5px solid var(--teal);
  break-after: avoid;
  bookmark-level: 3;
}
.part-reference {
  font-size: 8.15pt;
  line-height: 1.34;
  margin: 0 0 2.5mm;
  padding-left: 4.5mm;
  text-indent: -4.5mm;
  text-align: left;
  overflow-wrap: anywhere;
  /* Sin esto, una referencia puede partirse entre columna/página: la
     sangría negativa (pensada solo para la primera línea real) se vuelve
     a aplicar a la línea que abre el fragmento siguiente, lo que la
     desplaza fuera del área de la columna y recorta sus primeros
     caracteres. Al impedir el corte, cada referencia (un párrafo corto)
     se mueve entera a donde quepa, igual que en .references-columns p. */
  break-inside: avoid;
}

/* ------------------------------------------------------------------------- */
/* RENDERIZADO POR UNIDADES */
/* ------------------------------------------------------------------------- */
.running-part {
  string-set: current-part content();
  height: 0;
  max-height: 0;
  overflow: hidden;
  color: transparent;
  font-size: 0;
  line-height: 0;
  margin: 0; padding: 0;
}
.running-section {
  string-set: current-section content();
  height: 0;
  max-height: 0;
  overflow: hidden;
  color: transparent;
  font-size: 0;
  line-height: 0;
  margin: 0; padding: 0;
}
.render-landscape { page: landscape; }
.render-fullpage { page: figurepage; }
.fullpage-unit { page: figurepage; }
.fullpage-unit > .figure.fullpage {
  page: auto;
  break-before: auto;
  break-after: auto;
}
.landscape-unit {
  page: landscape;
}
.landscape-unit > .figure.landscape {
  page: auto;
  break-before: auto;
  break-after: auto;
  margin-top: 0;
}
.render-unit > .section-major,
.render-unit > .executive-summary,
.render-unit > .part-opener,
.render-unit > .references-section,
.render-unit > .toc {
  break-before: auto;
}

/* ------------------------------------------------------------------------- */
/* MISC */
/* ------------------------------------------------------------------------- */
.small-note { font-size: 7.5pt; color: var(--tertiary); }
.page-break { break-before: page; }
.no-print { display: none; }

/* ========================================================================== */
/* V4 - SISTEMA EDITORIAL: PROGRESIÓN / MEDICIÓN / PROFUNDIDAD               */
/* ========================================================================== */
:root {
  --blue: #173CF5;
  --blue-2: #3557F2;
  --blue-3: #5670F1;
  --blue-4: #7F91F2;
  --blue-5: #A8B4F5;
  --blue-6: #D7DCFA;
  --ink: #14171d;
  --ink-2: #303641;
  --muted: #69717e;
  --paper: #fbfaf7;
  --rule: #d9dde6;
  --rule-strong: #bfc6d6;
  --panel: #f4f6fb;
  --font-serif: "Noto Serif", "DejaVu Serif", Georgia, serif;
  --font-sans: "Inter", "Noto Sans", Arial, sans-serif;

  /* Compatibilidad con reglas heredadas del V3 */
  --teal: var(--blue);
  --teal-dark: #2949c7;
  --teal-deep: #233da3;
  --orange: var(--blue-3);
  --graphite: var(--ink-2);
  --text: #242830;
  --cream: var(--paper);
  --title-dark: var(--ink);
  --secondary: #59616d;
  --tertiary: #747c87;
  --line: var(--rule);
  --line-soft: #e7e9ef;
  --teal-bg: #f3f5ff;
  --warm-bg: #faf9f6;
}

@page {
  size: A4;
  margin: 17mm 17mm 18mm 17mm;
  @top-left {
    font-family: var(--font-sans);
    font-size: 6.25pt;
    font-weight: 500;
    letter-spacing: .03em;
    color: #7b8290;
  }
  @top-right {
    font-family: var(--font-sans);
    font-size: 6.25pt;
    font-weight: 500;
    color: #7b8290;
  }
  @bottom-left {
    content: "Facultad de Ingeniería · Universidad de Santiago de Chile";
    font-family: var(--font-sans);
    font-size: 6.1pt;
    color: #8a909a;
  }
  @bottom-right {
    content: counter(page);
    font-family: var(--font-sans);
    font-size: 7pt;
    font-weight: 650;
    color: #565e6b;
  }
}

html { font-size: 9.15pt; }
body {
  color: #252931;
  font-family: var(--font-serif);
  font-size: 9.15pt;
  line-height: 1.50;
  background: white;
}
p { margin-bottom: 3.25mm; }
strong { color: #171b22; }
a { color: #2949c7; }

/* PORTADA ----------------------------------------------------------------- */
.cover {
  background: var(--paper);
  color: var(--ink);
}
.cover::before {
  content: "";
  position: absolute;
  left: 23mm;
  right: 23mm;
  top: 23mm;
  height: 0;
  border-top: .55pt solid #222730;
  background: transparent;
}
.cover::after {
  content: "";
  position: absolute;
  z-index: 1;
  width: 61mm;
  height: 61mm;
  border-radius: 50%;
  left: 108mm;
  top: 93mm;
  background: #173CF5;
  box-shadow:
    0 10mm 0 #294AF4,
    0 20mm 0 #3D5CF3,
    0 30mm 0 #526CF2,
    0 40mm 0 #6B80F2,
    0 50mm 0 #8797F3,
    0 60mm 0 #A1AEF4,
    0 70mm 0 #BBC4F6,
    0 80mm 0 #D4DAF9;
}
.cover-inner {
  padding: 20mm 23mm 20mm;
  z-index: 2;
}
.cover-header {
  display: flex;
  justify-content: space-between;
  font-family: var(--font-sans);
  font-size: 6.7pt;
  font-weight: 600;
  letter-spacing: .06em;
  color: #3f4652;
  margin-bottom: 25mm;
}
.cover-kicker {
  font-family: var(--font-sans);
  font-size: 7.4pt;
  font-weight: 600;
  letter-spacing: .08em;
  color: #6b7380;
  text-transform: uppercase;
  margin: 0 0 5mm;
}
.cover-title {
  max-width: 105mm;
  font-family: var(--font-sans);
  font-size: 36pt;
  line-height: .98;
  letter-spacing: -.035em;
  font-weight: 750;
  color: #11151b;
  margin: 0;
}
.cover-subtitle {
  margin-top: 8mm;
  font-family: var(--font-sans);
  font-size: 12pt;
  font-weight: 500;
  color: #303641;
  position: relative;
  padding-top: 5mm;
}
.cover-subtitle::before {
  content: "";
  position: absolute;
  top: 0;
  left: 0;
  width: 15mm;
  height: .8mm;
  background: var(--blue);
}
.cover-meta {
  margin-top: auto;
  max-width: 78mm;
  padding-top: 4mm;
  border-top: .55pt solid #303641;
  font-family: var(--font-sans);
  color: #3f4652;
}
.cover-meta .faculty {
  font-size: 8.4pt;
  font-weight: 650;
  line-height: 1.25;
}
.cover-meta .institution {
  font-size: 7.6pt;
  margin-top: 1.3mm;
  color: #686f7a;
}
.cover-meta .descriptor {
  margin-top: 4mm;
  font-size: 6.3pt;
  line-height: 1.35;
  letter-spacing: .07em;
  color: #6c7380;
  font-weight: 550;
}

/* RESUMEN / ÍNDICE -------------------------------------------------------- */
.executive-summary > h1,
.toc h1 {
  font-family: var(--font-sans);
  color: var(--ink);
  font-size: 22pt;
  line-height: 1.02;
  letter-spacing: -.025em;
  border-top: .6pt solid #303641;
  padding-top: 6mm;
  margin-bottom: 8mm;
  position: relative;
}
.executive-summary > h1::before,
.toc h1::before {
  content: "";
  display: inline-block;
  width: 3.4mm;
  height: 3.4mm;
  border-radius: 50%;
  background: var(--blue);
  margin-right: 3mm;
  vertical-align: 1.5mm;
}
.executive-summary .summary-body > p:first-of-type {
  font-size: 10pt;
  line-height: 1.55;
  color: #333946;
}
.executive-summary strong { color: #1d327e; }
.toc-part a { color: #1b2f7a; }
.toc-num { color: var(--blue); }

/* PORTADILLAS DE PARTE ---------------------------------------------------- */
.part-opener {
  background: var(--paper);
  padding: 22mm 22mm 18mm;
  color: var(--ink);
}
.part-opener::before {
  content: "";
  position: absolute;
  left: 22mm;
  right: 22mm;
  top: 22mm;
  height: 0;
  border-top: .55pt solid #262b33;
  background: transparent;
}
.part-opener::after {
  content: "";
  position: absolute;
  z-index: 1;
  width: 98mm;
  height: 98mm;
  border-radius: 50%;
  right: -39mm;
  top: 56mm;
  background: #173CF5;
  box-shadow:
    -11mm 11mm 0 #3557F2,
    -22mm 22mm 0 #5A72F1,
    -33mm 33mm 0 #8293F2,
    -44mm 44mm 0 #AEB9F5,
    -55mm 55mm 0 #D7DCFA;
}
.part-opener .part-content,
.part-opener .part-lower { z-index: 2; }
.part-opener .part-kicker {
  color: #5e6570;
  text-transform: uppercase;
  letter-spacing: .08em;
  font-size: 6.8pt;
  font-weight: 600;
  margin: 8mm 0 14mm;
}
.part-opener h1.part-heading { color: var(--ink); }
.part-opener .part-number {
  color: var(--blue);
  font-size: 16pt;
  line-height: 1;
  letter-spacing: .01em;
  margin-bottom: 5mm;
  font-weight: 650;
}
.part-opener .part-name {
  max-width: 108mm;
  font-size: 29pt;
  line-height: 1.02;
  letter-spacing: -.028em;
  font-weight: 720;
  color: #12161d;
}
.part-opener .part-lower {
  top: 168mm;
  left: 22mm;
  right: 22mm;
}
.part-description {
  max-width: 108mm;
  font-family: var(--font-serif);
  color: #3b414a;
  font-size: 9.4pt;
  line-height: 1.47;
  margin-bottom: 6mm;
}
.part-mini-index {
  grid-template-columns: 1fr 1fr;
  gap: 0 10mm;
  border-top: .7pt solid #2e343d;
  padding-top: 3mm;
}
.part-mini-index a {
  grid-template-columns: 12mm 1fr;
  min-height: 13mm;
  padding: 3mm 0 3.5mm;
  border-bottom: .45pt solid #d8dce4;
  font-size: 8.6pt;
  color: #3c424c;
}
.part-mini-index .mini-num {
  color: var(--blue);
  font-size: 10.5pt;
  font-weight: 700;
}

/* CUERPO ------------------------------------------------------------------ */
.flow-columns { column-gap: 9mm; }
h2, h3, h4 { font-family: var(--font-sans); color: var(--ink); }
h2 {
  font-size: 13.7pt;
  line-height: 1.10;
  letter-spacing: -.015em;
  margin: 0 0 4mm;
  padding: 3mm 0 1.6mm;
  border-top: .55pt solid #cfd4df;
  border-bottom: 0;
  font-weight: 700;
}
h2::before {
  content: "";
  display: inline-block;
  width: 2.7mm;
  height: 2.7mm;
  border-radius: 50%;
  background: var(--blue);
  margin-right: 2.3mm;
  vertical-align: 1.0mm;
}
h3 {
  font-size: 10.8pt;
  line-height: 1.18;
  margin: 4.8mm 0 2.2mm;
  color: #1e3b9c;
  font-weight: 680;
}
h4 {
  font-size: 9.3pt;
  line-height: 1.22;
  color: #3b424d;
  font-weight: 650;
}
li::marker { color: #2445df; }
.attrs {
  background: #f6f7fb;
  border-top: .6pt solid #cad0dd;
  border-bottom: .6pt solid #e1e4ea;
}
.attrs li::marker { color: var(--blue); }

/* FIGURAS ----------------------------------------------------------------- */
.figure { margin: 6.5mm 0 8mm; }
.figure-frame {
  border-top: 0;
  border-bottom: 0;
  padding: 0;
  background: transparent;
}
.figure img { max-height: 150mm; }
.figure.hero img { max-height: 170mm; }
.figure figcaption {
  border-top: .45pt solid #d9dde6;
  padding-top: 2.2mm;
  margin-top: 2.5mm;
  font-family: var(--font-serif);
  font-size: 7.35pt;
  line-height: 1.38;
  color: #626975;
}
.figure figcaption strong {
  font-family: var(--font-sans);
  color: #1d327e;
  font-weight: 700;
}

/* TABLAS ------------------------------------------------------------------ */
.table-title {
  font-size: 7.5pt;
  color: #626975;
  margin: 5mm 0 2mm;
}
.table-title strong {
  color: #1d327e;
  font-weight: 700;
}
table {
  font-size: 7.2pt;
  line-height: 1.27;
}
th {
  background: #f0f2f8;
  color: #20252d;
  border-top: 1.2pt solid #2f50d1;
  border-bottom: .55pt solid #b9c1d5;
  padding: 2.2mm 2.4mm;
  font-weight: 700;
}
td {
  padding: 1.9mm 2.4mm;
  border-bottom: .45pt solid #e0e3e9;
}
tbody tr:nth-child(even) td { background: #fafbfc; }
.table-note { color: #777e89; font-size: 7.2pt; }

/* CAJAS / CITAS ----------------------------------------------------------- */
blockquote {
  background: #f7f8fb;
  border-left: 2.5px solid #c7cddd;
  color: #3d434c;
}
.keybox {
  background: #f3f5ff;
  border-left: 3px solid var(--blue);
}
.question {
  background: #fafafa;
  border-top: .6pt solid #cfd4df;
  border-bottom: .6pt solid #e2e5ea;
}
.chain {
  background: #f5f6fa;
  border-left: 0;
  border-top: .8pt solid #2f50d1;
  border-bottom: .45pt solid #d8dde8;
}

/* REFERENCIAS ------------------------------------------------------------- */
.part-references-heading,
.references-heading {
  border-top: .55pt solid #cfd4df !important;
}


/* V4.1 - discos sólidos (evita depender de box-shadow, no soportado por WeasyPrint) */
.cover::after,
.part-opener::after { display: none; }
.cover-title span { display: block; }
.cover-title { max-width: 112mm; font-size: 35pt; }
.cover-visual {
  position: absolute;
  z-index: 1;
  left: 109mm;
  top: 87mm;
  width: 66mm;
  height: 145mm;
}
.cover-disc {
  position: absolute;
  left: 0;
  width: 61mm;
  height: 61mm;
  border-radius: 50%;
}
.cover-disc.disc-1 { top: 0;    background: #173CF5; }
.cover-disc.disc-2 { top: 10mm; background: #294AF4; }
.cover-disc.disc-3 { top: 20mm; background: #3D5CF3; }
.cover-disc.disc-4 { top: 30mm; background: #526CF2; }
.cover-disc.disc-5 { top: 40mm; background: #6B80F2; }
.cover-disc.disc-6 { top: 50mm; background: #8797F3; }
.cover-disc.disc-7 { top: 60mm; background: #A7B2F5; }
.cover-disc.disc-8 { top: 70mm; background: #D1D8F9; }
/* orden visual: círculo superior intenso por encima de los más claros */
.cover-disc.disc-1 { z-index: 8; }
.cover-disc.disc-2 { z-index: 7; }
.cover-disc.disc-3 { z-index: 6; }
.cover-disc.disc-4 { z-index: 5; }
.cover-disc.disc-5 { z-index: 4; }
.cover-disc.disc-6 { z-index: 3; }
.cover-disc.disc-7 { z-index: 2; }
.cover-disc.disc-8 { z-index: 1; }

.part-visual {
  position: absolute;
  z-index: 1;
  right: -26mm;
  top: 60mm;
  width: 105mm;
  height: 112mm;
}
.part-disc {
  position: absolute;
  width: 88mm;
  height: 88mm;
  border-radius: 50%;
}
.part-disc-1 { right: 0;     top: 0;     background: #173CF5; z-index: 6; }
.part-disc-2 { right: 10mm;  top: 10mm;  background: #3557F2; z-index: 5; }
.part-disc-3 { right: 20mm;  top: 20mm;  background: #5A72F1; z-index: 4; }
.part-disc-4 { right: 30mm;  top: 30mm;  background: #8293F2; z-index: 3; }
.part-disc-5 { right: 40mm;  top: 40mm;  background: #AEB9F5; z-index: 2; }
.part-disc-6 { right: 50mm;  top: 50mm;  background: #D7DCFA; z-index: 1; }
.part-opener .part-content,
.part-opener .part-lower { position: relative; z-index: 3; }

/* El folio final se estampa en merge_and_finalize; no usar counter(page) por unidad. */
@page { @bottom-right { content: none; } }
.table-note { margin-top: 1.4mm; margin-bottom: 5mm; }


/* ========================================================================== */
/* V5 - SISTEMA HÍBRIDO MONOCROMÁTICO                                        */
/* Línea 1: resultados / Línea 2: técnico / Línea 3: conceptual              */
/* ========================================================================== */
:root {
  --ink: #121212;
  --ink-2: #333333;
  --muted: #666666;
  --paper: #fbfaf7;
  --paper-2: #f5f4f1;
  --gray-1: #eeeeec;
  --gray-2: #d8d8d5;
  --gray-3: #b8b8b5;
  --gray-4: #8e8e8b;
  --gray-5: #5e5e5c;
  --gray-6: #2f2f2e;
  --rule: #c9c9c6;
  --rule-soft: #e4e4e1;
  --blue: #222222;
  --blue-2: #444444;
  --blue-3: #666666;
  --blue-4: #888888;
  --blue-5: #aaaaaa;
  --blue-6: #dddddd;
  --teal: #2d2d2d;
  --teal-dark: #2d2d2d;
  --teal-deep: #181818;
  --orange: #777777;
  --graphite: #333333;
  --text: #252525;
  --cream: var(--paper);
  --title-dark: #111111;
  --secondary: #5d5d5b;
  --tertiary: #7a7a77;
  --line: var(--rule);
  --line-soft: var(--rule-soft);
  --teal-bg: #f4f4f2;
  --warm-bg: #f8f7f4;
  --font-serif: "Noto Serif", "DejaVu Serif", Georgia, serif;
  --font-sans: "Inter", "Noto Sans", Arial, sans-serif;
}

@page {
  margin: 17mm 17mm 18mm 17mm;
  @top-left { color: #777; font-size: 6.1pt; letter-spacing: .03em; }
  @top-right { color: #777; font-size: 6.1pt; }
  @bottom-left { color: #898986; font-size: 6pt; }
}

body { color: #262626; background: #fff; line-height: 1.49; }
a { color: #222; }
strong { color: #171717; }

/* PORTADA V5 -------------------------------------------------------------- */
.cover-v5 {
  background: var(--paper);
  color: var(--ink);
  overflow: hidden;
}
.cover-v5::before,
.cover-v5::after,
.cover-visual { display: none !important; }
.cover-v5 .cover-inner {
  position: relative;
  z-index: 5;
  height: 100%;
  padding: 19mm 22mm 19mm;
}
.cover-v5 .cover-header {
  margin: 0;
  padding: 0 0 4mm;
  border-bottom: .55pt solid #202020;
  color: #202020;
  font-size: 6.6pt;
  font-weight: 600;
  letter-spacing: .03em;
}
.cover-v5 .cover-title {
  margin: 13mm 0 0;
  max-width: 118mm;
  font-size: 38pt;
  line-height: .96;
  letter-spacing: -.042em;
  font-weight: 760;
  color: #101010;
}
.cover-v5 .cover-subtitle {
  margin-top: 6mm;
  padding: 0;
  font-size: 10.8pt;
  color: #30302f;
  font-weight: 500;
}
.cover-v5 .cover-subtitle::before { display:none; }
.cover-geometry {
  position: absolute;
  z-index: 2;
  left: 22mm;
  top: 105mm;
  width: 162mm;
  height: 126mm;
  overflow: visible;
}
.geo-bar {
  position: absolute;
  display: block;
  width: 112mm;
  height: 13mm;
  border-radius: 7mm;
  transform: rotate(-43deg);
  transform-origin: center;
}
.gb-1 { left: 3mm;  top: 37mm; background:#dadad7; }
.gb-2 { left: 31mm; top: 18mm; background:#bdbdb9; }
.gb-3 { left: 59mm; top: 2mm;  background:#7e7e7b; }
.gb-4 { left: 31mm; top: 58mm; background:#ededeb; }
.gb-5 { left: 59mm; top: 42mm; background:#a1a19e; }
.gb-6 { left: 87mm; top: 26mm; background:#555553; }
.gb-7 { left: 86mm; top: 68mm; background:#d1d1ce; }
.geo-node {
  position:absolute;
  width: 11.5mm;
  height: 11.5mm;
  border-radius:50%;
  background:#333331;
}
.gn-1 { left: 7mm;  top: 78mm; background:#b9b9b6; }
.gn-2 { left: 36mm; top: 59mm; background:#767673; }
.gn-3 { left: 64mm; top: 41mm; background:#4d4d4b; }
.gn-4 { left: 93mm; top: 23mm; background:#2b2b2a; }
.gn-5 { left: 37mm; top: 100mm; background:#8e8e8b; }
.gn-6 { left: 65mm; top: 82mm; background:#5c5c59; }
.gn-7 { left: 94mm; top: 64mm; background:#2d2d2c; }
.gn-8 { left: 122mm;top: 47mm; background:#151515; }
.gn-9 { left: 67mm; top: 119mm; background:#777774; }
.gn-10{ left: 95mm; top: 101mm; background:#3f3f3d; }
.gn-11{ left: 123mm;top: 84mm; background:#bdbdba; }
.cover-footer {
  position:absolute;
  z-index:6;
  left:22mm;
  right:22mm;
  bottom:18mm;
  border-top:.55pt solid #202020;
  padding-top:4mm;
  display:grid;
  grid-template-columns: 43mm 1fr;
  gap:7mm;
  font-family:var(--font-sans);
}
.cover-footer-keywords {
  border-right:.5pt solid #a8a8a5;
  padding-right:5mm;
  font-size:7pt;
  font-weight:700;
  line-height:1.5;
}
.cover-footer-keywords span,
.cover-footer-meta span,
.cover-footer-meta strong { display:block; }
.cover-footer-meta { font-size:7.1pt; line-height:1.4; color:#444; }
.cover-footer-meta strong { font-size:7.5pt; color:#202020; margin-bottom:.5mm; }
.cover-footer-meta span:last-child { color:#6d6d69; margin-top:1.2mm; }

/* RESUMEN E ÍNDICE -------------------------------------------------------- */
.executive-summary > h1,
.toc h1 {
  color:#111;
  border-top:.55pt solid #1f1f1f;
  font-size:21.5pt;
  letter-spacing:-.025em;
}
.executive-summary > h1::before,
.toc h1::before { background:#2d2d2c; width:3mm; height:3mm; }
.executive-summary strong { color:#222; }
.toc-part a { color:#222; }
.toc-num { color:#4b4b49; }

/* PORTADILLAS ------------------------------------------------------------- */
.part-opener {
  background:var(--paper);
  color:#111;
  padding:20mm 21mm 16mm;
  overflow:hidden;
}
.part-opener::before {
  left:21mm; right:21mm; top:20mm;
  border-top:.55pt solid #202020;
}
.part-opener .part-content { position:relative; z-index:5; }
.part-opener .part-kicker {
  margin:7mm 0 9mm;
  font-size:6.4pt;
  color:#555;
  letter-spacing:.05em;
}
.part-opener .part-number {
  color:#222;
  font-size:11pt;
  font-weight:650;
  margin-bottom:4mm;
}
.part-opener .part-name {
  color:#111;
  max-width:135mm;
  font-size:30pt;
  line-height:1.01;
  letter-spacing:-.033em;
  font-weight:720;
}
.part-opener.title-long .part-name { font-size:26pt; max-width:142mm; }
.part-opener.title-xlong .part-name { font-size:22pt; line-height:1.03; max-width:150mm; }
.part-visual,
.part-disc { display:none !important; }
.part-visual-v5 {
  position:absolute;
  z-index:2;
  right:-20mm;
  top:76mm;
  width:130mm;
  height:105mm;
}
.pbar {
  position:absolute;
  width:96mm;
  height:15mm;
  border-radius:8mm;
  transform:rotate(-43deg);
}
.pbar-1 { left:0; top:15mm; background:#dededb; }
.pbar-2 { left:25mm; top:0; background:#bdbdba; }
.pbar-3 { left:22mm; top:43mm; background:#a0a09d; }
.pbar-4 { left:50mm; top:28mm; background:#5c5c59; }
.pbar-5 { left:48mm; top:67mm; background:#d0d0cd; }
.pnode { position:absolute; width:15mm; height:15mm; border-radius:50%; }
.pnode-1 { left:7mm; top:59mm; background:#282827; }
.pnode-2 { left:39mm; top:39mm; background:#777774; }
.pnode-3 { left:70mm; top:20mm; background:#333331; }
.pnode-4 { left:40mm; top:83mm; background:#555553; }
.pnode-5 { left:72mm; top:63mm; background:#232322; }
.pnode-6 { left:103mm;top:44mm; background:#aaa9a6; }
.part-opener .part-lower {
  position:absolute;
  z-index:6;
  left:21mm;
  right:21mm;
  top:169mm;
}
.part-description { max-width:105mm; font-size:8.8pt; line-height:1.42; color:#494947; }
.part-mini-index {
  border-top:.55pt solid #202020;
  padding-top:2.5mm;
  gap:0 9mm;
}
.part-mini-index a {
  min-height:11.5mm;
  padding:2.5mm 0 2.8mm;
  border-bottom:.45pt solid #dededb;
  font-size:7.8pt;
  line-height:1.22;
  color:#444442;
}
.part-mini-index .mini-num { color:#2e2e2d; font-size:9.5pt; }
.part-mini-index.index-few { grid-template-columns:1fr 1fr; }
.part-mini-index.index-few a { min-height:18mm; font-size:9.2pt; }
.part-mini-index.index-many {
  grid-template-columns:repeat(3, 1fr);
  gap:0 6mm;
}
.part-mini-index.index-many a {
  grid-template-columns:8mm 1fr;
  min-height:8.3mm;
  padding:1.6mm 0 1.8mm;
  font-size:6.65pt;
  line-height:1.15;
}
.part-mini-index.index-many .mini-num { font-size:7.5pt; }

/* Variación de portadilla por línea */
.part-opener.part-style-1 .part-visual-v5 { top:82mm; right:-12mm; opacity:.9; }
.part-opener.part-style-1 .part-name { max-width:122mm; }
.part-opener.part-style-2 .part-visual-v5 { transform:scale(.82); transform-origin:top right; top:80mm; right:-8mm; }
.part-opener.part-style-2 .part-lower { top:151mm; }
.part-opener.part-style-2 .part-mini-index { background:#f6f6f4; padding:3mm 4mm 0; }
.part-opener.part-style-3 .part-visual-v5 { top:71mm; right:-22mm; transform:scale(1.05); transform-origin:top right; }

/* BASE INTERIOR ----------------------------------------------------------- */
.flow-columns { column-gap:9mm; }
h2, h3, h4 { color:#151515; }
h2::before { background:#333332; }
h3 { color:#333331; }
li::marker { color:#555553; }
.figure-frame { background:transparent; }
.figure figcaption strong,
.table-title strong { color:#292928; }
.figure figcaption { border-top:.45pt solid #d7d7d4; color:#646461; }
blockquote { background:#f5f5f3; border-left:2.5px solid #bdbdba; }
.keybox { background:#f3f3f1; border-left:3px solid #444442; }
.chain { background:#f5f5f3; border-top:.8pt solid #444442; }
.attrs { background:#f5f5f3; }

/* LÍNEA 1 - equilibrada / científica / resultados ------------------------ */
.part-body.part-style-1 h2 {
  font-size:13.6pt;
  border-top:.55pt solid #bdbdba;
  padding-top:3mm;
  margin-bottom:4mm;
}
.part-body.part-style-1 h2::before { width:2.6mm; height:2.6mm; margin-right:2.2mm; }
.part-body.part-style-1 h3 { font-size:10.7pt; margin-top:4.7mm; }
.part-body.part-style-1 .figure.hero { margin:7mm 0 9mm; }
.part-body.part-style-1 .figure.hero img { max-height:172mm; }
.part-body.part-style-1 table { font-size:7.25pt; }
.part-body.part-style-1 th {
  background:#f0f0ee;
  color:#202020;
  border-top:1.0pt solid #3b3b3a;
  border-bottom:.55pt solid #bdbdba;
}
.part-body.part-style-1 tbody tr:nth-child(even) td { background:#fafaf9; }

/* LÍNEA 2 - técnica / analítica / metodología ----------------------------- */
.part-body.part-style-2 { font-size:8.95pt; line-height:1.46; }
.part-body.part-style-2 .flow-columns { column-gap:8.5mm; }
.part-body.part-style-2 h2 {
  font-size:12.6pt;
  border-top:1.1pt solid #3c3c3a;
  border-bottom:.45pt solid #d0d0cd;
  padding:2.6mm 0 1.5mm;
  margin-bottom:3.5mm;
}
.part-body.part-style-2 h2::before { display:none; }
.part-body.part-style-2 h3 {
  font-size:10.15pt;
  margin:4.3mm 0 2mm;
  padding-left:2.3mm;
  border-left:1.7pt solid #666663;
}
.part-body.part-style-2 h4 { font-size:9.1pt; }
.part-body.part-style-2 table { font-size:7.05pt; line-height:1.23; }
.part-body.part-style-2 th {
  background:#e7e7e4;
  color:#181818;
  border-top:1.15pt solid #313130;
  border-bottom:.65pt solid #9d9d99;
  padding:2mm 2.2mm;
}
.part-body.part-style-2 td { padding:1.75mm 2.2mm; }
.part-body.part-style-2 tbody tr:nth-child(even) td { background:#f4f4f2; }
.part-body.part-style-2 .table-title { margin-top:4.3mm; }
.part-body.part-style-2 blockquote { font-size:8.55pt; }

/* LÍNEA 3 - conceptual / editorial / mapas y síntesis -------------------- */
.part-body.part-style-3 { line-height:1.52; }
.part-body.part-style-3 h2 {
  font-size:14.8pt;
  line-height:1.08;
  letter-spacing:-.022em;
  border-top:.55pt solid #202020;
  padding-top:3.8mm;
  margin-bottom:5mm;
}
.part-body.part-style-3 h2::before {
  width:3.2mm; height:3.2mm; background:#30302f; vertical-align:1.1mm;
}
.part-body.part-style-3 h3 {
  font-size:11.1pt;
  margin:5.4mm 0 2.5mm;
  color:#292928;
}
.part-body.part-style-3 .figure { margin:7.5mm 0 9mm; }
.part-body.part-style-3 .figure.hero { margin:9mm 0 11mm; }
.part-body.part-style-3 .figure figcaption { padding-top:2.6mm; }
.part-body.part-style-3 blockquote.keybox,
.part-body.part-style-3 blockquote.chain {
  padding:5mm 6mm;
  margin:6mm 0;
}
.part-body.part-style-3 table th {
  background:#f2f2f0;
  color:#202020;
  border-top:.9pt solid #4c4c4a;
  border-bottom:.55pt solid #bfbfbc;
}

/* FIGURAS Y TABLAS MONOCROMAS -------------------------------------------- */
.figure img { filter: grayscale(100%); }
th { background:#efefed; color:#202020; border-top:1pt solid #3d3d3b; border-bottom:.55pt solid #bdbdba; }
tbody tr:nth-child(even) td { background:#f8f8f6; }
.table-note { color:#747471; }

/* Reduce riesgo de cortes en portadillas y encabezados */
.part-opener, .cover { break-inside:avoid; }
.part-opener .part-mini-index a { break-inside:avoid; }
h2, h3, h4 { break-after:avoid; }

"""


# =============================================================================
# UTILIDADES DE ARCHIVOS
# =============================================================================

def safe_extract_zip(zip_path: Path, dest: Path) -> List[Path]:
    extracted: List[Path] = []
    dest_resolved = dest.resolve()
    with zipfile.ZipFile(zip_path, "r") as zf:
        for info in zf.infolist():
            name = info.filename.replace("\\", "/")
            if name.endswith("/"):
                continue
            target = (dest / name).resolve()
            if os.path.commonpath([str(dest_resolved), str(target)]) != str(dest_resolved):
                raise ValueError(f"Ruta insegura dentro del ZIP: {info.filename}")
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info, "r") as src, open(target, "wb") as dst:
                shutil.copyfileobj(src, dst)
            extracted.append(target)
    return extracted


def collect_images(source: Path, dest: Path) -> Dict[str, Path]:
    """Copia/extracta imágenes y devuelve índice por basename en minúsculas."""
    dest.mkdir(parents=True, exist_ok=True)
    if source.is_dir():
        paths = [p for p in source.rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED_IMAGES]
    elif source.suffix.lower() == ".zip":
        paths = safe_extract_zip(source, dest / "_zip")
        paths = [p for p in paths if p.suffix.lower() in SUPPORTED_IMAGES]
    else:
        raise ValueError("--figuras debe ser una carpeta o un archivo .zip")

    index: Dict[str, Path] = {}
    duplicates: Dict[str, List[Path]] = {}
    figures_dir = dest / "figures"
    figures_dir.mkdir(parents=True, exist_ok=True)

    for p in paths:
        key = p.name.lower()
        if key in index:
            duplicates.setdefault(key, [index[key]]).append(p)
            continue
        target = figures_dir / p.name
        shutil.copy2(p, target)
        index[key] = target

    if duplicates:
        detail = "\n".join(f"  - {k}: " + ", ".join(str(x) for x in v) for k, v in duplicates.items())
        raise RuntimeError("Hay imágenes con basename duplicado; la referencia Markdown sería ambigua:\n" + detail)

    return index


def auto_pick_report_md(cwd: Path) -> Path:
    candidates = sorted(cwd.glob("*.md"))
    preferred = [p for p in candidates if "reporte_hito0" in p.name.lower() and "diseno" not in p.name.lower()]
    if len(preferred) == 1:
        return preferred[0]
    if len(candidates) == 1:
        return candidates[0]
    if not candidates:
        raise FileNotFoundError("No se encontró ningún .md en la carpeta actual")
    names = "\n  - " + "\n  - ".join(p.name for p in candidates)
    raise RuntimeError("Hay más de un .md. Indica --md explícitamente:" + names)


def auto_pick_figures(cwd: Path) -> Path:
    folder = cwd / "01_figuras"
    if folder.exists() and folder.is_dir():
        return folder
    zips = sorted(cwd.glob("*.zip"))
    preferred = [p for p in zips if "figura" in p.name.lower()]
    if len(preferred) == 1:
        return preferred[0]
    if len(zips) == 1:
        return zips[0]
    if not zips:
        raise FileNotFoundError("No se encontró 01_figuras ni un ZIP de figuras")
    names = "\n  - " + "\n  - ".join(p.name for p in zips)
    raise RuntimeError("Hay más de un ZIP. Indica --figuras explícitamente:" + names)


# =============================================================================
# MATEMÁTICA SIMPLE
# =============================================================================

def _replace_frac(expr: str) -> str:
    pattern = re.compile(r"\\frac\{([^{}]+)\}\{([^{}]+)\}")
    prev = None
    while prev != expr:
        prev = expr
        expr = pattern.sub(r"(\1)/(\2)", expr)
    return expr


def _replace_sqrt(expr: str) -> str:
    pattern = re.compile(r"\\sqrt\{((?:[^{}]|\{[^{}]*\})+)\}")
    prev = None
    while prev != expr:
        prev = expr
        expr = pattern.sub(r"√(\1)", expr)
    return expr


def latex_to_html(expr: str, css_class: str = "math-inline") -> str:
    expr = expr.strip()
    expr = _replace_frac(expr)
    expr = _replace_sqrt(expr)
    symbols = {
        r"\theta": "θ", r"\tau": "τ", r"\sigma": "σ", r"\rho": "ρ",
        r"\alpha": "α", r"\beta": "β", r"\omega": "ω", r"\chi": "χ",
        r"\geq": "≥", r"\leq": "≤", r"\neq": "≠", r"\approx": "≈",
        r"\times": "×", r"\cdot": "·", r"\pm": "±",
        r"\leftrightarrow": "↔", r"\rightarrow": "→", r"\to": "→",
    }
    for old, new in symbols.items():
        expr = expr.replace(old, new)
    expr = expr.replace(r"\,", " ").replace(r"\;", " ").replace(r"\:", " ").replace(r"\!", "")
    expr = expr.replace("'", "′")
    expr = html_lib.escape(expr, quote=False)
    expr = re.sub(r"_\{([^{}]+)\}", r"<sub>\1</sub>", expr)
    expr = re.sub(r"\^\{([^{}]+)\}", r"<sup>\1</sup>", expr)
    expr = re.sub(r"_([A-Za-z0-9]+)", r"<sub>\1</sub>", expr)
    expr = re.sub(r"\^([A-Za-z0-9+\-]+)", r"<sup>\1</sup>", expr)
    expr = expr.replace("{", "").replace("}", "")
    return f'<span class="{css_class}">{expr}</span>'


def preprocess_math(md_text: str) -> str:
    """Convierte la notación matemática simple usada en el reporte a HTML seguro."""
    # Display math: \[ ... \]
    def display_repl(m: re.Match) -> str:
        inner = m.group(1).strip().replace("\n", " ")
        rendered = latex_to_html(inner, "math-display-inner")
        return f"\n<div class=\"equation\">{rendered}</div>\n"

    md_text = re.sub(r"\\\[\s*(.*?)\s*\\\]", display_repl, md_text, flags=re.S)

    placeholders: Dict[str, str] = {}
    counter = 0

    def stash(expr: str) -> str:
        nonlocal counter
        token = f"HITO0MATHINLINE{counter}TOKEN"
        placeholders[token] = latex_to_html(expr)
        counter += 1
        return token

    md_text = re.sub(r"\\\((.+?)\\\)", lambda m: stash(m.group(1)), md_text)
    md_text = re.sub(r"(?<!\\)\$([^$\n]+?)(?<!\\)\$", lambda m: stash(m.group(1)), md_text)

    # Se restauran después del render Markdown.
    md_text += "\n<!--HITO0_MATH_PLACEHOLDERS:" + repr(placeholders) + "-->\n"
    return md_text


def restore_math_placeholders(html_text: str, md_with_marker: str) -> str:
    m = re.search(r"<!--HITO0_MATH_PLACEHOLDERS:(\{.*?\})-->", md_with_marker, flags=re.S)
    if not m:
        return html_text
    import ast
    placeholders = ast.literal_eval(m.group(1))
    for token, rendered in placeholders.items():
        html_text = html_text.replace(token, rendered)
    return html_text


# =============================================================================
# MARKDOWN -> HTML
# =============================================================================

def strip_html_comments(text: str) -> str:
    return re.sub(r"<!--.*?-->", "", text, flags=re.S)


def extract_cover_fields(text: str) -> Tuple[str, str, str]:
    title_match = re.search(r"^#\s+(.+?)\s*$", text, flags=re.M)
    title = title_match.group(1).strip() if title_match else "Reporte Institucional del Hito 0"
    institution_match = re.search(r"^\*\*(Facultad de Ingeniería.*?Chile)\*\*\s*$", text, flags=re.M)
    institution = institution_match.group(1).strip() if institution_match else "Facultad de Ingeniería, Universidad de Santiago de Chile"

    body = text
    if title_match:
        body = body[:title_match.start()] + body[title_match.end():]
    if institution_match:
        # Rebuscar tras quitar título, porque offsets originales ya no sirven.
        body = re.sub(r"^\*\*Facultad de Ingeniería.*?Chile\*\*\s*$", "", body, count=1, flags=re.M)
    return title, institution, body.strip()


def slugify(text: str) -> str:
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = text.lower()
    text = re.sub(r"[^a-z0-9]+", "-", text).strip("-")
    return text or "seccion"


def make_markdown_html(md_text: str) -> str:
    prepared = preprocess_math(md_text)
    md = MarkdownIt("commonmark", {"html": True, "breaks": False})
    md.enable("table")
    html_text = md.render(prepared)
    html_text = restore_math_placeholders(html_text, prepared)
    # El comentario auxiliar de placeholders no debe llegar al DOM.
    html_text = re.sub(r"<!--HITO0_MATH_PLACEHOLDERS:.*?-->", "", html_text, flags=re.S)
    return html_text


# =============================================================================
# TRANSFORMACIÓN EDITORIAL DEL DOM
# =============================================================================

def add_heading_ids(root: Tag) -> List[Tuple[int, str, str]]:
    used: Dict[str, int] = {}
    headings: List[Tuple[int, str, str]] = []
    for h in root.find_all(["h1", "h2", "h3", "h4"]):
        txt = h.get_text(" ", strip=True)
        base = slugify(txt)
        used[base] = used.get(base, 0) + 1
        ident = base if used[base] == 1 else f"{base}-{used[base]}"
        h["id"] = ident
        headings.append((int(h.name[1]), txt, ident))
    return headings


def next_tag_sibling(node: Tag) -> Optional[Tag]:
    sib = node.next_sibling
    while sib is not None:
        if isinstance(sib, Tag):
            return sib
        if isinstance(sib, NavigableString) and sib.strip():
            return None
        sib = sib.next_sibling
    return None


def image_ratio(path: Path) -> Optional[float]:
    try:
        if path.suffix.lower() == ".svg":
            return None
        with PILImage.open(path) as im:
            w, h = im.size
            return w / h if h else None
    except Exception:
        return None


def transform_figures(soup: BeautifulSoup, root: Tag, image_index: Dict[str, Path]) -> Tuple[List[str], List[str]]:
    used_images: List[str] = []
    warnings: List[str] = []
    for p in list(root.find_all("p")):
        img = p.find("img", recursive=False)
        if not img:
            continue
        src = img.get("src", "")
        basename = Path(src.replace("\\", "/")).name
        key = basename.lower()
        target = image_index.get(key)
        if not target:
            warnings.append(f"Figura referenciada pero no encontrada: {src}")
            continue

        used_images.append(basename)
        img["src"] = target.resolve().as_uri()

        caption = next_tag_sibling(p)
        if caption and caption.name == "p" and re.match(r"^Figura\s+\d+\.", caption.get_text(" ", strip=True), flags=re.I):
            caption.extract()
        else:
            caption = None

        figure = soup.new_tag("figure")
        classes = ["figure"]
        ratio = image_ratio(target)
        lowname = basename.lower()
        alt = img.get("alt", "").lower()

        # Panorámicas extremas: página horizontal para preservar legibilidad.
        if ratio is not None and ratio >= 4.0:
            classes.append("landscape")
        # Mapas de constructo, figuras multipanel o densas: hero.
        elif ("-cm-" in lowname or "construct map" in alt or (ratio is not None and ratio >= 2.15)):
            classes.append("hero")

        figure["class"] = classes
        frame = soup.new_tag("div", attrs={"class": "figure-frame"})
        p.unwrap()  # deja img en el lugar temporal
        img.extract()
        frame.append(img)
        figure.append(frame)
        if caption:
            fc = soup.new_tag("figcaption")
            for child in list(caption.contents):
                fc.append(child.extract())
            figure.append(fc)
        # El p fue desempaquetado; insertar figure antes de la posición de la imagen ya extraída.
        # Como la imagen quedó sin padre, usamos el antiguo punto: reemplazamos el nodo p ya no existe,
        # así que agregamos figure donde estaba mediante referencia al siguiente sibling si existe.
        # La forma más estable es usar insert_before antes de unwrap; por eso esta rama se reconstruye
        # colocando la figura en el final si se perdió referencia. Se corrige con marcador previo.
        marker = soup.new_tag("span", attrs={"class": "_figure_marker"})
        # p ya fue desempaquetado, por lo que insertamos usando el primer nodo siguiente cuando exista.
        # Para evitar cambios de orden, la transformación principal usa una segunda pasada más abajo.
        frame["data-source-name"] = basename

    # La primera pasada anterior no puede conservar posición de forma robusta tras unwrap.
    # Rehacemos figuras desde las imágenes aún sueltas mediante un procedimiento posicional.
    # Si ya quedaron dentro de figure-frame, no tocar.
    for frame in root.find_all("div", class_="figure-frame"):
        if frame.parent and frame.parent.name == "figure":
            continue

    # El algoritmo anterior se reemplaza realmente por el posicional de transform_figures_positional.
    return used_images, warnings


def transform_figures_positional(soup: BeautifulSoup, root: Tag, image_index: Dict[str, Path]) -> Tuple[List[str], List[str]]:
    """Agrupa imagen + caption conservando exactamente la posición del Markdown."""
    used_images: List[str] = []
    warnings: List[str] = []
    for p in list(root.find_all("p")):
        if not p.parent or p.find_parent("figure"):
            continue
        img = p.find("img", recursive=False)
        if not img:
            continue
        src = img.get("src", "")
        basename = Path(src.replace("\\", "/")).name
        target = image_index.get(basename.lower())
        if not target:
            warnings.append(f"Figura referenciada pero no encontrada: {src}")
            continue
        used_images.append(basename)
        img["src"] = target.resolve().as_uri()

        caption = next_tag_sibling(p)
        if caption and caption.name == "p" and re.match(r"^Figura\s+\d+\.", caption.get_text(" ", strip=True), flags=re.I):
            caption.extract()
        else:
            caption = None

        figure = soup.new_tag("figure")
        classes = ["figure"]
        ratio = image_ratio(target)
        lowname = basename.lower()
        alt = img.get("alt", "").lower()
        slug_match = re.match(r"figura[-_]\d+[-_](.+?)\.\w+$", lowname)
        fig_slug = slug_match.group(1) if slug_match else None

        if fig_slug in FORCE_HERO_FIGURE_SLUGS:
            classes.append("hero")
        elif ratio is not None and ratio >= 4.0:
            classes.append("landscape")
        elif ("-cm-" in lowname or "construct map" in alt or (ratio is not None and ratio >= 2.15)):
            classes.append("hero")
        figure["class"] = classes

        if fig_slug in COMPACT_FIGURE_MAX_HEIGHT_MM:
            img["style"] = f"max-height:{COMPACT_FIGURE_MAX_HEIGHT_MM[fig_slug]:.1f}mm"

        frame = soup.new_tag("div", attrs={"class": "figure-frame"})
        img.extract()
        frame.append(img)
        figure.append(frame)
        if caption:
            fc = soup.new_tag("figcaption")
            for child in list(caption.contents):
                fc.append(child.extract())
            figure.append(fc)
        p.replace_with(figure)
    return used_images, warnings


def transform_tables(soup: BeautifulSoup, root: Tag) -> None:
    """Clasifica tablas por densidad y agrupa título + tabla + nota.

    Las tablas pequeñas permanecen dentro de UNA columna narrativa. Las tablas
    anchas o densas se extraen al ancho completo durante la normalización del flujo.
    """
    for p in root.find_all("p"):
        txt = p.get_text(" ", strip=True)
        if re.match(r"^Tabla\s+\d+\.", txt, flags=re.I):
            p["class"] = list(dict.fromkeys(p.get("class", []) + ["table-title"]))
        elif re.match(r"^Panel\s+[A-Z0-9]+\.", txt, flags=re.I):
            p["class"] = list(dict.fromkeys(p.get("class", []) + ["panel-label"]))
        elif re.match(r"^(Nota\.|Nota:)", txt, flags=re.I):
            p["class"] = list(dict.fromkeys(p.get("class", []) + ["table-note"]))

    for table in list(root.find_all("table")):
        rows = table.find_all("tr")
        first_row = rows[0] if rows else None
        cols = len(first_row.find_all(["th", "td"], recursive=False)) if first_row else 0
        data_rows = max(0, len(rows) - 1)
        cell_texts = [c.get_text(" ", strip=True) for c in table.find_all(["th", "td"])]
        max_cell = max((len(t) for t in cell_texts), default=0)
        total_chars = sum(len(t) for t in cell_texts)

        # Heurística editorial: una tabla es realmente pequeña si, además de
        # pocas columnas/filas, sus celdas no contienen párrafos largos.
        small = (
            (cols <= 3 and data_rows <= 14 and max_cell <= 90 and total_chars <= 1800)
            or (cols <= 4 and data_rows <= 7 and max_cell <= 55 and total_chars <= 1200)
        )

        classes = list(table.get("class", []))
        if cols >= 7:
            classes.append("matrix")
        elif cols >= 5:
            classes.append("medium")
        table["class"] = list(dict.fromkeys(classes))

        # Capturar título/panel inmediatamente anteriores y nota posterior.
        before = []
        prev = table.find_previous_sibling()
        if isinstance(prev, Tag) and "panel-label" in prev.get("class", []):
            before.insert(0, prev)
            prev2 = prev.find_previous_sibling()
            if isinstance(prev2, Tag) and "table-title" in prev2.get("class", []):
                before.insert(0, prev2)
        elif isinstance(prev, Tag) and "table-title" in prev.get("class", []):
            before.append(prev)

        table_title_text = ""
        for node in before:
            if "table-title" in node.get("class", []):
                table_title_text = _normalize_search_text(node.get_text(" ", strip=True))
                break
        if any(t in table_title_text for t in FORCE_WIDE_TABLE_TITLES):
            small = False

        after = table.find_next_sibling()
        note = after if isinstance(after, Tag) and "table-note" in after.get("class", []) else None

        group = soup.new_tag(
            "div",
            attrs={"class": f"table-group {'small-table' if small else 'wide-table'}"}
        )
        anchor = before[0] if before else table
        anchor.insert_before(group)

        for node in before:
            group.append(node.extract())

        wrapper = soup.new_tag("div", attrs={"class": "table-block"})
        table.extract()
        wrapper.append(table)
        group.append(wrapper)

        if note is not None and note.parent:
            group.append(note.extract())


def transform_blockquotes(root: Tag) -> None:
    for bq in root.find_all("blockquote"):
        txt = bq.get_text(" ", strip=True)
        classes = bq.get("class", [])
        if re.match(r"^Nota\b", txt, flags=re.I):
            classes.append("note-box")
        elif len(txt) <= 380 and ("→" in txt or "↔" in txt):
            classes.append("chain")
        elif len(txt) <= 420 and (txt.startswith("¿") or ("?" in txt and len(txt.split()) <= 45)):
            classes.append("question")
        elif len(txt) <= 500 and bq.find("strong"):
            classes.append("keybox")
        bq["class"] = list(dict.fromkeys(classes))


def transform_attribute_list(root: Tag) -> None:
    for ol in root.find_all("ol"):
        items = ol.find_all("li", recursive=False)
        if len(items) == 10:
            joined = " ".join(li.get_text(" ", strip=True).lower() for li in items)
            if "comunicación" in joined and "liderazgo" in joined and "diseño" in joined:
                ol["class"] = list(set(ol.get("class", []) + ["attrs"]))


def split_part_heading(text: str) -> Tuple[str, str]:
    m = re.match(r"^(PARTE\s+[IVXLCDM]+)\s*(?:—|–|-|\.)?\s*(.*)$", text.strip(), flags=re.I)
    if m:
        return m.group(1).upper(), m.group(2).strip()
    return text.strip(), ""


def split_section_num(text: str) -> Tuple[str, str]:
    m = re.match(r"^(\d+(?:\.\d+)*\.?)(?:\s+)(.*)$", text.strip())
    if m:
        return m.group(1), m.group(2)
    return "", text.strip()


def top_level_segments(root: Tag) -> List[Tuple[Optional[Tag], List[Tag]]]:
    segments: List[Tuple[Optional[Tag], List[Tag]]] = []
    current_h1: Optional[Tag] = None
    nodes: List[Tag] = []
    for child in list(root.children):
        if not isinstance(child, Tag):
            continue
        if child.name == "h1":
            if current_h1 is not None or nodes:
                segments.append((current_h1, nodes))
            current_h1 = child
            nodes = []
        else:
            nodes.append(child)
    if current_h1 is not None or nodes:
        segments.append((current_h1, nodes))
    return segments


def build_toc(soup: BeautifulSoup, headings: List[Tuple[int, str, str]]) -> Tag:
    section = soup.new_tag("section", attrs={"class": "toc"})
    h = soup.new_tag("h1")
    h.string = "Índice general"
    section.append(h)
    listing = soup.new_tag("div", attrs={"class": "toc-list"})
    section.append(listing)

    for level, text, ident in headings:
        if level == 1 and (text.upper().startswith("PARTE ") or text.upper() == "REFERENCIAS"):
            div = soup.new_tag("div", attrs={"class": "toc-part"})
        elif level == 2:
            div = soup.new_tag("div", attrs={"class": "toc-sec"})
        else:
            continue
        a = soup.new_tag("a", href=f"#{ident}")
        num, label = split_section_num(text)
        if text.upper().startswith("PARTE "):
            num, label = split_part_heading(text)
        nspan = soup.new_tag("span", attrs={"class": "toc-num"})
        nspan.string = num
        tspan = soup.new_tag("span", attrs={"class": "toc-text"})
        tspan.string = label if label else text
        pspan = soup.new_tag("span", attrs={"class": "toc-page", "data-href": f"#{ident}"})
        a.extend([nspan, tspan, pspan])
        div.append(a)
        listing.append(div)
    return section


def build_cover(soup: BeautifulSoup, title: str, institution: str) -> Tag:
    """Portada V5: sistema monocromático de barras diagonales y nodos."""
    section = soup.new_tag("section", attrs={"class": "cover cover-v5"})
    inner = soup.new_tag("div", attrs={"class": "cover-inner"})

    header = soup.new_tag("div", attrs={"class": "cover-header"})
    left = soup.new_tag("span"); left.string = "INFORME DE MEDICIÓN PSICOMÉTRICA"
    right = soup.new_tag("span"); right.string = "ESTUDIO 2026"
    header.extend([left, right])

    h = soup.new_tag("h1", attrs={"class": "cover-title"})
    line1 = soup.new_tag("span"); line1.string = "Diagnostico"
    line2 = soup.new_tag("span"); line2.string = "Hito 0"
    h.extend([line1, line2])

    subtitle = soup.new_tag("div", attrs={"class": "cover-subtitle"})
    subtitle.string = "Medición Atributos i+e"

    visual = soup.new_tag("div", attrs={"class": "cover-geometry"})
    for i in range(1, 8):
        visual.append(soup.new_tag("span", attrs={"class": f"geo-bar gb-{i}"}))
    for i in range(1, 12):
        visual.append(soup.new_tag("span", attrs={"class": f"geo-node gn-{i}"}))

    footer = soup.new_tag("div", attrs={"class": "cover-footer"})
    leftblock = soup.new_tag("div", attrs={"class": "cover-footer-keywords"})
    for word in ["DIAGNÓSTICO", "EVIDENCIA", "VALIDACIÓN"]:
        sp = soup.new_tag("span"); sp.string = word; leftblock.append(sp)
    rightblock = soup.new_tag("div", attrs={"class": "cover-footer-meta"})
    faculty = soup.new_tag("strong")
    if "," in institution:
        f, rest = institution.split(",", 1)
    else:
        f, rest = institution, ""
    faculty.string = f.strip()
    rightblock.append(faculty)
    if rest.strip():
        inst = soup.new_tag("span"); inst.string = rest.strip(); rightblock.append(inst)
    desc = soup.new_tag("span"); desc.string = "Perfil de ingreso · atributos de innovación y emprendimiento"; rightblock.append(desc)
    footer.extend([leftblock, rightblock])

    inner.extend([header, h, subtitle, footer])
    section.extend([visual, inner])
    return section


def part_style_for_heading(text: str) -> str:
    """Mapea Partes a las tres líneas editoriales aprobadas.

    Línea 1: equilibrada/científica (Parte II)
    Línea 2: técnica/analítica (Partes IV y VI)
    Línea 3: conceptual/editorial (Partes I, III, V y VII)
    """
    m = re.match(r"^PARTE\s+([IVXLCDM]+)", text.strip(), flags=re.I)
    roman = m.group(1).upper() if m else ""
    return {
        "I": "3", "II": "1", "III": "3", "IV": "2",
        "V": "3", "VI": "2", "VII": "3",
    }.get(roman, "1")

def build_part_opener(soup: BeautifulSoup, h1: Tag, nodes: List[Tag]) -> Tuple[Tag, List[Tag]]:
    text = h1.get_text(" ", strip=True)
    number, name = split_part_heading(text)
    style_id = part_style_for_heading(text)
    name_len = len(name)
    length_class = "title-short" if name_len < 35 else ("title-long" if name_len < 62 else "title-xlong")
    opener = soup.new_tag("section", attrs={"class": f"part-opener part-style-{style_id} {length_class}"})
    content = soup.new_tag("div", attrs={"class": "part-content"})
    kicker = soup.new_tag("div", attrs={"class": "part-kicker"})
    kicker.string = "DIAGNOSTICO HITO 0 · MEDICIÓN ATRIBUTOS i+e"
    heading = soup.new_tag("h1", attrs={"class": "part-heading", "id": h1.get("id", slugify(text))})
    ns = soup.new_tag("span", attrs={"class": "part-number"}); ns.string = number
    nm = soup.new_tag("span", attrs={"class": "part-name"}); nm.string = name
    heading.extend([ns, nm]); content.extend([kicker, heading]); opener.append(content)

    visual = soup.new_tag("div", attrs={"class": "part-visual-v5"})
    for i in range(1, 6):
        visual.append(soup.new_tag("span", attrs={"class": f"pbar pbar-{i}"}))
    for i in range(1, 7):
        visual.append(soup.new_tag("span", attrs={"class": f"pnode pnode-{i}"}))
    opener.append(visual)

    lower = soup.new_tag("div", attrs={"class": "part-lower"})

    # Párrafos y notas introductorias previas al primer H2 pertenecen a la
    # portadilla. Esto evita crear una falsa "sección" adicional (caso Parte III).
    desc_nodes: List[Tag] = []
    remaining = list(nodes)
    current_chars = 0
    while remaining:
        n = remaining[0]
        if n.name == "h2":
            break
        if n.name in {"p", "blockquote"} and current_chars < 1500:
            item = remaining.pop(0)
            desc_nodes.append(item)
            current_chars += len(item.get_text(" ", strip=True))
        elif n.name == "hr":
            remaining.pop(0)
        else:
            break
    if desc_nodes:
        desc = soup.new_tag("div", attrs={"class": "part-description"})
        for n in desc_nodes: desc.append(n)
        lower.append(desc)

    h2s = [n for n in remaining if n.name == "h2"]
    mini_classes = ["part-mini-index"]
    if len(h2s) <= 4:
        mini_classes.append("index-few")
    elif len(h2s) >= 9:
        mini_classes.append("index-many")
    mini = soup.new_tag("div", attrs={"class": " ".join(mini_classes)})
    for n in h2s:
        txt = n.get_text(" ", strip=True)
        num, label = split_section_num(txt)
        a = soup.new_tag("a", href=f"#{n.get('id')}")
        s_num = soup.new_tag("span", attrs={"class": "mini-num"}); s_num.string = num
        s_label = soup.new_tag("span", attrs={"class": "mini-label"}); s_label.string = label
        a.extend([s_num, s_label]); mini.append(a)
    if mini.contents: lower.append(mini)
    opener.append(lower)
    return opener, remaining



def normalize_part_column_flows(soup: BeautifulSoup, container: Tag) -> None:
    """Convierte el contenido de una Parte en un flujo secuencial estricto.

    Regla invariable de composición:
        columna izquierda -> columna derecha -> página siguiente.

    Los títulos H2/H3/H4, listas, párrafos y cambios de sección permanecen en el
    MISMO tramo multicolumna. Solo los elementos que realmente requieren ancho
    completo interrumpen temporalmente el tramo. Antes y después de ellos se
    crea un nuevo ``flow-columns`` con ``column-fill:auto``; nunca se balancea.
    """
    direct = [x for x in list(container.children) if isinstance(x, Tag)]
    if not direct:
        return

    def is_full_width(node: Tag) -> bool:
        classes = set(node.get("class", []))

        # Figuras normales pueden vivir en una columna. Solo las densas/hero y
        # las que requieren página especial salen del flujo de columna.
        if node.name == "figure":
            return bool(classes.intersection({"hero", "landscape", "fullpage"}))

        # Tablas pequeñas permanecen en una columna; matrices/tablas anchas no.
        if node.name == "div" and "table-group" in classes:
            return "wide-table" in classes

        # Elementos conceptuales cuyo diseño depende realmente del ancho total.
        # Las ecuaciones NO están aquí: permanecen en una columna.
        if classes.intersection({"chain", "question", "attrs"}):
            return True
        if node.name == "blockquote" and classes.intersection({"chain", "question"}):
            return True

        return False

    for node in direct:
        node.extract()

    chunk: List[Tag] = []

    def flush() -> None:
        nonlocal chunk
        if not chunk:
            return
        flow = soup.new_tag("div", attrs={"class": "flow-columns"})
        for x in chunk:
            flow.append(x)
        container.append(flow)
        chunk = []

    for node in direct:
        if is_full_width(node):
            flush()
            container.append(node)
        else:
            chunk.append(node)
    flush()


def _normalize_search_text(text: str) -> str:
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = text.casefold()
    text = re.sub(r"\s+", " ", text)
    return text


def _reference_signature(reference_text: str) -> Optional[Tuple[str, List[str]]]:
    """Devuelve (año, alias de autor) para asociar citas y referencias.

    La asociación es deliberadamente conservadora: exige que el año y al menos
    un alias de autor aparezcan próximos en el texto de la Parte.
    """
    match = re.search(r"\((\d{4}[a-z]?)\)", reference_text)
    if not match:
        return None

    year = match.group(1)
    author_block = reference_text[:match.start()].strip().rstrip(". ")
    first_author = author_block.split(",", 1)[0].strip()
    if not first_author:
        return None

    aliases = [first_author]
    words = re.findall(r"[A-Za-zÁÉÍÓÚÑÜáéíóúñü]+", first_author)
    if len(words) >= 2:
        stop = {"de", "del", "la", "las", "los", "y", "en", "of", "the", "and", "for", "on", "in"}
        acronym = "".join(w[0] for w in words if w.casefold() not in stop).upper()
        if len(acronym) >= 2:
            aliases.append(acronym)

    # Algunas entidades institucionales suelen citarse mediante una forma corta.
    if "american educational research association" in _normalize_search_text(first_author):
        aliases.append("AERA")
    if "organisation for economic co-operation" in _normalize_search_text(first_author):
        aliases.extend(["OECD", "OCDE"])
    if "subsecretaria de educacion superior" in _normalize_search_text(first_author):
        aliases.append("Subsecretaría de Educación Superior")
    if "ministerio de educacion" in _normalize_search_text(first_author):
        aliases.append("Ministerio de Educación")

    # Deduplicar preservando orden.
    unique: List[str] = []
    seen = set()
    for alias in aliases:
        key = _normalize_search_text(alias)
        if key and key not in seen:
            seen.add(key)
            unique.append(alias)
    return year, unique


def references_used_in_part(part_text: str, reference_nodes: List[Tag]) -> List[Tag]:
    """Selecciona, en orden bibliográfico, las referencias citadas en una Parte."""
    normalized_part = _normalize_search_text(part_text)
    selected: List[Tag] = []

    for ref in reference_nodes:
        ref_text = ref.get_text(" ", strip=True)
        signature = _reference_signature(ref_text)
        if not signature:
            continue
        year, aliases = signature
        found = False
        for alias in aliases:
            a = re.escape(_normalize_search_text(alias))
            y = re.escape(year.casefold())
            # La ventana evita falsos positivos debidos a años numéricos aislados.
            if re.search(a + r".{0,220}?" + y, normalized_part, flags=re.S) or re.search(
                y + r".{0,220}?" + a, normalized_part, flags=re.S
            ):
                found = True
                break
        if found:
            selected.append(ref)
    return selected


def append_part_references(
    soup: BeautifulSoup,
    container: Tag,
    part_heading: str,
    reference_nodes: List[Tag],
) -> int:
    """Añade al final de una Parte solo las referencias efectivamente utilizadas."""
    part_text = container.get_text(" ", strip=True)
    refs = references_used_in_part(part_text, reference_nodes)
    if not refs:
        return 0

    number, _ = split_part_heading(part_heading)
    ident = slugify(f"referencias-{number}")
    h = soup.new_tag("h2", id=ident, attrs={"class": "part-references-heading"})
    h.string = f"Referencias de la {number}"
    container.append(h)

    for ref in refs:
        p = soup.new_tag("p", attrs={"class": "part-reference"})
        # Copiar el contenido HTML de la referencia para conservar cursivas/enlaces.
        for child in list(ref.contents):
            if isinstance(child, NavigableString):
                p.append(NavigableString(str(child)))
            else:
                # BeautifulSoup no permite reutilizar el mismo nodo en varios lugares;
                # se clona mediante serialización mínima.
                clone = BeautifulSoup(str(child), "html.parser")
                for c in list(clone.contents):
                    p.append(c)
        container.append(p)
    return len(refs)


def compose_document(
    fragment_html: str,
    title: str,
    institution: str,
    image_index: Dict[str, Path],
) -> Tuple[str, List[str], List[str]]:
    soup = BeautifulSoup("<html><head></head><body></body></html>", "html.parser")
    fragment = BeautifulSoup(f"<div id='source-root'>{fragment_html}</div>", "html.parser")
    root = fragment.find(id="source-root")
    assert root is not None

    headings = add_heading_ids(root)
    used_images, warnings = transform_figures_positional(soup, root, image_index)
    transform_tables(soup, root)
    transform_blockquotes(root)
    transform_attribute_list(root)

    body = soup.body
    assert body is not None
    body.append(build_cover(soup, title, institution))

    segments = top_level_segments(root)

    # La bibliografía general del Markdown se usa como fuente maestra para
    # construir referencias específicas al final de cada Parte. No se reproduce
    # como una segunda bibliografía global redundante.
    reference_nodes: List[Tag] = []
    for h1, nodes in segments:
        if h1 is not None and h1.get_text(" ", strip=True).upper() == "REFERENCIAS":
            reference_nodes = [n for n in nodes if isinstance(n, Tag) and n.name == "p"]
            break

    toc_headings: List[Tuple[int, str, str]] = []
    toc_placeholder: Optional[Tag] = None

    # Recolectar índice visible desde estructura original, excluyendo la
    # bibliografía global porque será sustituida por referencias por Parte.
    for level, txt, ident in headings:
        if level == 1 and txt.upper() in {"RESUMEN EJECUTIVO", "REFERENCIAS"}:
            continue
        toc_headings.append((level, txt, ident))

    for h1, nodes in segments:
        if h1 is None:
            for n in nodes:
                body.append(n)
            continue

        txt = h1.get_text(" ", strip=True)
        upper = txt.upper()

        if upper == "RESUMEN EJECUTIVO":
            sec = soup.new_tag("section", attrs={"class": "executive-summary"})
            new_h = soup.new_tag("h1", id=h1.get("id", "resumen-ejecutivo"))
            new_h.string = txt
            sec.append(new_h)
            sb = soup.new_tag("div", attrs={"class": "summary-body"})
            for n in nodes:
                if n.name != "hr":
                    sb.append(n)
            sec.append(sb)
            body.append(sec)
            toc_placeholder = soup.new_tag("div", attrs={"id": "_toc_placeholder"})
            body.append(toc_placeholder)

        elif upper.startswith("PARTE "):
            opener, remaining = build_part_opener(soup, h1, nodes)
            body.append(opener)

            # Una Parte constituye UN único flujo editorial. Los H2/H3/H4 no
            # crean contenedores multicolumna nuevos. Esto garantiza que terminar
            # una sección no cambie el orden izquierda -> derecha -> página.
            style_id = part_style_for_heading(txt)
            part_body = soup.new_tag("div", attrs={"class": f"part-body part-style-{style_id}"})
            for n in remaining:
                if n.name != "hr":
                    part_body.append(n)

            append_part_references(soup, part_body, txt, reference_nodes)

            normalize_part_column_flows(soup, part_body)
            body.append(part_body)

        elif upper == "REFERENCIAS":
            # Ya distribuida por Partes.
            continue

        else:
            sec = soup.new_tag("section")
            sec.append(h1)
            for n in nodes:
                sec.append(n)
            body.append(sec)

    # Insertar el TOC una vez conocidas también las referencias por Parte.
    toc = build_toc(soup, toc_headings)
    if toc_placeholder is not None:
        toc_placeholder.replace_with(toc)
    else:
        body.insert(1, toc)

    head = soup.head
    assert head is not None
    meta_charset = soup.new_tag("meta", charset="utf-8")
    meta_author = soup.new_tag("meta", attrs={"name": "author", "content": institution})
    title_tag = soup.new_tag("title")
    title_tag.string = title
    style = soup.new_tag("style")
    style.string = MASTER_CSS
    head.extend([meta_charset, meta_author, title_tag, style])

    used_lower = {x.lower() for x in used_images}
    for key, path in sorted(image_index.items()):
        if key not in used_lower:
            warnings.append(f"Imagen disponible pero no referenciada por el Markdown: {path.name}")

    return str(soup), warnings, used_images


# =============================================================================
# RENDERIZADO SEGMENTADO Y ENSAMBLAJE
# =============================================================================

def html_shell(head_html: str, body_html: str, body_class: str = "render-unit") -> str:
    return f"<!doctype html><html>{head_html}<body class=\"{body_class}\">{body_html}</body></html>"


def split_render_units(final_html: str) -> Tuple[str, List[dict], Tag]:
    """Agrupa el documento por Partes y conserva el flujo interno continuo.

    Solo figuras que requieren una página realmente especial (landscape o
    fullpage) se extraen como unidades. Las figuras hero y las tablas anchas
    permanecen dentro de la Parte, entre tramos ``flow-columns`` secuenciales.
    """
    soup = BeautifulSoup(final_html, "html.parser")
    head_html = str(soup.head)
    body = soup.body
    assert body is not None

    units: List[dict] = []
    toc_node: Optional[Tag] = None
    current_part = ""

    def append_part_body_with_specials(part_body: Tag) -> None:
        nonlocal current_part
        style_id = part_style_for_heading(current_part) if current_part else "1"
        current = soup.new_tag("div", attrs={"class": f"part-body part-style-{style_id}"})

        def flush_current() -> None:
            nonlocal current
            if current.find(True):
                units.append({
                    "kind": "partbody",
                    "part": current_part,
                    "title": "",
                    "node": current,
                })
            style_id = part_style_for_heading(current_part) if current_part else "1"
        current = soup.new_tag("div", attrs={"class": f"part-body part-style-{style_id}"})

        for child in list(part_body.children):
            if not isinstance(child, Tag):
                continue
            classes = set(child.get("class", []))
            if child.name == "figure" and classes.intersection({"landscape", "fullpage"}):
                flush_current()
                kind = "landscape" if "landscape" in classes else "fullpage"
                # Buscar el último H2 previo como contexto de encabezado.
                prev_h2 = part_body.find_previous("h2")
                title = prev_h2.get_text(" ", strip=True) if prev_h2 else ""
                units.append({
                    "kind": kind,
                    "part": current_part,
                    "title": title,
                    "node": child.extract(),
                })
            else:
                current.append(child.extract())
        flush_current()

    for child in list(body.children):
        if not isinstance(child, Tag):
            continue
        classes = set(child.get("class", []))

        if "toc" in classes:
            toc_node = child.extract()
        elif "cover" in classes:
            units.append({"kind": "cover", "part": "", "node": child.extract()})
        elif "executive-summary" in classes:
            units.append({"kind": "summary", "part": "", "node": child.extract()})
        elif "part-opener" in classes:
            h = child.find("h1", class_="part-heading")
            current_part = h.get_text(" ", strip=True) if h else ""
            units.append({"kind": "part", "part": current_part, "node": child.extract()})
        elif "part-body" in classes:
            append_part_body_with_specials(child)
        else:
            units.append({"kind": "section", "part": current_part, "node": child.extract()})

    if toc_node is None:
        toc_node = soup.new_tag("section", attrs={"class": "toc"})
        h = soup.new_tag("h1")
        h.string = "Índice general"
        toc_node.append(h)

    return head_html, units, toc_node


def unit_body_html(unit: dict) -> str:
    part = unit.get("part", "")
    title = unit.get("title", "")
    if unit.get("kind") in {"landscape", "fullpage"}:
        inside = ""
        if part:
            inside += f'<div class="running-part">{html_lib.escape(part)}</div>'
        if title:
            inside += f'<div class="running-section">{html_lib.escape(title)}</div>'
        inside += str(unit["node"])
        wrapper = "landscape-unit" if unit["kind"] == "landscape" else "fullpage-unit"
        return f'<section class="{wrapper}">{inside}</section>'

    prefix = ""
    if part and unit.get("kind") not in {"part", "cover"}:
        prefix += f'<div class="running-part">{html_lib.escape(part)}</div>'
    if title and unit.get("kind") in {"partbody", "section"}:
        prefix += f'<div class="running-section">{html_lib.escape(title)}</div>'
    return prefix + str(unit["node"])


def render_html_to_pdf(html_text: str, output_pdf: Path, base_url: Path) -> None:
    HTML(string=html_text, base_url=str(base_url)).write_pdf(str(output_pdf))


def pdf_page_count(path: Path) -> int:
    doc = fitz.open(path)
    n = doc.page_count
    doc.close()
    return n


# Parámetros del contenedor de página principal (deben reflejar el @page de
# MASTER_CSS). Se usan para medir, en el PDF ya renderizado, cuánto espacio
# queda libre al final de una página antes de un elemento de ancho completo.
_PAGE_TOP_MARGIN_MM = 18.0
_PAGE_BOTTOM_MARGIN_MM = 19.0
_PT_PER_MM = 72.0 / 25.4


def _find_marker_page(doc: "fitz.Document", marker: str) -> Optional[int]:
    for pno in range(1, doc.page_count):
        if doc[pno].search_for(marker):
            return pno
    return None


def _figure_marker(fig: Tag) -> Optional[str]:
    caption = fig.find("figcaption")
    if caption is None:
        return None
    m = re.match(r"^(Figura\s+\d+\.)", caption.get_text(" ", strip=True))
    return m.group(1) if m else None


def tighten_pushed_hero_figures(
    unit_node: Tag,
    pdf_path: Path,
    min_scale: float = 0.52,
    min_abs_mm: float = 55.0,
    gap_threshold_mm: float = 20.0,
    # Espacio no medible directamente en el render "empujado" (colapso del
    # margen superior de .figure con el margen inferior del párrafo previo,
    # más un margen de seguridad frente a redondeos): se resta del hueco
    # disponible para no proponer un tamaño que en la práctica no entre.
    # Calibrado empíricamente contra el documento real (ver notas de commit).
    safety_mm: float = 16.0,
) -> List[Tuple[Tag, str]]:
    """Si una figura .hero queda empujada al inicio de una página y deja un
    hueco grande al final de la página anterior, reduce el alto máximo de
    ESA imagen (solo esa, vía un estilo inline) lo justo para que quepa en
    el hueco disponible, evitando la página con mucho espacio en blanco.

    Nunca reduce una imagen por debajo de ``min_scale`` de su alto ya
    renderizado ni de ``min_abs_mm``, para no volverla ilegible: si el
    hueco no alcanza para un tamaño razonable, la figura se deja intacta
    (mejor un hueco en blanco que un diagrama denso ilegible).

    Devuelve la lista de ``(img_tag, estilo_previo)`` modificados, para que
    el llamador pueda revertir los que, tras volver a renderizar, resulten
    no haber alcanzado a moverse a la página anterior (ver
    ``revert_unplaced_hero_figures``).
    """
    doc = fitz.open(pdf_path)
    top_y = (_PAGE_TOP_MARGIN_MM + 2) * _PT_PER_MM
    bottom_y = (297 - _PAGE_BOTTOM_MARGIN_MM) * _PT_PER_MM

    changed: List[Tuple[Tag, str]] = []
    for fig in unit_node.find_all("figure", class_=lambda c: c and "hero" in c):
        img = fig.find("img")
        marker = _figure_marker(fig)
        if img is None or marker is None:
            continue

        target_page = _find_marker_page(doc, marker)
        if target_page is None or target_page == 0:
            continue

        page = doc[target_page]
        image_info = page.get_image_info()
        if not image_info:
            continue
        img_bbox = image_info[0]["bbox"]
        if img_bbox[1] > top_y + 10 * _PT_PER_MM:
            continue  # la imagen no arranca pegada al tope: no fue empujada
        img_h_mm = (img_bbox[3] - img_bbox[1]) / _PT_PER_MM

        prev_page = doc[target_page - 1]
        prev_blocks = [
            b for b in prev_page.get_text("blocks")
            if b[4].strip() and top_y <= b[1] <= bottom_y - 2 * _PT_PER_MM
        ]
        if not prev_blocks:
            continue
        prev_bottom = max(b[3] for b in prev_blocks)
        gap_mm = (bottom_y - prev_bottom) / _PT_PER_MM
        if gap_mm < gap_threshold_mm:
            continue

        content_blocks = [b for b in page.get_text("blocks") if b[4].strip() and b[1] >= top_y]
        cap_block = next((b for b in content_blocks if b[4].strip().startswith(marker)), None)
        if cap_block is None:
            continue
        overhead_mm = (cap_block[3] - img_bbox[1]) / _PT_PER_MM - img_h_mm

        available_for_image_mm = gap_mm - overhead_mm - safety_mm
        floor_mm = max(img_h_mm * min_scale, min_abs_mm)
        if available_for_image_mm < floor_mm or available_for_image_mm >= img_h_mm:
            continue  # no alcanza a un tamaño legible, o ya cabía: no tocar

        previous_style = img.get("style", "")
        existing_style = previous_style.rstrip("; ")
        new_style = f"max-height:{available_for_image_mm:.1f}mm"
        img["style"] = f"{existing_style}; {new_style}" if existing_style else new_style
        changed.append((img, previous_style))

    doc.close()
    return changed


def revert_unplaced_hero_figures(candidates: List[Tuple[Tag, str]], pdf_path: Path) -> int:
    """Tras volver a renderizar con los tamaños reducidos de
    ``tighten_pushed_hero_figures``, revierte las figuras que, pese al
    ajuste, siguieron arrancando al tope de una página nueva: en ese caso
    reducir el tamaño no evitó el salto de página y solo restaría
    legibilidad sin beneficio alguno.
    """
    doc = fitz.open(pdf_path)
    top_y = (_PAGE_TOP_MARGIN_MM + 2) * _PT_PER_MM
    reverted = 0
    for img, previous_style in candidates:
        fig = img.find_parent("figure")
        marker = _figure_marker(fig) if fig is not None else None
        still_pushed = True
        if marker is not None:
            target_page = _find_marker_page(doc, marker)
            if target_page is not None:
                image_info = doc[target_page].get_image_info()
                if image_info and image_info[0]["bbox"][1] > top_y + 10 * _PT_PER_MM:
                    still_pushed = False
        if still_pushed:
            if previous_style:
                img["style"] = previous_style
            else:
                del img["style"]
            reverted += 1
    doc.close()
    return reverted


def extract_pdf_toc(path: Path) -> List[List]:
    doc = fitz.open(path)
    toc = doc.get_toc(simple=True)
    doc.close()
    return toc


def build_toc_with_pages(toc_node: Tag, page_lookup: Dict[str, int]) -> str:
    """Reemplaza los contadores CSS del índice por páginas físicas ya conocidas."""
    soup = BeautifulSoup(str(toc_node), "html.parser")
    for div in soup.find_all(["div"], class_=lambda x: x and ("toc-part" in x or "toc-sec" in x)):
        a = div.find("a")
        if not a:
            continue
        href = a.get("href", "")
        ident = href[1:] if href.startswith("#") else href
        page = page_lookup.get(ident)
        page_span = div.find("span", class_="toc-page")
        if page_span:
            page_span.attrs.pop("data-href", None)
            page_span.string = str(page) if page else ""
        # El índice visible usa números de página; la navegación interactiva se
        # proporciona mediante el panel de marcadores del PDF ensamblado.
        a.attrs.pop("href", None)
    return str(soup)


def collect_heading_page_lookup(unit_records: List[dict], toc_pages: int, front_cover_pages: int, summary_pages: int) -> Dict[str, int]:
    """Asocia IDs HTML con páginas reales usando el outline de cada unidad."""
    lookup: Dict[str, int] = {}
    offset = 0
    for rec in unit_records:
        final_start = offset + 1
        offset += rec["pages"]
        if rec["kind"] == "summary":
            offset += toc_pages

        node = rec.get("node")
        title_to_ids: Dict[str, List[str]] = {}
        if isinstance(node, Tag):
            for h in node.find_all(["h1", "h2"], recursive=True):
                ident = h.get("id")
                if ident:
                    key = re.sub(r"\s+", " ", h.get_text(" ", strip=True)).casefold()
                    title_to_ids.setdefault(key, []).append(ident)

        local_toc = rec.get("local_toc") or extract_pdf_toc(rec["pdf"])
        for level, title, page in local_toc:
            key = re.sub(r"\s+", " ", str(title).strip()).casefold()
            ids = title_to_ids.get(key)
            if ids:
                ident = ids.pop(0)
                lookup[ident] = final_start + int(page) - 1
    return lookup


def merge_and_finalize(
    unit_records: List[dict],
    toc_pdf: Path,
    output: Path,
) -> Tuple[int, List[List]]:
    """Une unidades, inserta TOC, normaliza marcadores y estampa folios globales."""
    merged = fitz.open()
    global_toc: List[List] = []
    page_offset = 0
    toc_inserted = False
    skip_folio_pages = set()

    def append_pdf(path: Path, kind: str) -> None:
        nonlocal page_offset
        src = fitz.open(path)
        local_toc = src.get_toc(simple=True)
        start = page_offset
        merged.insert_pdf(src)
        for level, title, page in local_toc:
            global_toc.append([level, title, start + page])
        if kind in {"cover", "part", "fullpage"}:
            for pno in range(start, start + src.page_count):
                skip_folio_pages.add(pno)
        page_offset += src.page_count
        src.close()

    for rec in unit_records:
        append_pdf(rec["pdf"], rec["kind"])
        if rec["kind"] == "summary" and not toc_inserted:
            src = fitz.open(toc_pdf)
            toc_start = page_offset
            merged.insert_pdf(src)
            # Marcador explícito para el índice general.
            global_toc.append([1, "Índice general", toc_start + 1])
            page_offset += src.page_count
            src.close()
            toc_inserted = True

    # Si no hubo resumen ejecutivo, insertar índice después de portada.
    if not toc_inserted:
        # Caso no esperado en el reporte actual; mantener salida funcional.
        pass

    # Normalizar niveles de marcadores para evitar saltos inválidos.
    normalized: List[List] = []
    last_level = 0
    for level, title, page in global_toc:
        level = max(1, int(level))
        if last_level and level > last_level + 1:
            level = last_level + 1
        normalized.append([level, title, int(page)])
        last_level = level
    if normalized:
        merged.set_toc(normalized)
    # Solicitar a lectores compatibles que abran el panel de navegación.
    try:
        merged.set_pagemode("UseOutlines")
        merged.set_pagelayout("OneColumn")
    except Exception:
        pass

    # Folio global: 7 pt sans, esquina inferior derecha. Portada/portadillas sin folio.
    for i in range(merged.page_count):
        if i in skip_folio_pages:
            continue
        page = merged[i]
        rect = page.rect
        page.insert_text(
            fitz.Point(rect.width - 48, rect.height - 22),
            str(i + 1),
            fontsize=7,
            fontname="helv",
            color=(0.30, 0.34, 0.35),
            overlay=True,
        )

    output.parent.mkdir(parents=True, exist_ok=True)
    merged.save(output, garbage=4, deflate=True)
    pages = merged.page_count
    merged.close()
    return pages, normalized


# =============================================================================
# CLI Y GENERACIÓN
# =============================================================================

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Genera el Reporte Hito 0 con el diseño editorial maestro V4 mediante HTML/CSS -> PDF."
    )
    p.add_argument("--md", type=Path, help="Ruta al Markdown maestro")
    p.add_argument("--figuras", type=Path, help="Carpeta 01_figuras o ZIP de figuras")
    p.add_argument("--salida", type=Path, help="PDF de salida")
    p.add_argument("--html-debug", type=Path, help="Opcional: guardar HTML completo intermedio")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    cwd = Path.cwd()
    try:
        md_path = (args.md or auto_pick_report_md(cwd)).expanduser().resolve()
        figures_source = (args.figuras or auto_pick_figures(cwd)).expanduser().resolve()
    except Exception as exc:
        print(f"ERROR al seleccionar entradas: {exc}")
        return 2

    if not md_path.exists():
        print(f"ERROR: no existe el Markdown: {md_path}")
        return 2
    if not figures_source.exists():
        print(f"ERROR: no existe la fuente de figuras: {figures_source}")
        return 2

    output = (args.salida or md_path.with_name(md_path.stem + "_EDITORIAL.pdf")).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)

    print("=" * 78)
    print("GENERADOR EDITORIAL HITO 0 — HTML/CSS + WEASYPRINT — v5")
    print("=" * 78)
    print(f"Markdown : {md_path}")
    print(f"Figuras  : {figures_source}")
    print(f"Salida   : {output}")
    print()

    try:
        raw = strip_html_comments(md_path.read_text(encoding="utf-8-sig"))
        title, institution, body_md = extract_cover_fields(raw)

        with tempfile.TemporaryDirectory(prefix="hito0_editorial_") as tmp:
            build_dir = Path(tmp)
            image_index = collect_images(figures_source, build_dir)
            fragment_html = make_markdown_html(body_md)
            final_html, warnings, used_images = compose_document(fragment_html, title, institution, image_index)

            if args.html_debug:
                debug_path = args.html_debug.expanduser().resolve()
                debug_path.parent.mkdir(parents=True, exist_ok=True)
                debug_path.write_text(final_html, encoding="utf-8")
                print(f"HTML debug: {debug_path}")

            head_html, units, toc_node = split_render_units(final_html)
            unit_records: List[dict] = []

            print(f"Unidades editoriales a renderizar: {len(units)}")
            for idx, unit in enumerate(units, start=1):
                unit_pdf = build_dir / f"unit_{idx:03d}_{unit['kind']}.pdf"
                if unit["kind"] == "landscape":
                    body_class = "render-unit render-landscape"
                elif unit["kind"] == "fullpage":
                    body_class = "render-unit render-fullpage"
                else:
                    body_class = "render-unit"
                unit_html = html_shell(head_html, unit_body_html(unit), body_class=body_class)
                label = unit.get("title") or unit.get("part") or unit["kind"]
                print(f"  [{idx:02d}/{len(units):02d}] {unit['kind']}: {label[:72]}")
                render_html_to_pdf(unit_html, unit_pdf, build_dir)
                if unit["kind"] == "partbody":
                    candidates = tighten_pushed_hero_figures(unit["node"], unit_pdf)
                    if candidates:
                        unit_html = html_shell(head_html, unit_body_html(unit), body_class=body_class)
                        render_html_to_pdf(unit_html, unit_pdf, build_dir)
                        reverted = revert_unplaced_hero_figures(candidates, unit_pdf)
                        if reverted:
                            unit_html = html_shell(head_html, unit_body_html(unit), body_class=body_class)
                            render_html_to_pdf(unit_html, unit_pdf, build_dir)
                        print(f"       figuras ajustadas para evitar página en blanco: {len(candidates) - reverted}")
                pages = pdf_page_count(unit_pdf)
                print(f"       páginas: {pages}")
                rec = dict(unit)
                rec.update({"pdf": unit_pdf, "pages": pages, "local_toc": extract_pdf_toc(unit_pdf)})
                unit_records.append(rec)

            cover_pages = sum(r["pages"] for r in unit_records if r["kind"] == "cover")
            summary_pages = sum(r["pages"] for r in unit_records if r["kind"] == "summary")

            # Primera pasada de TOC para conocer su extensión.
            toc_tmp0 = build_dir / "toc0.pdf"
            render_html_to_pdf(html_shell(head_html, str(toc_node)), toc_tmp0, build_dir)
            toc_pages = pdf_page_count(toc_tmp0)

            # Construir páginas reales con TOC insertado tras el resumen.
            page_lookup = collect_heading_page_lookup(unit_records, toc_pages, cover_pages, summary_pages)
            toc_html = build_toc_with_pages(toc_node, page_lookup)
            toc_pdf = build_dir / "toc.pdf"
            render_html_to_pdf(html_shell(head_html, toc_html), toc_pdf, build_dir)
            toc_pages2 = pdf_page_count(toc_pdf)

            if toc_pages2 != toc_pages:
                toc_pages = toc_pages2
                page_lookup = collect_heading_page_lookup(unit_records, toc_pages, cover_pages, summary_pages)
                toc_html = build_toc_with_pages(toc_node, page_lookup)
                render_html_to_pdf(html_shell(head_html, toc_html), toc_pdf, build_dir)

            pages, outline = merge_and_finalize(unit_records, toc_pdf, output)

        print("\nPDF generado correctamente.")
        print(f"Páginas: {pages}")
        print(f"Figuras incorporadas: {len(used_images)}")
        print(f"Marcadores PDF: {len(outline)}")
        print(f"Archivo: {output}")

        if warnings:
            seen = set()
            unique = []
            for w in warnings:
                if w not in seen:
                    seen.add(w); unique.append(w)
            print("\nADVERTENCIAS:")
            for w in unique:
                print(f"  - {w}")
        else:
            print("Advertencias: ninguna")
        return 0

    except Exception as exc:
        print("\nERROR durante la generación:")
        print(f"  {type(exc).__name__}: {exc}")
        print("\nDependencias recomendadas:")
        print("  pip install weasyprint markdown-it-py beautifulsoup4 pillow pymupdf")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
