#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Generador editorial del Diagnostico Hito 0 - Medición Atributos i+e.

Pipeline:
    Markdown -> HTML semántico -> CSS paginado -> PDF (WeasyPrint)

El diseño implementa la especificación "Diseño Maestro Editorial — Diagnóstico
Hito 0 · Línea editorial «Red / Curva / Medición»"
(Diseno_Maestro_Editorial_Hito0_Red_Cian.md): portada y portadillas en cian
institucional (#2EB2C7) con un motivo de red de nodos y vínculos en trayectoria
ascendente sobre un eje vertical discontinuo; en el cuerpo esa identidad se
reduce a reglas finas, barras verticales junto a títulos, captions y cabeceras
de tabla en cian. Tres líneas editoriales asignadas por Parte (A - resultados/
psicometría, B - técnica/metodología, C - conceptual/Construct Maps),
tipografía Roboto Condensed (títulos) + Noto Sans (cuerpo), A4, narrativa a
dos columnas, figuras complejas y tablas anchas a ancho completo, tablas
pequeñas en una columna, encabezados corridos, folios, índice navegable,
marcadores PDF, flujo de columnas estrictamente secuencial y referencias
específicas al final de cada Parte. Las figuras estadísticas/psicométricas no
se recolorean: permanecen en negro y grises.

Dependencias:
    pip install weasyprint markdown-it-py beautifulsoup4 pillow pymupdf
    apt-get install fonts-roboto-unhinted fonts-noto-core   # Roboto Condensed y Noto Sans reales;
                                                             # sin ellos, cae a Liberation Sans/Arial.

Uso recomendado:
    python3 generar_reporte_hito0_editorial_v6.py \
      --md "Reporte_Hito0_Integrado_MEJORADO_v16_AJUSTADO_FACTCHECK.md" \
      --figuras "01_figuras_AJUSTADAS_v16.zip" \
      --salida "Reporte_Hito0_Editorial_V6.pdf"

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
import math
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
#
# "parametros-paso-pcm2" (Figura 12) y "funcionamiento-categorias-pcm"
# (Figura 14) tuvieron aquí un tope de 65/68 mm heredado del layout v3, muy
# por debajo de su alto natural a ancho completo (~132 mm y ~85 mm
# respectivamente para su relación de aspecto real), lo que las dejaba
# ilegibles pese a ocupar el ancho completo ("hero"). El mecanismo dinámico
# de ajuste (tighten_pushed_hero_figures/revert_unplaced_hero_figures, más
# abajo) ya resuelve el problema de espacio en blanco que motivó ese tope
# fijo, así que se retira para ambas y se dejan a su alto natural (sujeto al
# techo general de figura "hero" definido en el CSS).
COMPACT_FIGURE_MAX_HEIGHT_MM: Dict[str, float] = {}


MASTER_CSS = r"""
:root {
  --cyan: #2EB2C7;
  --cyan-deep: #168DA2;
  --cyan-dark: #0D6F80;
  --cyan-mid: #48BDCE;
  --cyan-light: #A7E0E7;
  --cyan-pale: #EAF7F9;
  --paper: #F4FBFC;
  --ink: #171B1D;
  --graphite: #3D4548;
  --gray: #6C777A;
  --rule: #D8E9EC;
  --font-display: "Roboto Condensed", "Arial Narrow", "Liberation Sans", sans-serif;
  --font-sans: "Noto Sans", "Liberation Sans", Arial, sans-serif;
}

@page {
  size: A4;
  margin: 17mm 16mm 18mm 16mm;
  @top-left {
    content: string(current-part);
    font-family: var(--font-display);
    font-size: 6.5pt;
    letter-spacing: .04em;
    text-transform: uppercase;
    color: var(--cyan-deep);
  }
  @top-right {
    content: string(current-section);
    font-family: var(--font-sans);
    font-size: 6.3pt;
    color: var(--gray);
  }
  @bottom-left {
    content: "Facultad de Ingeniería · Universidad de Santiago de Chile";
    font-family: var(--font-sans);
    font-size: 6.2pt;
    color: var(--gray);
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
  margin: 15mm 16mm 16mm 16mm;
  @top-left {
    content: string(current-part);
    font-family: var(--font-display);
    font-size: 6.5pt;
    letter-spacing: .04em;
    text-transform: uppercase;
    color: var(--cyan-deep);
  }
  @top-right {
    content: string(current-section);
    font-family: var(--font-sans);
    font-size: 6.3pt;
    color: var(--gray);
  }
  @bottom-left {
    content: "Facultad de Ingeniería · Universidad de Santiago de Chile";
    font-family: var(--font-sans);
    font-size: 6.2pt;
    color: var(--gray);
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

html { font-size: 8.95pt; }
body {
  margin: 0;
  color: var(--ink);
  font-family: var(--font-sans);
  font-size: 8.95pt;
  line-height: 1.49;
  text-rendering: optimizeLegibility;
  hyphens: auto;
  background: #fff;
}

p {
  margin: 0 0 3.1mm;
  text-align: justify;
  text-align-last: left;
  orphans: 2;
  widows: 2;
}

strong { font-weight: 700; color: #10161a; }
em { font-style: italic; }
a { color: var(--cyan-dark); text-decoration: none; }
code {
  font-family: "Liberation Mono", monospace;
  font-size: .92em;
  color: var(--cyan-dark);
  overflow-wrap: anywhere;
}

/* =========================================================================
   RED / NODOS — motivo geométrico compartido por portada y portadillas
   ========================================================================= */
.net-line { stroke-linecap: round; }
.net-node { }

/* ------------------------------------------------------------------------- */
/* PORTADA                                                                  */
/* ------------------------------------------------------------------------- */
.cover {
  page: cover;
  break-after: page;
  height: 297mm;
  width: 210mm;
  position: relative;
  overflow: hidden;
  background: var(--cyan);
  color: #fff;
}
.cover-network {
  position: absolute;
  left: 0; top: 0;
  width: 210mm;
  height: 297mm;
  z-index: 1;
}
.cover-axis {
  position: absolute;
  left: 147mm;
  top: 16mm;
  bottom: 16mm;
  width: 0;
  border-left: .8pt dashed rgba(255,255,255,.55);
  z-index: 2;
}
.cover-inner {
  position: relative;
  z-index: 3;
  height: 100%;
  width: 130mm;
  padding: 15mm 0 15mm 15mm;
  display: flex;
  flex-direction: column;
}
.cover-header {
  font-family: var(--font-display);
  font-size: 9.8pt;
  font-weight: 700;
  letter-spacing: .06em;
  line-height: 1.65;
  text-transform: uppercase;
  color: #eafcff;
  max-width: 108mm;
}
.cover-title {
  margin: 68mm 0 0;
  font-family: var(--font-display);
  font-weight: 700;
  font-size: 61pt;
  line-height: .95;
  letter-spacing: -.01em;
  text-transform: uppercase;
  color: #fff;
}
.cover-title span { display: block; }
.cover-subtitle {
  margin-top: 6mm;
  display: flex;
  flex-direction: column;
  font-family: var(--font-sans);
  font-size: 25pt;
  font-weight: 500;
  line-height: 1.12;
  color: #eafcff;
  max-width: 110mm;
}
.cover-tagline {
  margin-top: 13mm;
  font-family: var(--font-sans);
  font-size: 11.5pt;
  font-weight: 500;
  color: #fff;
}
.cover-footer {
  position: absolute;
  left: 15mm;
  bottom: 15mm;
  z-index: 3;
  font-family: var(--font-sans);
  font-size: 9.3pt;
  font-weight: 500;
  color: #eafcff;
}

/* ------------------------------------------------------------------------- */
/* RESUMEN EJECUTIVO                                                        */
/* ------------------------------------------------------------------------- */
.executive-summary {
  page: auto;
  break-before: page;
  break-after: page;
  padding-top: 2mm;
}
.executive-summary > h1 {
  font-family: var(--font-display);
  color: var(--ink);
  font-size: 21pt;
  font-weight: 700;
  line-height: 1.05;
  margin: 0 0 6mm;
  padding: 0 0 4mm 4mm;
  border-left: 1.3mm solid var(--cyan);
  bookmark-level: 1;
}
.executive-summary .summary-body { column-count: 1; }
.executive-summary .summary-body > p:first-of-type {
  font-size: 9.9pt;
  line-height: 1.53;
  color: var(--graphite);
  margin-bottom: 5mm;
}
.executive-summary strong { color: var(--cyan-dark); }

/* ------------------------------------------------------------------------- */
/* ÍNDICE                                                                   */
/* ------------------------------------------------------------------------- */
.toc {
  break-before: page;
  break-after: page;
  font-family: var(--font-sans);
}
.toc h1 {
  font-family: var(--font-display);
  font-size: 21pt;
  font-weight: 700;
  color: var(--ink);
  padding: 0 0 4mm 4mm;
  border-left: 1.3mm solid var(--cyan);
  margin: 0 0 8mm;
}
.toc-list { column-count: 2; column-gap: 10mm; }
.toc-part, .toc-sec {
  break-inside: avoid;
  margin: 0 0 2.6mm;
}
.toc-part { margin-top: 3.5mm; }
.toc-part a, .toc-sec a {
  display: flex;
  gap: 2mm;
  align-items: baseline;
  color: var(--graphite);
}
.toc-part a { font-weight: 700; color: var(--cyan-dark); font-family: var(--font-display); }
.toc-sec a { font-size: 8.1pt; }
.toc-sec { padding-left: 5mm; }
.toc-num { color: var(--cyan); font-weight: 700; min-width: 9mm; }
.toc-text { flex: 1; }
.toc-page::after { content: target-counter(attr(data-href), page); }
.toc-page {
  margin-left: auto;
  color: var(--gray);
  font-variant-numeric: tabular-nums;
}

/* ------------------------------------------------------------------------- */
/* PORTADILLAS DE PARTE                                                     */
/* ------------------------------------------------------------------------- */
.part-opener {
  page: part;
  break-before: page;
  break-after: page;
  break-inside: avoid;
  height: 297mm;
  width: 210mm;
  margin: 0;
  position: relative;
  overflow: hidden;
  background: #fff;
  font-family: var(--font-sans);
}
.part-opener::before {
  content: "";
  position: absolute;
  left: 0; right: 0; top: 0;
  height: 61%;
  background: var(--cyan);
  z-index: 1;
}
.part-opener .part-network {
  position: absolute;
  right: 0; top: 0;
  width: 130mm;
  height: 181mm;
  z-index: 2;
}
.part-opener .part-axis {
  position: absolute;
  left: 147mm;
  top: 14mm;
  height: 155mm;
  width: 0;
  border-left: .7pt dashed rgba(255,255,255,.6);
  z-index: 3;
}
.part-opener .part-content {
  position: relative;
  z-index: 4;
  padding: 20mm 20mm 0;
}
.part-opener .part-kicker {
  color: #eafcff;
  font-family: var(--font-display);
  text-transform: uppercase;
  letter-spacing: .1em;
  font-size: 8pt;
  font-weight: 700;
  margin: 0 0 9mm;
}
.part-opener .part-kicker .no-upper { text-transform: none; }
.part-opener h1.part-heading {
  margin: 0;
  color: #fff;
  bookmark-level: 1;
  string-set: current-part content();
}
.part-opener .part-number {
  display: block;
  font-family: var(--font-display);
  font-size: 13.5pt;
  font-weight: 700;
  line-height: 1;
  letter-spacing: .02em;
  margin-bottom: 5mm;
  color: #eafcff;
}
.part-opener .part-name {
  display: block;
  max-width: 118mm;
  font-family: var(--font-display);
  font-weight: 700;
  font-size: 31pt;
  line-height: 1.04;
  color: var(--ink);
}
.part-opener.title-long .part-name { font-size: 27pt; max-width: 130mm; }
.part-opener.title-xlong .part-name { font-size: 23.5pt; line-height: 1.08; max-width: 138mm; }
.part-opener .part-lower {
  position: absolute;
  z-index: 4;
  top: 61%;
  left: 20mm;
  right: 20mm;
  bottom: 16mm;
  display: flex;
  flex-direction: column;
}
.part-description {
  max-width: 158mm;
  color: var(--graphite);
  font-size: 9.3pt;
  line-height: 1.46;
  margin: 8mm 0 6mm;
}
.part-description p { text-align: justify; text-align-last: left; margin-bottom: 3mm; }
.part-description blockquote {
  margin: 0 0 5mm;
  padding: 3.5mm 4.5mm;
  background: var(--paper);
  border-left: 3px solid var(--cyan);
}
.part-mini-index {
  margin-top: auto;
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 0 12mm;
  border-top: 1.1pt solid var(--cyan);
  padding-top: 4mm;
}
.part-mini-index a {
  display: grid;
  grid-template-columns: 14mm 1fr;
  gap: 2.5mm;
  align-items: start;
  min-height: 15mm;
  padding: 3.6mm 0 4mm;
  border-bottom: .5pt solid var(--rule);
  font-size: 9.4pt;
  line-height: 1.24;
  color: var(--graphite);
}
.part-mini-index .mini-num {
  font-family: var(--font-display);
  color: var(--cyan-dark);
  font-weight: 700;
  font-size: 12.5pt;
  line-height: 1;
}
.part-mini-index .mini-label { font-weight: 550; }
.part-mini-index.index-few { grid-template-columns: 1fr 1fr; }
.part-mini-index.index-few a {
  min-height: 21mm;
  padding-top: 5mm;
  font-size: 10.6pt;
  line-height: 1.25;
}
.part-mini-index.index-few .mini-num { font-size: 14.5pt; }
.part-mini-index.index-many {
  grid-template-columns: repeat(3, 1fr);
  gap: 0 7mm;
}
.part-mini-index.index-many a {
  grid-template-columns: 9mm 1fr;
  min-height: 11.5mm;
  padding: 2.4mm 0 2.6mm;
  font-size: 7.7pt;
}
.part-mini-index.index-many .mini-num { font-size: 9.6pt; }

/* ------------------------------------------------------------------------- */
/* CUERPO EN DOS COLUMNAS                                                   */
/* ------------------------------------------------------------------------- */
.section-major { break-before: auto; margin: 0; }
.section-major + .section-major { margin-top: 5mm; }

.flow-columns {
  column-count: 2;
  column-gap: 9mm;
  column-fill: balance;
}
.flow-columns::after { content: ""; display: block; clear: both; }
/* Tramo justo antes de una figura/tabla a todo el ancho: a una columna, no
   a dos. Con column-count:2 (aunque sea balance), un párrafo corto se
   parte justo antes del salto y dejaba 1-2 líneas sueltas -muy estiradas
   por el justify- al tope de la columna derecha; con una sola columna el
   texto simplemente continúa hasta terminar, sin "saltar" a una segunda
   columna, y de paso ese texto ancho funciona como remate/introducción de
   la figura o tabla que sigue. (column-fill:auto con altura automática NO
   sirve para esto: es un bug confirmado de WeasyPrint que deja TODO el
   contenido en la columna izquierda y una derecha en blanco.) */
.flow-columns-preflush { column-count: 1; }

h2, h3, h4 { font-family: var(--font-display); break-after: avoid; }
h2 {
  font-size: 15.4pt;
  font-weight: 700;
  line-height: 1.1;
  margin: 0 0 4.2mm;
  padding: .5mm 0 .5mm 4mm;
  border-left: 1.3mm solid var(--cyan);
  color: var(--ink);
  string-set: current-section content();
  bookmark-level: 2;
}
h3 {
  font-size: 11pt;
  font-weight: 700;
  line-height: 1.15;
  margin: 5mm 0 2.5mm;
  color: var(--cyan-dark);
  bookmark-level: 3;
}
h4 {
  font-family: var(--font-sans);
  font-size: 9.3pt;
  font-weight: 700;
  line-height: 1.2;
  margin: 4mm 0 2mm;
  color: var(--graphite);
  bookmark-level: 4;
}

ul, ol { margin: 2mm 0 4mm; padding-left: 5mm; }
li { margin-bottom: 2.1mm; break-inside: auto; }
li::marker { color: var(--cyan); }

.attrs {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 2.2mm 7mm;
  margin: 4mm 0 6mm;
  padding: 5mm 6mm 5mm 11mm;
  background: var(--paper);
  border-top: .6pt solid var(--cyan-light);
  border-bottom: .6pt solid var(--cyan-light);
  font-family: var(--font-sans);
}
.attrs li { margin: 0; padding-left: 1mm; }
.attrs li::marker { color: var(--cyan-dark); font-weight: 700; }

hr { border: 0; border-top: .6pt solid var(--rule); margin: 5mm 0; }

/* ------------------------------------------------------------------------- */
/* FIGURAS — no se recolorean; el cian pertenece al sistema, no a los datos */
/* ------------------------------------------------------------------------- */
.figure { margin: 6.5mm 0 8mm; break-inside: avoid; width: 100%; }
.figure.hero { break-before: auto; }
.figure.landscape { page: landscape; break-before: page; break-after: page; }
.figure.fullpage {
  page: figurepage;
  break-before: page;
  break-after: page;
  margin: 0;
  width: 100%;
}
.figure.fullpage .figure-frame { border: 0; padding: 0; }
.figure.fullpage img { width: 100%; max-height: 250mm; object-fit: contain; }
.figure.fullpage figcaption { margin-top: 3mm; }
.figure-frame { background: #fff; text-align: center; }
.figure img {
  width: 100%;
  height: auto;
  max-height: 150mm;
  object-fit: contain;
  display: block;
  margin: 0 auto;
}
.figure.hero img { max-height: 172mm; }
.figure.landscape img { max-height: 155mm; }
.figure figcaption {
  font-family: var(--font-sans);
  font-size: 7.3pt;
  line-height: 1.36;
  color: var(--gray);
  text-align: left;
  margin-top: 2.4mm;
  padding-top: 2.2mm;
  border-top: .5pt solid var(--rule);
  break-inside: avoid;
}
.figure figcaption strong {
  font-family: var(--font-display);
  color: var(--cyan-dark);
  font-weight: 700;
}

/* ------------------------------------------------------------------------- */
/* TABLAS                                                                   */
/* ------------------------------------------------------------------------- */
.table-title {
  font-family: var(--font-sans);
  font-size: 7.6pt;
  line-height: 1.35;
  color: var(--gray);
  margin: 5mm 0 2.3mm;
  break-after: avoid;
}
.table-title strong {
  font-family: var(--font-display);
  color: var(--cyan-dark);
  font-weight: 700;
}
.panel-label {
  font-family: var(--font-display);
  font-size: 7.8pt;
  font-weight: 700;
  color: var(--graphite);
  margin: 3mm 0 1.5mm;
  break-after: avoid;
}
.table-group { margin: 4mm 0 5mm; }
.table-group.small-table { margin: 3mm 0 4.5mm; break-inside: avoid; }
.table-group.wide-table { break-inside: auto; }
.table-block { margin: 0; break-inside: auto; }
.table-group.small-table table { font-size: 7.7pt; line-height: 1.28; }
.table-group.small-table .table-title { margin-top: 0; font-size: 7.8pt; }
table {
  width: 100%;
  border-collapse: collapse;
  font-family: var(--font-sans);
  font-size: 7.3pt;
  line-height: 1.28;
}
thead { display: table-header-group; }
tr { break-inside: avoid; }
th {
  background: var(--cyan-pale);
  color: var(--ink);
  text-align: left;
  padding: 2.3mm 2.5mm;
  font-weight: 700;
  vertical-align: bottom;
  border-top: 1pt solid var(--cyan);
  border-bottom: .6pt solid var(--cyan-light);
}
td {
  padding: 2mm 2.5mm;
  border-bottom: .5pt solid var(--rule);
  vertical-align: top;
}
tbody tr:nth-child(even) td { background: #fbfdfd; }
table.matrix { font-size: 6.2pt; line-height: 1.22; }
table.medium { font-size: 6.8pt; line-height: 1.24; }
.table-note { font-size: 7.4pt; color: var(--gray); margin-top: 1.3mm; margin-bottom: 5mm; }

/* ------------------------------------------------------------------------- */
/* CAJAS / CITAS / CADENAS                                                  */
/* ------------------------------------------------------------------------- */
blockquote {
  margin: 3.5mm 0 4.5mm;
  padding: 3.5mm 4.5mm;
  background: #fff;
  border-left: 2.2px solid var(--cyan-light);
  color: var(--graphite);
  break-inside: auto;
}
blockquote.note-box, blockquote.keybox, blockquote.question, blockquote.chain { break-inside: avoid; }
blockquote p:last-child { margin-bottom: 0; }
.keybox {
  margin: 5mm 0 6mm;
  padding: 5mm 6mm;
  background: var(--paper);
  border-left: 3px solid var(--cyan);
  font-size: 9.4pt;
  line-height: 1.47;
  break-inside: avoid;
}
.note-box {
  margin: 5mm 0;
  padding: 4.5mm 5.5mm;
  background: var(--paper);
  border-left: 3px solid var(--cyan-dark);
  break-inside: avoid;
}
.question {
  margin: 5mm 0 6mm;
  padding: 6mm 8mm;
  background: var(--cyan);
  color: #fff;
  font-family: var(--font-display);
  font-size: 13.5pt;
  font-weight: 700;
  line-height: 1.3;
  text-align: center;
  break-inside: avoid;
}
.question p { text-align: center; margin: 0; }
.chain {
  margin: 5mm 0 6mm;
  padding: 4mm 5mm;
  border-top: 1pt solid var(--cyan);
  border-bottom: .5pt solid var(--cyan-light);
  background: var(--cyan-pale);
  text-align: center;
  font-family: var(--font-display);
  font-weight: 700;
  color: var(--cyan-dark);
  font-size: 8.9pt;
  line-height: 1.6;
  break-inside: avoid;
}
.chain p { text-align: center; margin: 0; }

/* ------------------------------------------------------------------------- */
/* ECUACIONES                                                               */
/* ------------------------------------------------------------------------- */
.math-inline { font-family: var(--font-sans); font-style: italic; white-space: nowrap; }
.equation {
  text-align: center;
  font-family: var(--font-sans);
  font-style: italic;
  font-size: 13pt;
  padding: 3mm 4mm;
  margin: 3mm 0 4mm;
  color: var(--cyan-dark);
  background: var(--cyan-pale);
  border-top: .6pt solid var(--cyan-light);
  border-bottom: .6pt solid var(--cyan-light);
  break-inside: avoid;
}
.equation sub, .math-inline sub { font-size: .72em; }
.equation sup, .math-inline sup { font-size: .72em; }

/* ------------------------------------------------------------------------- */
/* REFERENCIAS                                                             */
/* ------------------------------------------------------------------------- */
.references-section { break-before: page; }
.references-section > h1 {
  font-family: var(--font-display);
  font-size: 20pt;
  font-weight: 700;
  line-height: 1.05;
  color: var(--ink);
  margin: 0 0 6mm;
  padding: 0 0 3.2mm 4mm;
  border-left: 1.3mm solid var(--cyan);
  string-set: current-section content();
  bookmark-level: 1;
}
.references-columns { column-count: 2; column-gap: 9mm; font-size: 8.15pt; line-height: 1.36; }
.references-columns p {
  margin: 0 0 2.6mm;
  padding-left: 5mm;
  text-indent: -5mm;
  text-align: left;
  overflow-wrap: anywhere;
  break-inside: avoid;
}

.part-references-heading {
  font-family: var(--font-display);
  font-size: 11pt;
  font-weight: 700;
  line-height: 1.15;
  color: var(--cyan-dark);
  margin: 7mm 0 3mm;
  padding-top: 2.5mm;
  border-top: 1pt solid var(--cyan);
  break-after: avoid;
  bookmark-level: 3;
}
.part-reference {
  font-size: 8.1pt;
  line-height: 1.34;
  margin: 0 0 2.5mm;
  padding-left: 4.5mm;
  text-indent: -4.5mm;
  text-align: left;
  overflow-wrap: anywhere;
  break-inside: avoid;
}

/* ------------------------------------------------------------------------- */
/* RENDERIZADO POR UNIDADES (mecánica del pipeline; no tocar sin motivo)    */
/* ------------------------------------------------------------------------- */
.running-part {
  string-set: current-part content();
  height: 0; max-height: 0; overflow: hidden;
  color: transparent; font-size: 0; line-height: 0; margin: 0; padding: 0;
}
.running-section {
  string-set: current-section content();
  height: 0; max-height: 0; overflow: hidden;
  color: transparent; font-size: 0; line-height: 0; margin: 0; padding: 0;
}
.render-landscape { page: landscape; }
.render-fullpage { page: figurepage; }
.fullpage-unit { page: figurepage; }
.fullpage-unit > .figure.fullpage { page: auto; break-before: auto; break-after: auto; }
.landscape-unit { page: landscape; }
.landscape-unit > .figure.landscape { page: auto; break-before: auto; break-after: auto; margin-top: 0; }
.render-unit > .section-major,
.render-unit > .executive-summary,
.render-unit > .part-opener,
.render-unit > .references-section,
.render-unit > .toc { break-before: auto; }

/* ------------------------------------------------------------------------- */
/* LÍNEA A — resultados / psicometría (Parte II)                            */
/* ------------------------------------------------------------------------- */
.part-body.part-style-1 h2 { font-size: 15.1pt; margin-bottom: 4mm; }
.part-body.part-style-1 h3 { font-size: 10.8pt; margin-top: 4.7mm; }
.part-body.part-style-1 .figure.hero { margin: 7mm 0 9mm; }
.part-body.part-style-1 .figure.hero img { max-height: 174mm; }
.part-body.part-style-1 table { font-size: 7.2pt; }

/* ------------------------------------------------------------------------- */
/* LÍNEA B — técnica / metodología (Partes IV y VI)                         */
/* ------------------------------------------------------------------------- */
.part-body.part-style-2 { font-size: 8.8pt; line-height: 1.46; }
.part-body.part-style-2 .flow-columns { column-gap: 8.5mm; }
.part-body.part-style-2 h2 {
  font-size: 13.8pt;
  padding: .4mm 0 .4mm 3.4mm;
  border-left-width: 1.1mm;
  margin-bottom: 3.6mm;
}
.part-body.part-style-2 h3 {
  font-size: 10pt;
  margin: 4.3mm 0 2mm;
  padding-left: 2.3mm;
  border-left: 1.4pt solid var(--cyan-mid);
  color: var(--cyan-dark);
}
.part-body.part-style-2 h4 { font-size: 9pt; }
.part-body.part-style-2 table { font-size: 7pt; line-height: 1.23; }
.part-body.part-style-2 th { padding: 2mm 2.2mm; }
.part-body.part-style-2 td { padding: 1.75mm 2.2mm; }
.part-body.part-style-2 .table-title { margin-top: 4.3mm; }
.part-body.part-style-2 blockquote { font-size: 8.5pt; }
.part-body.part-style-2 .keybox { background: var(--paper); }

/* ------------------------------------------------------------------------- */
/* LÍNEA C — conceptual / Construct Maps (Partes I, III, V, VII)            */
/* ------------------------------------------------------------------------- */
.part-body.part-style-3 { line-height: 1.52; }
.part-body.part-style-3 h2 {
  font-size: 16.4pt;
  line-height: 1.08;
  margin-bottom: 5mm;
}
.part-body.part-style-3 h3 { font-size: 11.3pt; margin: 5.4mm 0 2.5mm; }
.part-body.part-style-3 .figure { margin: 7.5mm 0 9mm; }
.part-body.part-style-3 .figure.hero { margin: 9mm 0 11mm; }
.part-body.part-style-3 .figure figcaption { padding-top: 2.6mm; }
.part-body.part-style-3 blockquote.keybox,
.part-body.part-style-3 blockquote.chain { padding: 5mm 6mm; margin: 6mm 0; }

/* Reduce riesgo de cortes en portadillas y encabezados */
.part-opener, .cover { break-inside: avoid; }
.part-opener .part-mini-index a { break-inside: avoid; }
h2, h3, h4 { break-after: avoid; }

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


def _copy_as_grayscale(src: Path, dst: Path) -> None:
    """Convierte una imagen ráster a escala de grises real.

    El diseño maestro exige que las figuras se integren "al sistema
    monocromático" y el CSS declara ``filter: grayscale(100%)`` sobre
    ``.figure img`` con esa intención. Se comprobó empíricamente que
    WeasyPrint 69 ignora en silencio la propiedad ``filter`` (la imagen se
    embebe sin alterar), así que ese único CSS no garantiza monocromatismo
    si alguna figura futura no llega ya en grises. Se convierte aquí, en la
    imagen misma, para que la garantía sea real y no dependa de que el
    origen ya esté en grises.
    """
    with PILImage.open(src) as im:
        if im.mode in ("RGBA", "LA") or ("transparency" in im.info):
            rgba = im.convert("RGBA")
            gray = rgba.convert("L").convert("RGBA")
            gray.putalpha(rgba.getchannel("A"))
            gray.save(dst)
        else:
            im.convert("L").save(dst)


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
        if p.suffix.lower() == ".svg":
            # No hay una forma segura de forzar escala de grises en un SVG
            # arbitrario sin parsearlo; se conserva tal cual (no se usan
            # SVG en el material actual del reporte).
            shutil.copy2(p, target)
        else:
            _copy_as_grayscale(p, target)
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


def _mesh_edges(nodes: List[Tuple[float, float]], threshold: float, max_k: int = 3) -> List[Tuple[int, int]]:
    """Construye los vínculos de una malla triangulada por vecino-más-cercano.

    Cada nodo se conecta con sus ``max_k`` vecinos más próximos que estén a
    una distancia menor que ``threshold``. Con nodos ya dispuestos a lo largo
    de una trayectoria (ver ``_cover_network_svg``/``_part_network_svg``) este
    criterio reproduce, sin necesidad de una triangulación de Delaunay
    completa, la textura de "cinta" triangulada de la portada aprobada.
    """
    edges: set = set()
    n = len(nodes)
    for i in range(n):
        xi, yi = nodes[i]
        dists = sorted(
            ((math.hypot(xi - nodes[j][0], yi - nodes[j][1]), j) for j in range(n) if j != i)
        )
        count = 0
        for dist, j in dists:
            if dist > threshold or count >= max_k:
                break
            edges.add((i, j) if i < j else (j, i))
            count += 1
    return sorted(edges)


def _network_svg(
    width_mm: float,
    height_mm: float,
    nodes: List[Tuple[float, float]],
    edges: List[Tuple[int, int]],
    node_r: float = 1.3,
    stroke: str = "#ffffff",
    stroke_opacity: float = 0.85,
    node_fill: str = "#ffffff",
    node_opacity: float = 1.0,
) -> str:
    """Genera el motivo de "red" (malla triangulada, trayectoria ascendente)
    compartido por la portada y las portadillas de Parte, como SVG inline.

    ``nodes`` son las coordenadas (x_mm, y_mm) de cada nodo y ``edges`` los
    pares de índices de ``nodes`` que se conectan entre sí.
    """
    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {width_mm} {height_mm}" '
        f'width="{width_mm}mm" height="{height_mm}mm" preserveAspectRatio="none">'
    ]
    for i, j in edges:
        x1, y1 = nodes[i]
        x2, y2 = nodes[j]
        parts.append(
            f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{stroke}" '
            f'stroke-opacity="{stroke_opacity}" stroke-width="0.5" stroke-linecap="round"/>'
        )
    for x, y in nodes:
        parts.append(f'<circle cx="{x}" cy="{y}" r="{node_r}" fill="{node_fill}" fill-opacity="{node_opacity}"/>')
    parts.append("</svg>")
    return "".join(parts)


# Coordenadas (mm, sobre una página A4 completa de 210x297mm) de los nodos de
# la malla en red de la portada aprobada (Muestra_Editorial_Hito0_Red_Cian.pdf),
# obtenidas por detección de círculos sobre el PDF de referencia. Se reutilizan
# tal cual para que la portada generada reproduzca fielmente ese diseño.
_COVER_NETWORK_NODES: List[Tuple[float, float]] = [
    (190.8, 17.1), (182.6, 29.8), (191.5, 37.0), (176.9, 50.4), (193.9, 53.0),
    (182.8, 66.7), (168.4, 69.4), (191.0, 78.8), (179.4, 83.5), (162.1, 86.9),
    (171.5, 98.0), (183.3, 98.2), (155.2, 102.5), (166.2, 115.2), (147.4, 121.3),
    (151.0, 131.8), (132.7, 134.0), (118.6, 143.5), (136.1, 146.8), (102.0, 155.1),
    (121.3, 156.6), (107.4, 168.9), (99.0, 185.1), (80.3, 187.6), (86.7, 203.4),
    (99.1, 208.5), (71.3, 209.2), (87.1, 223.3), (67.5, 228.6), (104.2, 230.1),
    (81.9, 242.4), (63.9, 249.1), (95.1, 249.7), (113.1, 251.8), (76.7, 262.5),
    (95.5, 267.2), (61.3, 269.9), (77.8, 274.7),
]

# Ídem para la malla (recortada) de las portadillas de Parte, referida al
# contenedor local ``.part-network`` (130x181mm, alineado al borde derecho de
# la página) — coordenadas de página trasladadas en x = página_x - 80mm.
_PART_NETWORK_NODES: List[Tuple[float, float]] = [
    (125.5, 20.3), (122.1, 32.5), (129.8, 34.8), (114.3, 40.3), (123.2, 44.7),
    (109.9, 51.5), (117.6, 54.8), (102.1, 59.2), (109.9, 62.5), (97.7, 69.2),
    (104.3, 72.6), (88.8, 74.8), (97.6, 80.3), (82.1, 84.8), (89.8, 89.2),
    (73.2, 92.6), (81.0, 98.1), (63.3, 102.5), (82.0, 111.5), (66.5, 115.9),
    (73.2, 122.5), (57.8, 127.0), (74.2, 136.0), (60.9, 140.5), (67.6, 146.0),
    (53.2, 149.3), (69.8, 158.2), (55.4, 161.6), (65.4, 169.3), (49.8, 172.7),
]


def _cover_network_svg() -> str:
    edges = _mesh_edges(_COVER_NETWORK_NODES, threshold=24.0, max_k=3)
    return _network_svg(210, 297, _COVER_NETWORK_NODES, edges, node_r=1.35, stroke_opacity=0.85)


def _part_network_svg() -> str:
    edges = _mesh_edges(_PART_NETWORK_NODES, threshold=22.0, max_k=3)
    return _network_svg(130, 181, _PART_NETWORK_NODES, edges, node_r=1.3, stroke_opacity=0.85)


def build_cover(soup: BeautifulSoup, title: str, institution: str) -> Tag:
    """Portada V6: línea editorial «Red / Curva / Medición» (fondo cian, red ascendente, eje discontinuo)."""
    section = soup.new_tag("section", attrs={"class": "cover cover-v6"})
    inner = soup.new_tag("div", attrs={"class": "cover-inner"})

    network = soup.new_tag("div", attrs={"class": "cover-network"})
    network.append(BeautifulSoup(_cover_network_svg(), "html.parser"))
    axis = soup.new_tag("div", attrs={"class": "cover-axis"})

    # Identificación institucional pequeña, dos líneas (Facultad / Universidad),
    # en el borde superior izquierdo — sin columna derecha, tal como la portada
    # aprobada (Muestra_Editorial_Hito0_Red_Cian.pdf).
    if "," in institution:
        f, rest = institution.split(",", 1)
    else:
        f, rest = institution, ""
    header = soup.new_tag("div", attrs={"class": "cover-header"})
    hline1 = soup.new_tag("div"); hline1.string = f.strip().upper()
    header.append(hline1)
    if rest.strip():
        hline2 = soup.new_tag("div"); hline2.string = rest.strip().upper()
        header.append(hline2)

    h = soup.new_tag("h1", attrs={"class": "cover-title"})
    line1 = soup.new_tag("span"); line1.string = "Diagnóstico"
    line2 = soup.new_tag("span"); line2.string = "Hito 0"
    # Espacio explícito (ver nota equivalente en build_part_opener): evita que
    # WeasyPrint concatene sin separación el título del marcador/outline del PDF.
    h.extend([line1, NavigableString(" "), line2])

    subtitle = soup.new_tag("div", attrs={"class": "cover-subtitle"})
    sub1 = soup.new_tag("span"); sub1.string = "Medición Atributos i+e"
    sub2 = soup.new_tag("span"); sub2.string = "Perfil de Ingreso"
    subtitle.extend([sub1, sub2])

    # Cierre editorial (sección 2.1 del diseño maestro): la Facultad en flujo
    # normal bajo el subtítulo, y un descriptor de cierre pequeño anclado al
    # borde inferior izquierdo — replica la composición de la portada aprobada.
    tagline = soup.new_tag("div", attrs={"class": "cover-tagline"})
    tagline.string = f.strip()

    footer = soup.new_tag("div", attrs={"class": "cover-footer"})
    footer.string = "Reporte académico · 2026"

    inner.extend([header, h, subtitle, tagline])
    section.extend([network, axis, inner, footer])
    return section


def part_style_for_heading(text: str) -> str:
    """Mapea Partes a las tres líneas editoriales del diseño Red/Cian.

    Línea A (id interno "1"): resultados/psicometría (Parte II)
    Línea B (id interno "2"): técnica/metodología (Partes IV y VI)
    Línea C (id interno "3"): conceptual/Construct Maps (Partes I, III, V y VII)
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
    kicker.string = "DIAGNÓSTICO HITO 0 · MEDICIÓN ATRIBUTOS "
    kicker_ie = soup.new_tag("span", attrs={"class": "no-upper"}); kicker_ie.string = "i+e"
    kicker.append(kicker_ie)
    heading = soup.new_tag("h1", attrs={"class": "part-heading", "id": h1.get("id", slugify(text))})
    ns = soup.new_tag("span", attrs={"class": "part-number"}); ns.string = number
    nm = soup.new_tag("span", attrs={"class": "part-name"}); nm.string = name
    # El espacio explícito entre ambos <span> no es visible (part-name es
    # display:block), pero evita que WeasyPrint concatene sin separación el
    # texto de ambos nodos al generar el título del marcador/outline del PDF
    # (p. ej. "PARTE ICONTEXTO Y FUNDAMENTOS" en vez de "PARTE I CONTEXTO...").
    heading.extend([ns, NavigableString(" "), nm]); content.extend([kicker, heading]); opener.append(content)

    network = soup.new_tag("div", attrs={"class": "part-network"})
    network.append(BeautifulSoup(_part_network_svg(), "html.parser"))
    opener.append(network)
    axis = soup.new_tag("div", attrs={"class": "part-axis"})
    opener.append(axis)

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

    def flush(pre_break: bool = False) -> None:
        nonlocal chunk
        if not chunk:
            return
        css_class = "flow-columns flow-columns-preflush" if pre_break else "flow-columns"
        flow = soup.new_tag("div", attrs={"class": css_class})
        for x in chunk:
            flow.append(x)
        container.append(flow)
        chunk = []

    for node in direct:
        if is_full_width(node):
            # El tramo justo antes de un elemento a todo el ancho (figura/tabla)
            # se marca con una clase aparte (.flow-columns-preflush): en vez de
            # repartirse en dos columnas, se renderiza a una sola columna para
            # que el texto no "salte" a una segunda columna justo antes del
            # elemento a todo el ancho (ver la regla CSS homónima).
            flush(pre_break=True)
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
    """Busca la página donde aparece el marcador (el inicio del caption) de
    una figura.

    Se intenta primero con el marcador completo (hasta ~60 caracteres del
    caption) y solo se cae a fragmentos más cortos —terminando en el
    mínimo "Figura N."— si ninguna página contiene el fragmento largo. Un
    marcador tan corto como "Figura 6." puede coincidir con una referencia
    cruzada dentro del caption de OTRA figura anterior (p. ej. "...
    documentada en la Figura 6." al cierre del caption de la Figura 5), lo
    que hace que se detecte la página equivocada y la función deje la
    figura intacta sin motivo. Un fragmento largo del caption real es, en
    la práctica, único en el documento.
    """
    short_match = re.match(r"^Figura\s+\d+\.", marker)
    short = short_match.group(0) if short_match else marker
    candidates = [marker]
    if len(marker) > 30:
        candidates.append(marker[:30])
    if short != marker and short not in candidates:
        candidates.append(short)
    for cand in candidates:
        for pno in range(1, doc.page_count):
            if doc[pno].search_for(cand):
                return pno
    return None


def _figure_marker(fig: Tag) -> Optional[str]:
    caption = fig.find("figcaption")
    if caption is None:
        return None
    text = caption.get_text(" ", strip=True)
    if not re.match(r"^Figura\s+\d+\.", text):
        return None
    return text[:60].strip()


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
            color=(0x16 / 255, 0x8D / 255, 0xA2 / 255),  # cian profundo #168DA2
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
        description="Genera el Reporte Hito 0 con el diseño editorial maestro V6 (Red/Cian) mediante HTML/CSS -> PDF."
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
    print("GENERADOR EDITORIAL HITO 0 — HTML/CSS + WEASYPRINT — v6 (Red/Cian)")
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
