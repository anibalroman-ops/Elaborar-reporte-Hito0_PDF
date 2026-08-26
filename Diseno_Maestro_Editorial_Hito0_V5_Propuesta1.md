# Diseño Maestro Editorial Hito 0 — V5 · Propuesta 1

## Concepto general

Sistema editorial híbrido monocromático basado en la portada seleccionada: fondo blanco cálido, tipografía negra, barras diagonales redondeadas y nodos circulares en una escala de grises. El motivo geométrico representa progresión, posición, relación y medición y se utiliza con máxima intensidad en portada y portadillas, y de manera muy restringida en el cuerpo.

## Arquitectura editorial por Parte

| Parte | Línea | Función editorial |
|---|---|---|
| I. Contexto y fundamentos | 3 — Conceptual/editorial | Apertura conceptual, mayor aire y jerarquía de títulos |
| II. Primera aproximación institucional | 1 — Equilibrada/científica | Resultados, gráficos, tablas y lectura continua |
| III. Construcción inicial de los Construct Maps | 3 — Conceptual/editorial | Arquitectura conceptual y comparación de mapas |
| IV. Juicio experto y evidencia de contenido | 2 — Técnica/analítica | Métodos, matrices, V de Aiken, análisis cualitativo y triangulación |
| V. Refinamiento y reespecificación | 3 — Conceptual/editorial | Construct Maps, decisiones de refinamiento y síntesis conceptual |
| VI. Desarrollo de la nueva medición | 2 — Técnica/analítica | Items Design, Outcome Space, pilotaje, modelos y agenda de validación |
| VII. Cierre institucional | 3 — Conceptual/editorial | Síntesis, aprendizajes, recomendaciones y cierre estratégico |

## Portada

- A4 vertical.
- Fondo blanco cálido.
- Encabezado superior con regla fina.
- Título: **Diagnostico Hito 0**.
- Subtítulo: **Medición Atributos i+e**.
- Sistema geométrico: barras diagonales sólidas y nodos circulares en grises.
- Pie editorial con Facultad, Universidad y descriptor del reporte.

## Portadillas

- Comparten la geometría de la portada, ampliada y recortada.
- Número de Parte + nombre en gran escala.
- Miniíndice inferior.
- Partes con muchas secciones usan índice compacto en tres columnas para evitar cortes.
- Las Partes técnicas (IV y VI) usan una presentación más estructurada y densa.

## Línea 1 — Equilibrada / científica

- Dos columnas con medianil amplio.
- Gráficos y figuras como protagonistas.
- Títulos de sección sobrios con marcador circular.
- Tablas claras, con cabecera gris muy suave y reglas horizontales.
- Espaciado medio y lectura continua.

## Línea 2 — Técnica / analítica

- Densidad tipográfica ligeramente mayor.
- Títulos H2 con regla superior más marcada.
- H3 con marcador lateral lineal.
- Tablas más estructuradas y compactas.
- Adecuada para procedimientos, matrices, resultados técnicos y trazabilidad.

## Línea 3 — Conceptual / editorial

- Mayor espacio negativo.
- Títulos de sección más grandes.
- Figuras conceptuales y Construct Maps con mayor respiración.
- Cajas de síntesis con tratamiento sobrio.
- Adecuada para marcos conceptuales, mapas, síntesis y cierres.

## Figuras

Se conserva la clasificación del generador:

1. `inline`: una columna.
2. `hero`: ancho completo dentro de la Parte.
3. `landscape`: página A4 horizontal.
4. `fullpage`: página dedicada.

Todas se integran al sistema monocromático sin marcos pesados y con captions discretos.

## Tablas

- Sin líneas verticales dominantes.
- Cabeceras en gris claro.
- Regla superior oscura.
- Filas alternas solo cuando ayudan a la lectura.
- Tablas pequeñas permanecen en columna; las densas pasan a ancho completo según la heurística existente.

## Tipografía

- Titulares y navegación: Inter / Noto Sans.
- Texto largo: Noto Serif.
- Máximo dos familias tipográficas.

## Paleta

- Negro: `#121212`
- Grafito: `#333333`
- Gris oscuro: `#5E5E5C`
- Gris medio: `#8E8E8B`
- Gris claro: `#D8D8D5`
- Gris muy claro: `#EEEEEC`
- Papel: `#FBFAF7`

## Implementación

El diseño está implementado en `generar_reporte_hito0_editorial_v5.py`. El Markdown y el ZIP de figuras permanecen como fuentes independientes: el script aplica automáticamente la línea editorial correspondiente a cada Parte.
