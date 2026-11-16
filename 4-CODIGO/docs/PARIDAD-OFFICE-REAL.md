# Paridad con Microsoft Office de verdad (desde el 15 de noviembre de 2026)

Encargo del equipo: con una computadora con Windows y Office 2024 (Word, Excel y PowerPoint reales),
analizar todo lo que Office ofrece, armar los archivos más complejos posibles con todas sus
herramientas y lograr que NOM-COLAB (web y app) los muestre **idénticos**, sin importar la combinación.

## 1. Cómo se mide (reproducible)

La computadora con Windows (`flast-desktop-win`, Tailscale `100.65.29.5`, usuario `user`, SSH con la
llave de la laptop) tiene Office Professional Plus 2024 (16.0.17932, 64 bits) y Python 3.15 con pywin32 y
Pillow. Office solo guarda archivos si corre **dentro de la sesión de escritorio**: los programas se lanzan
con una tarea programada (`correr.ps1` → `lanzador.py`). Word y Excel abren su propia copia oculta;
PowerPoint es una sola copia para todo el equipo y nunca se cierra si alguien tiene presentaciones
abiertas. Excel y PowerPoint se abren visibles pero **fuera de la pantalla** (sin ventana no dibujan la
cuadrícula ni crean gráficas). Esa Windows tiene la cola de impresión desactivada a propósito: Excel no
deja tocar la configuración de página ni exportar a PDF (no se activó).

| Paso | Programa | Qué hace |
|---|---|---|
| Armar los archivos | `pruebas/office/windows/hacer_word.py`, `hacer_excel.py`, `hacer_powerpoint.py` | Usan las herramientas de Office por automatización; cada función queda anotada en `NOMBRE.funciones.json` (si Office la aceptó) |
| La referencia | `pruebas/office/windows/verdad.py` | Word: cada renglón de cada página (texto exacto y posición, de `Pages/Rectangles/Lines`) y el PDF de Word. Excel: el texto que muestra cada celda (con `####`), colores ya aplicados por el formato condicional, anchos y altos, gráficas exportadas por Excel y la captura de la ventana a 100 %. PowerPoint: cada forma (posición, giro, relleno, renglones de su texto) y cada diapositiva exportada a 1920×1080 |
| Manejarlo desde Linux | `pruebas/office/office_real.py enviar | correr ID script.py … | traer PATRÓN` | Copia los programas y las imágenes, corre en Windows y trae los resultados a `~/.cache/nom-pruebas/office-real/` |
| Comparar | `pruebas/office/comparar_real.py ARCHIVOS… [--movil]` | Importa cada archivo en NOM-COLAB (Chromium) y mide: Word, renglones idénticos y en su página; Excel, texto de cada celda, fondos y anchos; PowerPoint, parecido de cada diapositiva. Deja imágenes Office \| NOM-COLAB \| diferencias en `office-real/comparar/` |

Las imágenes de prueba (fotos, logotipo con transparencia, SVG, video y audio cortos) las hace
`windows/imagenes.py` en Linux, para que todos los archivos lleven las mismas.

## 2. Los 28 archivos

| Word | Excel | PowerPoint |
|---|---|---|
| w01 letra (37 fuentes, 15 tamaños, 13 subrayados, efectos, ligaduras, 12 alfabetos) | x01 formato (44 formatos de número, letra, 14 tramas, degradados, 12 bordes, alineaciones, combinadas, 29 estilos de celda) | p01 texto (viñetas, numeraciones, fuentes, efectos, autoajuste, columnas, texto vertical, WordArt, deformado) |
| w02 párrafo (alineaciones, sangrías, espacios, interlineados, bordes, sombreados, tabuladores, capitales) | x02 fórmulas (unas 150 funciones, matrices dinámicas, LAMBDA, errores, nombres, 3D, tabla de datos) | p02 las 181 formas predefinidas, ajustes, giros, grupos anidados, conectores, forma libre, combinar formas |
| w03 estilos y listas (títulos, estilos propios heredados, 7 viñetas, 7 numeraciones, 7 multinivel) | x03 formato condicional (20 tipos: barras, escalas, íconos, fórmulas…) y validación | p03 degradados, 48 tramas, rellenos de imagen y textura, líneas, sombras, resplandor, biseles, 3D, 42 estilos rápidos |
| w04 secciones (horizontal, columnas, oficio centrado, bordes de página, números de línea, pares e impares, marca de agua, A5) | x04 gráficas (29 tipos, combinada, formato detallado, estilos, minigráficos) | p04 imágenes (recortes, a forma, color, efectos artísticos, estilos, SVG) |
| w05 tablas (8 estilos, combinadas, anidadas, bordes por celda, texto girado, flotante, 80 filas con encabezado repetido) | x05 tablas, filtros, tablas dinámicas, segmentaciones | p05 tablas con estilo y gráficas |
| w06 objetos (imágenes con 6 ajustes, formas, rellenos, efectos, cuadros enlazados, WordArt, SmartArt, 8 gráficas, ecuaciones) | x06 imágenes, formas, SmartArt, notas, vínculos, controles de formulario, grupos, conectores | p06 36 diseños de SmartArt |
| w07 referencias (tabla de contenido, notas, marcadores, vínculos, títulos y referencias cruzadas, citas y bibliografía, índice, campos, controles) | x07 hojas (inmovilizadas, sin cuadrícula, zoom, derecha a izquierda, esquema, ocultas, protegidas, divididas) | p07 todos los diseños, 14 temas, fondos, pie y número |
| w08 comentarios y control de cambios de dos personas | x08 registro escolar y presupuesto | p08 transiciones, animaciones, notas, secciones, medios |
| w09 informe largo (tema, portada, índice, capítulos numerados, notas, figura, tabla, gráfica, ecuación, anexo horizontal) | | p09 presentación realista con tema |
| w10 combinaciones raras | | p10 combinaciones raras (A4 vertical, fuera de la diapositiva, grupos girados y volteados…) |

Funciones que Office no dejó automatizar (anotadas en los `.funciones.json`): agrupar formas y conectores
fuera de un lienzo en Word; comentarios con hilo en Excel (Office 2024 por volumen no los tiene); gráficas
nuevas (cascada, rectángulos, histograma…) desde PowerPoint (sí están en Excel); configuración de impresión
de Excel (sin cola de impresión).

## 3. Medida inicial (0.14.2, 15-nov-2026, computadora)

| Word | Renglones idénticos | Páginas Word / NOM |
|---|---|---|
| w01 letra | 63.5 % | 6 / 5 |
| w02 párrafo | 27.2 % | 9 / 8 |
| w03 estilos y listas | 64.2 % | 7 / 6 |
| w04 secciones | 6.1 % | 9 / 8 |
| w05 tablas | 62.9 % | 8 / 8 |
| w06 objetos | 23.6 % | 18 / 13 |
| w07 referencias | 56.1 % | 6 / 6 |
| w08 revisión | 11.4 % | 2 / 1 |
| w09 informe | 58.9 % | 9 / 8 |
| w10 raros | 16.0 % | 6 / 5 |

| Excel | Texto de celdas igual | Fondos iguales | Anchos de columna |
|---|---|---|---|
| x01 formato | 84.9 % | 26.2 % | 0 % |
| x02 fórmulas | 78.4 % | 0 % | 0 % |
| x03 condicional | 97.5 % | 0 % | 0 % |
| x04 gráficas | 100 % | — | 0 % |
| x05 tablas dinámicas | 94.1 % | 0 % | 0 % |
| x06 objetos | 100 % | — | 0 % |
| x07 hojas | 72.7 % (8 de 10 hojas; 2 ocultas) | — | 0 % |
| x08 registro | 100 % | 100 % | 0 % |

| PowerPoint | Parecido medio de la imagen (mínimo) |
|---|---|
| p01 texto | 96.2 % (87.8) |
| p02 formas | 94.1 % (88.2) |
| p03 rellenos y efectos | 86.7 % (69.4) |
| p04 imágenes | 90.7 % (42.1) |
| p05 tablas y gráficas | 92.1 % (80.8) |
| p06 SmartArt | 93.3 % (85.9) |
| p07 temas | 85.3 % (29.4) |
| p08 animaciones | 98.5 % (97.7) |
| p09 presentación | 34.8 % (29.3) |
| p10 raros | 96.8 % (92.9) |

(El parecido de imagen es indulgente: una diapositiva casi blanca saca un número alto aunque le falten
detalles. Para Word y Excel las medidas son exactas: un renglón o una celda son iguales o no.)

## 4. Lo que falta, por orden de impacto

1. **Fuentes.** Office 2024 usa Aptos (Word), Aptos Narrow (Excel) y Aptos Display (títulos); NOM-COLAB
   las pinta con Carlito, que mide distinto, así que ningún renglón corta igual. Faltan también equivalentes
   de Arial Narrow, Arial Black, Calibri Light, Georgia, Verdana, Tahoma, Segoe UI, Consolas y otras.
   Plan: fuentes libres con **las mismas medidas** (ancho de cada letra y alto de renglón) que las de
   Office, como Carlito con Calibri: se toman las medidas de la fuente de Windows y el dibujo de una fuente
   libre (licencia OFL); nunca se copia el dibujo de una fuente de Microsoft.
2. **Word, párrafo:** interlineado como lo calcula Word (múltiplo del alto natural de cada fuente, no del
   tamaño), el espacio después de 8 pt del estilo Normal de Office 2024, espacio contextual, sangrías
   negativas, tipos de subrayado, espaciado y escala de letra.
3. **Excel, cuadrícula:** alto de fila y ancho de columna como Excel (de la fuente del libro), la letra del
   libro en las celdas, celdas de texto que parecen número, fechas del sistema 1900 (0 y negativas), `####`.
4. **PowerPoint:** las 187 geometrías predefinidas de Office; temas, patrones y diseños (fondos, formas
   decorativas, fuentes y estilos de texto heredados).
5. Después: tablas, imágenes flotantes, columnas, encabezados por sección (Word); bordes, tramas,
   degradados, barras de datos, escalas e íconos (Excel); degradados, efectos, imágenes, tablas, gráficas y
   SmartArt (PowerPoint); y la vuelta: lo que exporta NOM-COLAB abierto en Office sin reparar y visto igual.

## 5. Avance

### Tanda 1 (15 y 16 de noviembre de 2026, versión 0.15.0)

**Fuentes y párrafo (Word).** 32 familias con las medidas de las de Office (avances, alto de renglón,
subrayado, interletraje y ligaduras) sobre dibujos libres OFL (`herramientas/fuentes/`); interlineado de
Word (múltiplo del alto natural de cada fuente), justificado que encoge espacios hasta 20 % como Word,
interletraje y ligaduras solo si el documento los pide.

**Excel (importador versión 3).** Formato de números completo de Excel (fracciones, colores, condiciones,
contabilidad, escalas), cuadrícula con las medidas de Excel (ancho de columna y alto de fila exactos, letra del
libro, `####`), tablas con referencias estructuradas y nombres definidos, y funciones nuevas. El tinte de los
colores del tema como lo calcula Excel (HLS de Windows con HLSMAX = 240: 37 de 37 colores iguales a Excel; los
grises redondean), los 64 colores indexados, bordes con su estilo y color, los 60 estilos de tabla y 84 de
tabla dinámica de fábrica (ECMA-376) con la precedencia de Excel, botones de filtro, y el formato condicional
completo (reglas de valor, texto, fórmula, repetidos, primeros y últimos, promedio, escalas de color con el
paso truncado de Excel, barras de datos y juegos de íconos). Lo importado antes conserva su versión: lo que se
descarga se compara con la misma lectura.

**PowerPoint.** Las 187 geometrías predefinidas de Office con sus ajustes y las formas libres; el texto con la
herencia de PowerPoint (estilos del patrón, marcadores, listas, párrafo y tramo), márgenes internos, anclaje,
autoajuste y columnas; cada patrón con su tema (colores, fuentes, líneas, rellenos de fondo) y su mapa de
colores; fondos heredados (duotono, mosaico a la resolución de la imagen, degradados con todas sus paradas);
las formas del patrón y del diseño debajo de la diapositiva; rellenos con degradado, imagen o trama (las 54
tramas de Office medidas en PowerPoint real); líneas y conectores con volteos, punteado y puntas de flecha.

| PowerPoint | Antes (mínimo) | Después (mínimo) |
|---|---|---|
| p01 texto | 96.2 % (87.8) | 97.6 % (90.2) |
| p02 formas | 94.1 % (88.2) | 99.0 % (94.5) |
| p03 rellenos y efectos | 86.7 % (69.4) | 95.2 % (89.8) |
| p04 imágenes | 90.7 % (42.1) | 98.0 % (91.7) |
| p05 tablas y gráficas | 92.1 % (80.8) | 92.1 % (80.8) |
| p06 SmartArt | 93.3 % (85.9) | 94.0 % (86.3) |
| p07 temas | 85.3 % (29.4) | 98.8 % (96.1) |
| p08 animaciones | 98.5 % (97.7) | 99.5 % (98.8) |
| p09 presentación | 34.8 % (29.3) | 96.8 % (82.3) |
| p10 raros | 96.8 % (92.9) | 97.9 % (93.2) |

| Excel | Texto de celdas | Fondos | Anchos | Altos |
|---|---|---|---|---|
| x01 formato | 84.9 → 97.3 % | 26.2 → 67.5 % | 0 → 100 % | 98.2 % |
| x02 fórmulas | 78.4 → 98.3 % | 0 → 100 % | 0 → 100 % | 100 % |
| x03 condicional | 97.5 → 100 % | 0 → 100 % | 0 → 100 % | 100 % |
| x04 gráficas | 100 % | — | 0 → 100 % | 100 % |
| x05 tablas dinámicas | 94.1 → 100 % | 0 → 100 % | 0 → 100 % | 100 % |
| x06 objetos | 100 % | — | 0 → 100 % | 100 % |
| x07 hojas | 72.7 → 100 % | — | 0 → 100 % | 100 % |
| x08 registro | 100 % | 100 % | 0 → 100 % | 100 % |

| Word | Renglones idénticos (antes → después) | Páginas Word / NOM |
|---|---|---|
| w01 letra | 63.5 → 85.4 % | 6 / 6 |
| w02 párrafo | 27.2 → 82.4 % | 9 / 9 |
| w03 estilos y listas | 64.2 → 96.6 % | 7 / 7 |
| w04 secciones | 6.1 → 38.4 % | 9 / 8 (secciones de otro tamaño u orientación y columnas: falta) |
| w05 tablas | 62.9 → 85.7 % | 8 / 9 |
| w06 objetos | 23.6 → 57.4 % | 18 / 14 (objetos flotantes: falta) |
| w07 referencias | 56.1 → 91.2 % | 6 / 6 |
| w08 revisión | 11.4 → 94.3 % | 2 / 2 |
| w09 informe | 58.9 → 77.3 % | 9 / 9 |
| w10 raros | 16.0 → 18.1 % | 6 / 6 (división automática de palabras: falta) |

**Lo que sigue (por impacto):** gráficas de Office (un solo dibujante para Excel, Word y PowerPoint: hoy la hoja
no muestra las gráficas, imágenes ni formas del archivo y PowerPoint las dibuja con el estilo de NOM-COLAB);
Word: secciones con otro tamaño u orientación, columnas, objetos flotantes (cuadros de texto, formas,
imágenes) y la división automática de palabras; Excel: tramas y degradados en celdas, bordes diagonales, texto
girado y alineación vertical, filas de alto fijo que recortan; PowerPoint: tablas con su estilo y SmartArt;
la vuelta (lo exportado abierto en Office real) y la paridad en la app de Android.
