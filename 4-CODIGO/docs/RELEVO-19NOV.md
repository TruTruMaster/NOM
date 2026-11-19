# Relevo del 19 de noviembre de 2026

**Leer este primero.** Sigue a `RELEVO-18NOV-TARDE.md` (0.17.0 publicada el 18-nov a las 22:20).

Pedido vigente del dueño (18-nov, repetido el 19-nov): todo lo que `PARIDAD-OFFICE.md` marca «No» o «A medias» en
Word, Excel y PowerPoint queda en «Sí», **comprobado contra Office 2024 real en la PC de Windows** (nunca LibreOffice),
con su medida escrita en la fila. Cualquier archivo abre, se ve, se edita y se descarga al 100 %, y funciona con `.nom`,
cifrado, sincronía, sin Internet y el teléfono como servidor (`ApiTelefono` repite cada ruta nueva). Lo que no se toca
de un archivo de Office sale idéntico. Lo que de verdad no se pueda igualar no se salta: se anota aquí qué se intentó,
qué se midió y qué falta. Las filas «decisión del dueño» (p. ej. el Diseñador «sin IA») se quedan como están.

**Terminado** = `PARIDAD-OFFICE.md` sin «No» ni «A medias» (salvo decisiones del dueño o imposibles documentados), cada
«Sí» con su medida contra Office real, la regresión completa de Windows en verde y la versión publicada (0.18.0).

Orden: (0) auditar PARIDAD; (1) lo que queda de cómo se VE un archivo; (2) herramientas de Excel contra Excel real;
(3) crear lo que hoy solo se ve; (4) lo demás.

## 1. Hecho en esta sesión (commits en `main`, sin publicar todavía)

| Qué | Comprobado contra Office real |
|---|---|
| **Auditoría de PARIDAD**: «Escala de caracteres: no» y «Notas al final: no» eran viejas (hechas el 18-nov); corregidas con su medida. Párrafos de medidas de Word y PowerPoint al día | w01 137/137, w07 122/123 |
| **Renglones de Word** (`acomodarComoWord`, `modo-texto.js`): versalitas (mayúscula al 80 % del tamaño redondeado a medio punto), superíndice/subíndice (tamaño × `ySuperscriptYSize` de la fuente, redondeado a medio punto; también en pantalla, `--supk`), escala de caracteres (corta después de cada palabra con escala), corte después de una tabulación, sangrías negativas, recuadros (el borde sale al margen solo de su lado), la cita de Word sin los márgenes de la de NOM, columnas de distinto ancho (se revisa hasta que nada cambia; lo que sigue en otra columna se corta con su ancho), títulos H1-H3, celdas de tabla (con el ancho de Word de su columna: el navegador redondea los bordes a 1 px), junto a lo que flota (cada renglón con su ancho), letra capital. Arreglos: la holgura de 1 px pisaba el margen derecho de columnas, recuadros y citas; anchos con fracciones de píxel y el zoom exacto (clientWidth entero desviaba hasta 0.4 px); `text-wrap: balance/pretty` de la interfaz ya no toca el documento (partía los títulos a la mitad); mayúsculas del estilo de tabla | w01 94.9→**100**, w02 96.2→**100**, w03 96.6→**100**, w04 94.3→**100**, w05 85.7 (sin celdas)→**98.5** (con 744 renglones de celdas), w06 83.8→**99.3**, w07 96.5→**99.2**, w09 98.2→**100** (con celdas), t01 90.8 (sin celdas)→**95.4** (con celdas); w14 99.6; `prueba_acomodo` 11/11, ediciones docx 98/98, `prueba_funciones` 204/204 |
| **Fuentes con las medidas de Office**: Sitka Text (instancia de la variable), Lucida Sans Unicode, Ink Free (`armar_fuentes.py --solo` ya no borra las demás) | w01 100 % |
| **Medir mejor**: `verdad.py` saca los renglones de cada celda (Word los daba por fila); `comparar_real.py` los compara en orden de lectura, un objeto en el renglón es «/» (como lo lee Word), los renglones con ecuación no cuentan, la viñeta de una celda no es texto, el texto vertical se agrupa por columnas, un subíndice después de un superíndice no es renglón nuevo. Herramientas nuevas: `analisis/razones_acomodo.py` (por qué un párrafo no se acomoda), `analisis/inspeccionar.py` (abrir un archivo y correr JS) | — |
| **PowerPoint**: renglones donde los pone PowerPoint también en las miniaturas, al terminar de escribir en un cuadro y en el PNG de la diapositiva (que ahora lleva adentro las fuentes de la página; antes salía con las del equipo) | `analisis/png_diapositivas.py`: PNG de NOM contra PowerPoint p01 98.4 %, p07 98.9 %; prueba_pptx y pres1 sin errores |
| **Excel: Buscar objetivo** (Más → Datos, ruedita, paleta; diálogo y estado como Excel; Cancelar regresa el valor) | `datos_excel.py` 14/14 (los 4 de Buscar objetivo, también el sin solución) |
| **Excel: nombres definidos** (motor, importador v7 —ya no los sustituye: la barra muestra `=B2*IVA`—, `definedNames` al descargar, Administrador de nombres) | `libro_excel.py` (nuevo): Excel 2024 abre lo que descarga NOM con los 4 nombres y 6/6 celdas iguales; la vuelta los conserva |
| **Excel: cálculo manual, F9 y Mayús+F9** (`calcPr calcMode`; botón «Calcular») | `libro_excel.py`: NOM no recalcula sin F9; Excel abre en manual |
| **Excel: esquema** (agrupar/desagrupar filas y columnas, +/−, niveles, ocultar/mostrar detalle, borrar; importador v8) | `libro_excel.py`: mismos niveles y lo mismo contraído en Excel; validador de Microsoft 3/3; ediciones xlsx 30/30 |
| **Excel: mostrar fórmulas** (Ctrl+`; importador v9, `showFormulas`) | `libro_excel.py`: Excel abre con las fórmulas a la vista |
| **Excel: rastrear precedentes y dependientes** (flechas por niveles, Quitar flechas) | `libro_excel.py`: 7 celdas iguales a DirectPrecedents/DirectDependents |
| **Excel: evaluar fórmula** (paso a paso, en el orden del cuadro de Excel); potencias imposibles dan #¡NUM!/#¡DIV/0! (antes NaN) | `libro_excel.py`: 37 pasos de 7 fórmulas, calculados por Excel, dan lo mismo que la celda |
| **Excel: consolidar** (11 funciones; rótulos o posición; el orden de rótulos sale de medir Excel) | `datos_excel.py` 24/24 (10 de consolidar) |
| **Excel: tablas de datos** (`=TABLA(fila;columna)`: calcular, abrir —importador v10— y guardar como Excel); números en notación científica (Excel los escribe así; NOM los tomaba como texto) | x10-tablas-datos.xlsx (hecho en Excel): los 51 valores iguales; x01-x09 y t02 siguen en 100 % |

### 1a. Buscar objetivo: lo que no se pudo igualar

Excel no publica su método. Se midió (`windows/objetivo_pasos.py`, `windows/objetivo_eventos.py`): `MaxIterations` no
lo cambia, `MaxChange` sí; con x² = 2 desde 1 hace 7 cálculos y deja 1.4142112625 (f = 1.9999935, debajo de √2, así que
no es Newton puro: probablemente encierra la raíz y usa falsa posición); sin solución hace 101 cálculos. Se probaron
Newton con derivada numérica (pasos relativos y absolutos de 1e-10 a 0.1, hacia adelante y centrada), secante con 13
arranques, y Newton + falsa posición/Illinois: ninguno da sus 24 valores. No se pudo leer cada valor que prueba (el
evento SheetCalculate llega al final; una celda circular de historia rompe su búsqueda; con macros sí se podría, pero
pide cambiar la seguridad de VBA de la PC de prueba —no se hizo sin preguntar—). NOM usa Newton + Illinois con la
misma tolerancia; los dos quedan a menos de 0.001 del valor pedido y coinciden en si hay solución.

## 2. Cómo seguir

- Servidor de desarrollo: 8096 (interfaz en vivo desde `~/NOM`, datos en `~/NOM/trabajo/srv/datos-8096`).
- Mediciones de esta sesión: `~/NOM/trabajo/19nov/` (c1…c15 son corridas de `comparar_real.py --salida`; la verdad
  vieja de Word, sin celdas, en `verdad-vieja/`). La verdad nueva (con celdas) ya está en
  `~/.cache/nom-pruebas/office-real/verdad/` salvo w06 (Word falló al leer sus celdas: se dejó la vieja).
- Medir Word: `NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/office/comparar_real.py ARCHIVOS --salida DIR`
  (los 15 en ~6 min). Por qué no se acomoda un párrafo: `analisis/razones_acomodo.py ARCHIVO --reporte DIR`.
- Regresión: `python3 pruebas/regresion_windows.py --solo a,b` durante el trabajo; la completa antes de publicar.

## 3. Lo que falta (lista real, de la auditoría del 19-nov)

**Cómo se ve (punto 1)**
- Word: guiones automáticos con el diccionario de Word (w10 35 %, w13 59 %: juntar con Word real cientos de palabras
  partidas y usarlas como excepciones; acomodar los párrafos con guiones); texto a los dos lados de un dibujo, contornos
  (estrecho, a través) y la franja que sube a los párrafos de arriba del ancla (w12 86 %); tabla flotante (`tblpPr`) y
  filas de tabla entre hojas (w05); párrafos justificados o centrados con tabuladores; cuadros de texto (sus renglones
  no se miden todavía: falta sacarlos en la verdad); números de línea; el alto de un párrafo vacío con otra letra en la ¶;
  las páginas de t01 (126 en Word, 121 en NOM) y w06 (18/17).
- PowerPoint: SmartArt contra el video (p06, p11 89.5 %), transiciones y Transformar (Morph) contra el video.

**Excel (punto 2)**, lo que sigue: tablas de datos — falta el cuadro para crearlas desde NOM (el cálculo, abrir y guardar
ya están) y probar en Excel una creada en NOM —; relleno rápido (Ctrl+E: `Range.FlashFill` da la verdad), escenarios
(`Scenarios.Add`/`Show`/resumen), hoja de previsión y `FORECAST.ETS*` (el método de Excel no es público: medir como se
hizo con Buscar objetivo), `ODDFPRICE/ODDFYIELD/ODDLPRICE/ODDLYIELD`, `IMAGE`; disposición de página e impresión
(márgenes, orientación, tamaño, área de impresión, títulos, encabezado y pie, saltos, ajustar a N páginas, vistas de
saltos y de diseño; PDF de NOM contra el de Excel).
- **«####» con el hueco «_)»** (x10: 3 celdas): el formato de moneda de Excel deja al final un hueco del ancho de «)» y
  con él «$1,887.12» ya no cabe en 64 px. Se probó dibujar el hueco invisible del ancho exacto (`.xl-hueco::before`) y
  x10 subió a 98.2 %, pero x01 bajó a 97.7 % y t02 a 99.9 %: se quitó. Medido en Excel (`windows/medir_gatos.py`): el
  número cabe si a la columna le sobran más de ~3 px (entre 2.04 y 3.40). Falta ver qué celdas de x01 cambiaron con el
  hueco (probablemente se muestran «####» de más) antes de volver a ponerlo.

**Crear lo que hoy solo se ve (punto 3)** y **lo demás (punto 4)**: ver las filas «No» y «A medias» de
`PARIDAD-OFFICE.md` (§2-§5): tablas dinámicas, dar formato como tabla, estilos de celda, minigráficos, segmentaciones,
comentarios con hilo; estilos, temas, formas/SmartArt/gráficas/WordArt/cuadros de texto en Word, marcadores y
referencias cruzadas, tablas de ilustraciones y de autoridades, índice, citas de Word, combinar correspondencia de Word,
comparar, control de cambios completo, pegar con formato, controles de contenido, convertir texto↔tabla, dibujar tablas,
Excel incrustado; en PowerPoint ocultar diapositivas, presentaciones personalizadas, ensayar y grabar, secciones, patrón,
3D y estilos de forma, insertar video y audio, crear SmartArt; vistas de Word, ordenar párrafos, marcas ¶, estilos de
imagen; VBA, Power Query, Power Pivot y Python (como mínimo: que se conserven al descargar, con plan).

**Publicar**: nada de lo de este relevo está publicado (los commits están en `main` y en GitHub; no en la Pi ni en el
APK). La 0.18.0 sale cuando la regresión completa de Windows pase (`python3 pruebas/regresion_windows.py --carriles 10`).
Antes: correr `pruebas/office/libro_excel.py` y `datos_excel.py --solo-nom` (las dos nuevas de esta sesión).

## 4. Cuánto falta (al cerrar el 19-nov)

De `PARIDAD-OFFICE.md` quedan en «No» o «A medias» unas 60 filas o partes de fila: Word ~25 (estilos, temas, formas y
SmartArt en el texto, marcadores, índices, citas y correspondencia de Word, comparar, control de cambios completo,
pegar con formato, formularios, vistas, guiones con el diccionario de Word…), Excel ~15 (relleno rápido, escenarios,
previsión, disposición de página e impresión, tablas dinámicas y formato como tabla desde NOM, minigráficos,
segmentaciones, comentarios con hilo, VBA/Power Query/Python a conservar) y PowerPoint ~15 (secciones, ocultar
diapositivas, presentaciones personalizadas, ensayar y grabar, patrón, 3D, video y audio, crear SmartArt, Morph contra
el video). A este ritmo (unas 10-12 funciones medidas contra Office por sesión larga) son varias sesiones más.
