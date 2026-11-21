# Paridad con Office: el inventario de funciones contra lo que NOM tiene

17 de noviembre de 2026. Fuente: el inventario «Microsoft Office Complete Feature Inventory» que entregó el usuario
(32 páginas: cinta de Word, Excel y PowerPoint, los 75 campos de Word, funciones de hoja, VBA, Power Query, Power
Pivot y DAX, Python en Excel y las novedades de Microsoft 365 de 2025-2026). El estado de NOM sale de la revisión
completa del código de esa noche (`RELEVO-ESTADO-Y-PENDIENTES.md`, sección 9), de las pruebas automáticas y de las
mediciones contra Office 2024 real (`PARIDAD-OFFICE-REAL.md`). Lo que no se pudo comprobar en el código dice «sin
comprobar».

**Sí**: está y funciona · **A medias**: hay una parte (se dice cuál) · **No**: falta · **Se ve**: lo que trae un
archivo de Office se muestra (y se conserva al descargar), pero no se crea ni se edita desde NOM.

## 1. Errores del inventario (no copiarlos)

- **Excel no tiene «barra de menús»** desde 2007 (solo la cinta).
- **Los tipos de datos «Cotizaciones» no se consultan con `GETPIVOTDATA`** (esa es de tablas dinámicas): se leen con
  `FIELDVALUE` o con la notación de punto (`=A2.Precio`). El inventario lo dice dos veces.
- **Categoría «Lógicas» marcada como completa (16) sin `AND` (Y), `BYROW` ni `BYCOL`**, que sí son lógicas en Excel.
- **«75 de 75 campos de Word verificados»**: la lista no trae `DISPLAYBARCODE` ni `MERGEBARCODE` (Word 2013 en
  adelante), así que el «completo» no se puede creer; hay que tomar la lista de la página de Microsoft.
- **Muchas listas están cortadas** y el propio inventario lo admite (Ingeniería después de BITRSHIFT, Búsqueda después
  de COLUMN, Matemáticas entre EXP y SUMIF, Estadística después de F.INV, Financieras sin extraer). La tabla de
  «funciones financieras» que trae es la de **DAX** (que copia las de Excel), no la de la hoja.
- **`Ctrl+Alt+P` para «ayuda de fórmulas»** no es un atajo conocido de Excel: comprobarlo antes de copiarlo.
- La lista de la cinta de cada programa dice expresamente que no está completa (remite a `WordControls.txt`,
  `ExcelControls.txt` y `PowerPointControls.txt` del repositorio OfficeDev de Microsoft): para el inventario por
  comando de la cinta hay que usar esos archivos, no este documento.
- Las fechas y números de compilación de 2026 (versión 2611, `=COPILOT()` retirada el 14-sep-2026, etc.) no se
  comprobaron; no cambian lo que NOM tiene que hacer.

## 2. Lo común a los tres programas

| Office | NOM | Cómo está en NOM |
|---|---|---|
| Cinta con pestañas (Inicio, Insertar…) y grupos | A medias | Una barra con el nombre de cada herramienta que baja de renglón, el menú «…», la **ruedita** (mantener presionado o Mayús+F10) y la **paleta de comandos** (Ctrl+K). No hay pestañas. |
| Pestañas contextuales (Formato de imagen, de tabla…) | A medias | La ruedita y los menús cambian según lo elegido (tabla, imagen, figura, celda). |
| Backstage (Archivo) | A medias | Descargar en Word, Excel, PowerPoint, PDF, HTML, Markdown, CSV, SVG, PNG o `.nom`; proteger como en Office; versiones; plantillas; copia; enlace para ver. Sin «Opciones» de cada programa. |
| Barra de acceso rápido personalizable | No | |
| Info. sobre herramientas mejorada | Sí | Nombre, para qué sirve y atajo en cada botón. |
| Key Tips (Alt + letras) | No | Atajos de teclado sí («?» los muestra). |
| Buscar en la cinta (Alt+Q) | Sí, equivalente | Paleta de comandos (Ctrl+K; Alt+/ dentro del texto). |
| Cinta simplificada o clásica | No aplica | |
| Computadora, tableta y web | Sí | Navegador (computadora y celular), app de Android, instalar como programa. No hay app de iPhone ni de escritorio. |
| Modo oscuro | A medias | Tema claro, oscuro o automático de la interfaz; la hoja de un texto sigue blanca, como en Word. |
| Copilot, modo agente, DLP de Copilot | No (decisión) | NOM no escribe por la persona: busca fuentes, cita y arma, pero el texto lo escribe ella (ayuda, «Qué es NOM»). |
| Coautoría | **Sí, y más** | Letra por letra entre 50 personas, también **sin Internet** (un teléfono o una computadora comparte), con ramas, cifrado de extremo a extremo, recibo de proceso y quién escribió qué. |
| Formatos Open XML (ISO/IEC 29500) | Sí | Abrir y guardar `.docx`, `.xlsx` y `.pptx` (también con contraseña de Office): lo que no se toca sale idéntico y lo editado se cambia sobre el original; medido contra Office 2024. |
| `.doc`, `.xls`, `.ppt`, `.odt`, `.ods`, `.odp`, `.rtf` | No | Se pueden elegir al importar, pero solo dan un aviso. |
| Archivos con macros (`.docm`, `.xlsm`, `.pptm`) y VBA | No | Ni se abren, ni se conservan, ni corren. |
| Proteger (restringir edición, final, contraseña) | Sí | Como en Office (`a9e388a`, `1b0d8ea`). |
| Ortografía | A medias | El corrector del navegador. Sin el «Editor» de Word (gramática, estilo), sinónimos ni traductor. |
| Comentarios | Sí | Con respuestas, resueltos y @menciones, en los cuatro editores. Los de Excel con hilo (365) no se ven. |
| Accesibilidad | Sí | Teclado, lector de pantalla, alto contraste, reducir movimiento, texto alternativo. **Comprobar
accesibilidad** (21-nov: Archivo y la paleta, en texto, hoja y presentación) con las reglas del comprobador de Office:
errores (falta texto alternativo, tabla sin encabezado —en Word, «repetir como fila de encabezado»—, diapositiva sin
título), advertencias (contraste menor a 4.5 a 1 —3 a 1 en letra grande—, vínculos «haz clic aquí», celdas combinadas,
objetos que no van en línea con el texto, renglones o espacios en blanco para acomodar) y consejos (nivel de título
saltado, diapositivas con el mismo título, «Hoja1»/«Sheet1»), con «Ir» a cada uno; `prueba_accesibilidad.py` 7 de 7 con
un Word, un Excel y un PowerPoint hechos a propósito. Office no deja leer su comprobador desde fuera (no hay API): se
comparó con sus reglas publicadas. De paso: el texto alternativo de las imágenes de Word y Excel ya se lee al abrirlos. |
| Imprimir y PDF | Sí | Desde el navegador. |

## 3. Word

| Pestaña / grupo | Office | NOM |
|---|---|---|
| Inicio · Portapapeles | Cortar, copiar, pegar, pegado especial, copiar formato | Sí, salvo **pegar con formato** (a medias: negritas, listas y vínculos; faltan estilos, tablas, imágenes y colores desde Word o Google Docs) |
| Inicio · Fuente | Fuente, tamaño, negrita, cursiva, subrayado, tachado, sub/superíndice, color, resaltado, mayúsculas, borrar formato | Sí. Versalitas, espaciado entre letras, elevado/bajado y texto oculto: se leen y se conservan de Word, sin botón propio. **Escala de caracteres (`w:w`, 50-200 %): sí** (18-nov). Medido en Word 2024 (19-nov, `windows/medir_letras.py`, `analisis/letras_word.py`): la versalita es la mayúscula al 80 % del tamaño redondeado a medio punto; el superíndice y el subíndice, el tamaño × `ySuperscriptYSize` de la fuente redondeado a medio punto; la escala multiplica el ancho; NOM corta los renglones con esas reglas: **w01 (letra) 137 de 137 renglones idénticos a Word** |
| Inicio · Párrafo | Alineación, sangrías, interlineado, espacio, viñetas, numeración multinivel, ordenar, marcas ¶, bordes y sombreado | Sí lo principal (también numeración como la escribe Word); bordes y sombreado se ven; **ordenar: sí** (21-nov: Más → Formato → Ordenar párrafos…: por texto, número o fecha,
ascendente o descendente, con encabezado y distinguir mayúsculas; cada párrafo se lleva su formato; medido en Word 2024:
signos antes que números y números antes que letras, «ñ» después de «n», el guion y el apóstrofo no cuentan, el espacio
del final tampoco, mayúsculas solo si se pide; `ordenar_real.py`: 15 de 15 casos iguales a `Range.Sort` de Word —también
en inglés y fechas día/mes/año—, Word lee lo descargado en el mismo orden y Ctrl+Z lo regresa; fuera: números con «+»,
«%» o entre paréntesis); **mostrar marcas ¶: sí** (21-nov: Más → Formato → Mostrar marcas de formato: ¶ al final de
cada párrafo, · en cada espacio, → en cada tabulación, ° en el espacio de no separación, ↵ en el salto de renglón y el
texto oculto con subrayado punteado, como Word; no ocupan lugar —con w09, sus 262 renglones en el mismo lugar— y no se
imprimen; `prueba_marcas.py` 8 de 8). **Tabulaciones del archivo: sí** (18-nov: izquierda, derecha, centro y decimal con relleno de puntos, guiones o raya, las del estilo y las propias, «quitar» y la distancia de siempre; medido en w07 y w09: índices y tablas de contenido con sus puntos y el número en la parada; w09 98.2 % de renglones idénticos); **ponerlas desde NOM: sí** (21-nov: Más → Formato → Tabulaciones…, como el cuadro
de Word: posición en cm, izquierda, centro, derecha, decimal o barra, relleno de puntos, guiones, subrayado o punto medio,
eliminar y eliminar todas, y las predeterminadas del documento; `tabulaciones_real.py` 5 de 5: Word lee las mismas
paradas en los 9 párrafos y las mismas predeterminadas, y NOM pinta el documento con el 100 % de los renglones como Word;
la regla para arrastrarlas: no) |
| Inicio · Estilos | Galería de estilos, crear y modificar | A medias: Normal, Título 1-3 y los del archivo se ven; crear o cambiar estilos: no |
| Inicio · Edición | Buscar, reemplazar, seleccionar | Sí |
| Insertar · Páginas | Portada, página en blanco, salto | Salto sí; portada de la galería no («Armar reporte» hace la suya) |
| Insertar · Tablas | Insertar, dibujar, convertir, Excel incrustado | Insertar, combinar, bordes, rellenos, encabezado, anchos: sí; estilos de tabla de Word se ven; dibujar, convertir texto↔tabla y Excel incrustado: no |
| Insertar · Ilustraciones | Imágenes, formas, iconos, 3D, SmartArt, gráfico, captura | Imágenes sí (recortar, girar, texto alrededor, pie, texto alternativo). Formas, SmartArt y gráficos **se ven**; insertarlos en un texto: no. Iconos, modelos 3D y captura: no. **EMF/WMF: sí** (18-nov: `js/emf.js`; 535 metarchivos del corpus contra cómo los muestra Word real, `metarchivos_real.py`: parecido 96.2 %, 88.5 % en lo dibujado; las ecuaciones viejas y los objetos incrustados se ven con su vista previa; falta: las letras dentro del dibujo usan las del equipo) |
| Insertar · Vínculos y comentarios | Vínculo, marcador, referencia cruzada, comentario | Vínculo y comentario sí; marcador y referencia cruzada no |
| Insertar · Encabezado y pie | Encabezado, pie, número de página | Sí (primera página distinta, «1 de 5», por sección al importar) |
| Insertar · Texto | Cuadro de texto, WordArt, letra capital, firma, fecha, objeto | Fecha (fija) e insertar el contenido de otro Word: sí. Cuadros de texto y letra capital **se ven**; WordArt, firma y objetos: no |
| Insertar · Símbolos | Ecuación, símbolo | Sí (ecuaciones con notación tipo LaTeX, a OMML de Word) |
| Dibujar | Pluma, entintado | No en el texto (la pluma existe al presentar) |
| Diseño | Temas, conjuntos de estilos, colores, fuentes, marca de agua, color y bordes de página | Marca de agua sí; color y bordes de página se ven; temas, conjuntos de estilos, colores y fuentes del documento: no |
| Disposición | Márgenes, orientación, tamaño, columnas, saltos, números de línea, guiones, sangría, espacio, organizar | Márgenes, orientación, tamaño, saltos de página, sangría y espacio: sí. Columnas: las del archivo se ven como Word (18-nov: las de una sección continua se equilibran como Word; el índice a dos columnas de w07 ya no se encima); en NOM, «dos lado a lado» como bloque. Saltos de sección: se ven. **División automática de palabras: sí, como Word** (20-nov: el diccionario es el de Word 2024 —83 158 palabras preguntadas a Word real en español e inglés, cada una con los dos idiomas, aprendidas como patrones con patgen: 100 % y 99.99 % iguales a Word, 95.8 % y 78.4 % en palabras que no se le enseñaron; `herramientas/guiones/`—; cada párrafo con el idioma de su texto (`w:lang`); la regla de cuándo parte Word, medida en 3284 renglones —`windows/hacer_guiones2.py`, `analisis/guiones_decision.py`—: 99.4 % de las decisiones a la izquierda y 99.5 % en justificado; la zona de división no cambia nada en Word 2013 y después; «Limitar guiones consecutivos», «no dividir en mayúsculas» y «No dividir con guiones» por párrafo; lo que no cabe ni en un renglón entero se parte como Word —en un guion o en la última letra—. **w13-guiones 58.9 → 100 %** de renglones idénticos, w10 35.1 → 90.4 %, w15-guiones (976 renglones a la izquierda y justificados) 97.1 % y 99.1 %). **Números de línea: sí** (20-nov: se ven como Word —cada renglón del cuerpo, también los vacíos; no las celdas, cuadros de texto, notas, encabezado ni pie, ni el párrafo vacío del salto de sección; «Suprimir números de línea» no cuenta; los múltiplos de «cada»; a la izquierda de cada columna, a 0.25" o 0.125" con varias columnas si es automática; en la línea base; cada página, sección o continuos—, se ponen en Más y en la paleta —ninguno, continuos, cada página, cada sección, suprimir en el párrafo, opciones—, y se guardan en cada sección al descargar. `numlineas_real.py` contra Word 2024: w16-numlineas **92 de 92** números en su hoja, su renglón y su lugar; cambiados en NOM, descargados y abiertos en Word: las mismas opciones en las 4 secciones y **60 de 60** iguales) |
| Referencias | Tabla de contenido, notas al pie y al final, citas y bibliografía, títulos, tabla de ilustraciones, índice, tabla de autoridades | Tabla de contenido y notas al pie: sí. **Notas al final: sí** (18-nov: numeradas aparte —i, ii, iii o el formato del archivo—, después del último renglón con la raya corta de Word, también con secciones; w07 contra Word 2024: 122 de 123 renglones idénticos). Citas y bibliografía: **las de NOM** (APA, fuentes por DOI, ISBN o dirección); las de Word (`CITATION`, `Sources.xml`): no. Títulos «Figura 1.»: sí. Tabla de ilustraciones, índice (`XE`) y tabla de autoridades: no. Investigar (Wikipedia, libros, artículos): sí, extra |
| Correspondencia | Combinar correspondencia, sobres, etiquetas, reglas | **La de NOM** (texto con `{{campos}}` + una hoja = un documento por fila); la de Word (`MERGEFIELD`, origen de datos), sobres y etiquetas: no |
| Revisar | Ortografía, Editor, contar palabras, accesibilidad, traducir, idioma, comentarios, control de cambios, comparar, proteger | Contar palabras, comentarios y proteger: sí. **Control de cambios: sí** (21-nov: los cambios marcados de un Word —agregar, borrar y mover— entran como sugerencias
de NOM con su autor y su fecha, y las sugerencias salen como cambios marcados de Word; «Control de cambios» del documento
entra y sale y con él se empieza sugiriendo; Mostrar para revisión como Word —marcas simples, lo de siempre en Word,
todas las marcas, sin marcas y original—; `cambios_real.py`: 18 de 18 contra Word 2024 —w18-cambios.docx hecho en Word con
dos personas: los 6 cambios de texto iguales, descargado sin tocar Word ve los mismos 8, todo aceptado o rechazado en NOM
da el texto de Word con todo aceptado o rechazado, uno aceptado, otro rechazado y uno nuevo: Word ve los que quedan y el
nuevo con el nombre de quien lo escribió; uno hecho en NOM, sus sugerencias como cambios marcados; el validador de
Microsoft sin errores—; `comparar_real.py`: w08 100 %, w18 100 %, t01 igual. Los cambios de formato y los de dentro de
una tabla se ven aplicados y se conservan sin tocar al descargar; la marca de párrafo agregada o borrada, también). Comparar, traducir, idioma y Editor: no |
| Vista | Lectura, diseño de impresión, web, esquema, borrador, regla, cuadrícula, panel de navegación, zoom, ventanas, macros | Páginas o continuo, panel de esquema (títulos), zoom, modo enfoque: sí. Lectura, borrador, regla, cuadrícula y ventanas: no. Macros: no |
| Formato de imagen (contextual) | Recortar, ajustar, estilos, ajuste del texto, girar, tamaño | Recortar, girar, ajuste del texto, tamaño, texto alternativo: sí; estilos y correcciones de imagen: no (los filtros están en presentaciones) |
| Campos (los «75») | Insertar, actualizar, ver código | **Sí** (18-nov, `js/campos.js`): como Word, se ve el último valor (fecha y hora al día) y **Actualizar campos (F9)** calcula `DATE`/`TIME`/`CREATEDATE`/`SAVEDATE` con `\@`, `SEQ` (`\c \r \h \s`), `REF`, `PAGEREF` (`\p`), `PAGE`, `NUMPAGES`, `IF`, `=` (fórmulas y funciones, `\#`), `\*` (ROMAN, ALPHABETIC, CardText, OrdText, Ordinal, DollarText, Upper, Lower, Caps, FirstCap), `AUTHOR`, `TITLE`, `SUBJECT`, `DOCPROPERTY`, `STYLEREF`, `FILENAME`, `NUMWORDS`; **55 de 55 iguales a Word 2024** (`campos_real.py`). Ver códigos (Alt+F9) e insertar campo: sí. Falta: `MERGEFIELD` con el origen de datos de Word, `NOTEREF`, `TOC` actualizable, índice (`XE`/`INDEX`) |
| Controles de contenido y formularios | Casilla, lista, fecha, texto | Se conservan y su texto se edita; como formulario: no |
| VBA, modo seguro, automatización | | No |

Medido contra Word 2024 real (renglones idénticos, 20-nov noche; la verdad de Word trae también los renglones de cada
celda de tabla y `comparar_real.py` los compara: un objeto en el renglón cuenta como «/», como lo lee Word, y los
renglones con una ecuación no cuentan): **w01 100 %** (137), **w02 100 %** (239), **w03 100 %** (176), **w04 100 %**
(229), **w05 98.5 %** (744, con las celdas), **w06 99.3 %** (140), w07 98.4 % (121 de 123; uno de los dos es un campo DATE:
la verdad de Word es del 19-nov y NOM, como Word al abrirlo, pone la fecha de hoy), **w08 100 %**, **w09 100 %** (282, con las celdas), w10 90.4 %, **w11 100 %**, w12 90.7 % (21-nov; antes 86.1),
**w13 100 %**, **w14 99.6 %** (1064), t01 96.0 % (2642, con las celdas; antes 95.6), w17-dos-lados 92.2 % (886; 67.6 con el texto de un solo lado), w15-guiones-z18 97.1 % (971), **w15-guiones-lim2
99.1 %** (969). Al empezar el 19-nov: w01 94.9, w02 96.2, w03 96.6, w04 94.3, w05 85.7 (sin celdas), w06 83.8, w07 96.5,
w09 98.2, w10 35.1, w13 58.9, t01 90.8 (sin celdas). Lo que falta está en `RELEVO-20NOV-NOCHE.md` §3 (w12: texto a los dos
lados de un dibujo y contornos; w10: una palabra de 96 letras con su idioma inglés y la dirección web que sigue; t01: lo
de todos). Desde el 20-nov `comparar_real.py` separa los renglones de NOM por sus cortes (dos renglones que se enciman
con interlineado exacto ya no se juntan) y compara el texto de Wingdings/Symbol con las letras de siempre (Word lo da en
el área privada U+F0xx).

**Cómo corta Word un renglón** (medido con w14: `windows/hacer_renglones.py`, `analisis/acomodo_word.py`): a la
izquierda, todo lo que cabe con los anchos EXACTOS de la fuente, sin tolerancia; justificado (Word 2013 y después),
mete la palabra que no cabe encogiendo los espacios solo si se pasa del margen a lo más la mitad del hueco que
quedaría sin ella y cada espacio se encoge a lo más un 25 % (0 errores en 559 decisiones); corta también después de un
guion o una raya pegados a lo que sigue. NOM lo calcula con los anchos de sus fuentes (las mismas medidas que las de
Office) y lo fuerza con cortes de pantalla (`acomodarComoWord` en `modo-texto.js`; el navegador solo no puede: mide el
texto hasta un 0.1 % más angosto). Desde el 19-nov también: títulos, celdas de tabla (con el ancho de Word de su
columna), versalitas, escala de caracteres, superíndices, sangrías negativas, recuadros (el borde sale al margen solo de
su lado), columnas de distinto ancho, letra capital y junto a un objeto que flota (cada renglón con el ancho que le deja).
Desde el 20-nov también los guiones automáticos (ver Disposición). **Desde el 21-nov, texto a los dos lados de un dibujo**
(«Cuadrado, a ambos lados», lo de siempre en Word; medido con `windows/hacer_dos_lados.py`: 41 dibujos de 1, 2 y 3
pulgadas en varios lugares de la columna): cada renglón de la franja en dos tramos, primero el de la izquierda y luego
el de la derecha, y un tramo donde no cabe la palabra que sigue se queda vacío (con un hueco de pantalla, `.wd-hueco`).
Falta: párrafos justificados o centrados con tabuladores, contornos (estrecho, a través) y los cuadros de texto.

## 4. Excel

| Pestaña / grupo | Office | NOM |
|---|---|---|
| Hoja | Barra de fórmulas, filas, columnas, cuadrícula, tablas, tablas dinámicas, controles | Barra de fórmulas, filas y columnas (insertar, borrar, ancho, alto), hojas (nueva, renombrar, mover, duplicar, eliminar): sí. Tablas y tablas dinámicas **se ven** (100 % contra Excel); crearlas: no. Controles de formulario se ven con su estado |
| Teclas | F2, Entrar, Esc, Ctrl+Fin, Ctrl+D, Ctrl+R, Ctrl+Mayús+U | Sí lo principal (F2, Entrar, Esc, flechas, Ctrl+D/R, copiar fórmulas con referencias relativas, cuadrito de relleno) |
| Cálculo | Automático, automático salvo tablas, manual, F9 | **Sí** (19-nov): automático, automático excepto tablas de datos y manual (las fórmulas se quedan con su valor hasta F9 —libro— o Mayús+F9 —hoja—, con el botón «Calcular» de Excel); `calcMode` se lee y se escribe. `libro_excel.py`: NOM no recalcula sin F9 y Excel 2024 abre en manual lo que descarga NOM, con los mismos valores |
| Inicio · Fuente y alineación | Fuente, tamaño, negrita, color, relleno, bordes, alinear, ajustar texto, combinar, girar | Sí (también tramas y degradados de Excel, bordes diagonales, texto girado y apilado, sangría, reducir hasta ajustar) |
| Inicio · Número | Formatos de número, decimales, moneda, fecha, personalizado | Sí (código «como en Excel») |
| Inicio · Estilos | Formato condicional, dar formato como tabla, estilos de celda | Formato condicional: sí (también los del archivo). Dar formato como tabla y estilos de celda: se ven; crearlos: no |
| Inicio · Celdas | Insertar, eliminar, formato (alto, ancho, ocultar, proteger) | Sí; proteger hoja con sus permisos: sí |
| Inicio · Edición | Autosuma, rellenar, borrar, ordenar y filtrar, buscar y seleccionar | Sí (orden y filtro **solo para mí** o para todos: extra) |
| Insertar | Tabla dinámica, gráficas, minigráficos, imágenes, formas, cuadro de texto, vínculo, casilla, símbolos | Gráficas: sí (las de NOM; las de Excel se ven en su lugar, también 3D, cascada, histograma, Pareto, embudo, rectángulos, solar, cajas). Imágenes, formas, SmartArt y WordArt del archivo **se ven**; insertarlos: no. Tabla dinámica, minigráficos, segmentaciones y casilla en celda: no |
| Disposición de página | Temas, márgenes, orientación, área de impresión, saltos, títulos | No |
| Fórmulas | Biblioteca, nombres, auditoría, opciones de cálculo | Biblioteca y sugerencias al escribir: sí. **Nombres definidos: sí** (19-nov: Administrador de nombres y Definir nombre; la fórmula conserva el nombre —=B2*IVA, como Excel—; rango, número, fórmula o LAMBDA llamada como función; al descargar van en `definedNames`. `libro_excel.py`: Excel 2024 abre lo que descarga NOM con los mismos 4 nombres y 6 de 6 celdas iguales, y la vuelta los conserva). **Rastrear precedentes y dependientes, evaluar fórmula y mostrar fórmulas: sí** (19-nov: flechas por niveles y Quitar flechas; Evaluar fórmula paso a paso con la parte siguiente subrayada; Ctrl+`. `libro_excel.py`: precedentes y dependientes de 7 celdas iguales a DirectPrecedents/DirectDependents de Excel 2024; los 37 pasos de 7 fórmulas, calculados por Excel, dan lo mismo que la celda; Excel abre con las fórmulas a la vista) |
| Datos | Obtener datos (Power Query), ordenar, filtrar, texto en columnas, relleno rápido, quitar duplicados, validación, consolidar, Y si (buscar objetivo, escenarios, tablas de datos), previsión, agrupar, tipos de datos | Ordenar, filtrar y validación (lista desplegable): sí. **Quitar duplicados y texto en columnas: sí** (18-nov; `datos_excel.py`: 10 de 10 casos iguales a Excel 2024, con lo que se escribe en una celda entendido como Excel: 1e3, (7), fechas con guiones, formatos de moneda/porcentaje/miles). «Hacer un reporte con estos datos» y datos públicos de Wikidata: extra. **Buscar objetivo: sí** (19-nov: Datos en Más, la ruedita y la paleta, con el diálogo y el estado de Excel —Aceptar o Cancelar—; `datos_excel.py`: 4 de 4 casos como Excel 2024, también el sin solución, dentro de la tolerancia de Excel —su método no es público y su valor difiere en la sexta cifra; ver RELEVO-19NOV—). **Agrupar y esquema: sí** (19-nov: agrupar y desagrupar filas y columnas hasta 7 niveles, +/−, botones de nivel, ocultar y mostrar detalle, borrar esquema; Mayús+Alt+→/←; `libro_excel.py`: Excel 2024 ve los mismos niveles y lo mismo contraído, en filas y columnas). **Consolidar: sí** (19-nov: las 11 funciones, rótulos en fila, columna o los dos sin distinguir mayúsculas, o por posición, de esta hoja o de otras; el orden de los rótulos sale de medir Excel —un rótulo nuevo va antes del siguiente que ya estaba—; `datos_excel.py`: 10 de 10 casos iguales a Excel 2024). **Tablas de datos: sí** (19-nov: se calculan, se abren y se guardan como Excel —x10-tablas-datos.xlsx hecho en Excel 2024: los 51
valores iguales—; 21-nov: se crean desde NOM con Datos → Tabla de datos… en Más, la ruedita y la paleta, como Excel —el rango
completo y la celda de entrada de fila, de columna o las dos—; `datos_excel.py`: 3 de 3 iguales a `Range.Table` de Excel 2024
—una variable en columna, una en fila y dos variables—; `libro_excel.py`: Excel abre lo que descarga NOM con `=TABLE(B1,B2)` en
cada celda del interior y los mismos valores). **Escenarios: sí** (21-nov: el Administrador de escenarios en Más, la ruedita y
la paleta —agregar, modificar, eliminar, mostrar y resumen—, como Excel: hasta 32 celdas cambiantes sin fórmula, el comentario
«Creado por … el …», se leen y se guardan en el `<scenarios>` de la hoja; el informe resumen en una hoja nueva antes de la de
los escenarios, con los valores de cada uno, el nombre definido en vez de $B$1 y el formato de Excel —morado, gris, rayas,
grupos y sin cuadrícula—; `escenarios_real.py`: 19 de 19 contra Excel 2024 —x11-escenarios.xlsx hecho en Excel: los 4
escenarios, «Mostrar», los 30 números del resumen y sus 259 propiedades de formato iguales; cambiados o hechos en NOM, Excel
lee los mismos y su resumen da lo mismo—; `comparar_real.py`: x11 con el 100 % del texto y los fondos). **Relleno rápido:
sí** (21-nov: Ctrl+E, Más → Datos, la ruedita y la paleta; de uno o más ejemplos aprende cómo se arma cada celda con las
demás columnas de su fila —pedazos, palabras contadas desde el principio o el final, signos, texto fijo, MAYÚSCULAS,
minúsculas y Nombre Propio— y llena las vacías como si se escribieran; lee cada celda por su valor como Excel —escrito
«12.50» es «12.5»—, no toma un encabezado como ejemplo y en una fila que no sigue el patrón usa el siguiente programa que
sí, como Excel; `js/relleno-rapido.js`; `datos_excel.py`: 31 de 31 casos contra `Range.FlashFill` de Excel 2024 —uno a
propósito distinto: de «aNA péREZ» → «Ana Pérez», Excel da «Aluis Pgómez» y NOM «Luis Gómez»— y Ctrl+E/Ctrl+Z con el
teclado). Previsión: no |
| Revisar | Ortografía, notas, comentarios, proteger hoja y libro | Notas del archivo se ven; comentarios de NOM y proteger: sí |
| Vista | Saltos de página, diseño de página, cuadrícula, encabezados, zoom, inmovilizar, macros, celda de enfoque, modo oscuro | Inmovilizar (fijar filas y columnas): sí; cuadrícula como la del archivo. **Celda de enfoque: sí** (21-nov: Más → Fórmulas →
Celda de enfoque y la paleta: la fila y la columna de la celda elegida resaltadas, siguen a la selección, no cambian el
libro ni se imprimen; `prueba_enfoque.py` 6 de 6). Vista de saltos y de diseño de página (falta medir la impresión de Excel:
ver RELEVO-21NOV §4) y macros: no |
| VBA, Python en Excel, Power Query (M), Power Pivot y DAX, cubos | | No |

**Funciones.** NOM tiene **458** (con sus nombres en español y en inglés), con matrices dinámicas, `LET` y `LAMBDA`
(18-nov: +224; 20-nov: las 4 de bonos con período irregular). La lista de Excel 2024 sale del propio Excel (la biblioteca de tipos de `WorksheetFunction`, 416
funciones, `windows/lista_funciones.py`): de esas faltan solo las de idiomas de doble byte o tailandés (`BAHTTEXT`,
`THAI*`, `PHONETIC`), las que piden un servicio externo (`RTD`, `WEBSERVICE`, `FILTERXML`, `STOCKHISTORY`, `FIELDVALUE`),
`FORECAST.ETS*` (4) e `IMAGE`. **Bonos con el primer o el último período irregular** (`PRECIO.PER.IRREGULAR.1`/`.2`,
`RENDTO.PER.IRREGULAR.1`/`.2` = `ODDFPRICE`, `ODDLPRICE`, `ODDFYIELD`, `ODDLYIELD`; 20-nov): **66 de 75** casos iguales a
Excel 2024 (los cuatro ejemplos de la documentación de Microsoft, períodos cortos y largos, las cinco bases y las tres
frecuencias). Medido en Excel: en el último período no se aplica la regla de fin de mes y el largo normal de un cuasi
cupón es el real con las bases 1, 2 y 3; en el primero, los cuasi cupones entre la liquidación y el primer cupón se
cuentan con pasos desde la liquidación, y si el vencimiento no cae en el calendario del primer cupón da `#¡NUM!`. Los 9
que no: el último interés a fin de mes con cupones trimestrales o base real, y real/real con primer período largo y
primer cupón a fin de mes (`CONOCIDAS` de `funciones_excel.py`).
Comprobadas contra **Excel 2024 real** (`funciones_excel.py`: Excel escribe cada fórmula con `Formula2`, calcula y
devuelve el valor o el error exacto): **565 de 575** fórmulas iguales (20-nov; antes 499 de 500). La que no: `BESSELJ` con x ≥ 8 (Excel se aparta
del valor exacto en ~5e-7 y no se encontró su método; NOM da el exacto). Las funciones iterativas (`TIR`, `TIR.NO.PER`,
`TASA`) se comparan con la precisión que Excel documenta (1e-8).
No existen en Excel 2024 (son de Microsoft 365; dan `#NAME?` en la PC de prueba, así que no hay contra qué compararlas):
`REGEXTEST`, `REGEXEXTRACT`, `REGEXREPLACE`, `GROUPBY`, `PIVOTBY`, `TRIMRANGE`, `PERCENTOF`, `TRANSLATE`, `DETECTLANGUAGE`.

Medido contra Excel 2024 real: **texto de las celdas 100 % en los once libros** (x01 a x10 y t02, 21-nov: también x10, que
tenía tres «####» de más con la moneda «_)»); altos de fila 100 % salvo una fila de x01 (99.4 %); cada gráfica 92-98 % de
parecido. **Cuándo sale «####»** (21-nov, `windows/medir_gatos2.py`: 88 números con y sin «_)», la columna de píxel en
píxel, en libros con Calibri y con Aptos Narrow): Excel suma los anchos enteros de Windows de cada carácter —a 11 puntos,
los mismos con las dos fuentes: dígito 7, «.» y «,» 4, «-» 5, «$» 7, paréntesis 5, espacio 3—, el hueco «_)» cuenta como
un espacio, el relleno «*» no ocupa, y el número cabe si esa suma más 5 px no pasa del ancho de la columna; la fuente de
la celda no cambia nada (`js/anchos-excel.js`; con otros caracteres o tamaños, lo que mide el navegador).

## 5. PowerPoint

| Pestaña / grupo | Office | NOM |
|---|---|---|
| Inicio | Nueva diapositiva, diseños, secciones, fuente, párrafo, dibujo, organizar, editar | Diapositivas (nueva, duplicar, mover, eliminar), diseños, texto con formato por partes, viñetas, figuras, ordenar, alinear, agrupar (grupos dentro de grupos), capas: sí. **Secciones: sí** (21-nov: agregar sección en una diapositiva —sin secciones, lo de antes queda en «Sección predeterminada», como PowerPoint—, cambiar nombre, contraer y abrir —también todas—, mover arriba o abajo, quitar sección o con sus diapositivas; en la tira de miniaturas, en el menú, la ruedita y la paleta; se leen y se escriben como `p14:sectionLst`, y al borrar, agregar o mover diapositivas se rehacen con las de ahora —antes, borrar una diapositiva quitaba las secciones del archivo—. `presentacion_real.py` contra PowerPoint 2024: las 3 secciones de p13 iguales; cambiadas en NOM y hechas en NOM, iguales a lo esperado) |
| Insertar | Tabla, imágenes, captura, álbum, formas, iconos, 3D, SmartArt, gráfica, zoom, vínculo, acción, comentario, cuadro de texto, encabezado, WordArt, fecha, número, objeto, ecuación, símbolo, video, audio, grabación, Cameo | Tabla, imagen (recortar, foto dentro de una forma, filtros), formas de Office, iconos, gráfica, cuadro de texto, ecuación, vínculos y **botones de acción**, comentario: sí. SmartArt **se ve** (con o sin su dibujo) y se anima por partes; crearlo: no. **Video y audio del .pptx se reproducen al presentar**; insertarlos desde NOM: no (plan de medios, fases 2-3). Captura, álbum, 3D, zoom, WordArt, grabación y Cameo: no |
| Diseño | Temas, variantes, tamaño, formato del fondo, Diseñador | Temas, tamaño (16:9, 4:3, carta, cartel), fondo (color, degradado, imagen): sí. Diseñador e ideas de diseño: no (decisión: sin IA) |
| Transiciones | Galería, opciones, duración, sonido, avanzar | Sí: panel con **51 transiciones y 151 variantes** de PowerPoint, duración, avanzar al hacer clic o después de, sonido del archivo. **Transformar (Morph): a medias** (sin medir contra el video) |
| Animaciones | Galería, opciones, intervalos, panel, desencadenadores, copiar animación | Sí: **262 efectos** (entrada, énfasis, salida, trayectorias), opciones, intervalos, panel con línea de tiempo, trayectorias dibujadas, copiar animación, por párrafo/letra/palabra, gráficas por serie o categoría; medido 97 % contra el video de PowerPoint. Desencadenador por marcador de un medio: no |
| Presentación | Desde el principio, desde la actual, presentar en línea, personalizada, configurar, ocultar diapositiva, ensayar, grabar, subtítulos, Speaker Coach | Modo presentador (notas, siguiente, tiempo), **presentar en otra pantalla con un código** (también desde el teléfono), puntero láser y pluma: sí. **Desde el principio (F5) y desde la diapositiva actual (Mayús+F5), ocultar diapositiva, presentaciones personalizadas, ensayar intervalos y configurar presentación: sí** (21-nov: la oculta se ve atenuada y con su número tachado, no sale al presentar salvo si se empieza en ella o un vínculo lleva ahí; las personalizadas se hacen, copian, editan —qué diapositivas y en qué orden— y se presentan; ensayar mide cuánto se queda cada diapositiva, con pausa, y si se conserva queda como avance automático; configurar: repetir hasta Esc, usar los intervalos, sin animación y qué muestra «desde el principio» —todas, desde-hasta o una personalizada—; el tipo «examinada en exposición» o «en ventana» del archivo se conserva. Se leen y se escriben como PowerPoint: `show="0"`, `p:custShowLst`, `advTm` y `presProps.xml`. `presentacion_real.py` contra PowerPoint 2024 real: **31 de 31** —p13 hecho en PowerPoint (6 diapositivas, una oculta, 3 secciones, 2 personalizadas, 2 intervalos, repetir) leído igual; descargado sin tocar, PowerPoint dice lo mismo; cambiado en NOM y hecho en NOM desde cero, lo esperado; los tres abren sin avisos y el validador de Microsoft no da errores—; `prueba_presentacion.py` 10/10 por la interfaz). Presentar en línea, grabar, subtítulos en vivo y Speaker Coach: no |
| Revisar | Ortografía, accesibilidad, comentarios, comparar | Comentarios y **comprobar accesibilidad** (21-nov, ver la fila común): sí; ortografía: la del navegador; comparar: no |
| Vista | Normal, clasificador, página de notas, lectura, patrón, guías, regla, zoom | Normal, notas del orador, zoom, cuadrícula con ajuste: sí. **Patrón de diapositivas: se ve** (formas del patrón y del diseño); editarlo: no |
| Formato de forma (contextual) | Relleno, contorno, efectos, estilos, WordArt, tamaño, efectos 3D | Relleno (también degradado, trama, imagen), contorno, puntas, tamaño, giro: sí. Efectos 3D y estilos de forma de Office: no |
| Cameo, Teams, Presentaciones en vivo | | No (NOM tiene su proyector por código y la coautoría en vivo) |
| VBA, ActiveX | | No |

Medido contra PowerPoint 2024 real (parecido de imagen): p01 97.6 %, p02 99.0 %, p03 95.2 %, p04 98.0 %, p05 97.9 %,
p06 96.3 %, p07 98.8 %, p08 99.5 %, p09 98.7 %, p10 98.6 %, p11 89.5 %.

**Dónde pone PowerPoint cada renglón** (18-nov, `windows/hacer_interlineado.py` y `analisis/interlineado_ppt.py`: 12
fuentes, 5 tamaños, interlineado 80, 90, 100, 110, 120, 150, 200 % y exacto, midiendo la tinta de la línea base en
la imagen de PowerPoint): cada renglón baja exactamente el interlineado (1.2 × tamaño × %, o los puntos exactos
redondeados a punto entero) en todas las fuentes, y la primera línea base queda en
`max(1.2·tamaño·r + min(0, L − 1.2·tamaño), 0.75·L)` con r = (ascendente + espacio)/(ascendente + descendente +
espacio) de la tabla hhea. NOM mide dónde puso el navegador la línea base y mueve cada párrafo lo que falta
(`ajustarRenglones` en `diapositivas.js`: editor, presentar, proyector, enlace para ver). En p12: de hasta 10 pt de
diferencia a ±0.5 pt (un píxel de la medida); en p01 la diferencia mediana de los renglones baja de 2 a 1 px (37 de 55
exactos, antes 22) y en p07 de 4 a 1.5 px. Desde el 19-nov también en las miniaturas, al terminar de escribir en un cuadro y
en el PNG de la diapositiva, que lleva adentro las fuentes de la página (antes salía con las del equipo): PNG de NOM
contra la imagen de PowerPoint (`analisis/png_diapositivas.py`) p01 98.4 %, p07 98.9 %.

## 6. Lo que NOM tiene y Office no

Edición entre varios **sin Internet** (una computadora o un teléfono comparte, también por su punto de acceso);
**ramas** para lo editado sin conexión (el original no se rompe); **espacios cifrados de extremo a extremo**;
**recibo de proceso de escritura** verificable (tecleado contra pegado) y mapa de aportes; partes asignadas a una
persona; entrega con fecha límite; filtro y orden de la hoja solo para quien los pone; investigar y citar en APA sin
inteligencia artificial; datos a reporte y combinar correspondencia propios; diagramas UML y entidad-relación;
formato `.nom` comprimido; teléfono como servidor; 50 personas por espacio.

## 7. Qué sigue

El orden de construcción está en `RELEVO-ESTADO-Y-PENDIENTES.md`, sección 8 (y lo que falta por programa, en la 5).
Para el inventario por comando de la cinta: tomar los archivos de controles de Microsoft (`WordControls.txt`,
`ExcelControls.txt`, `PowerPointControls.txt` del repositorio OfficeDev) y poner una fila por comando con «Sí / A
medias / No» y su prueba, como pedía la T3 de `RELEVO-REVISION-PROFUNDA.md`.
