# Relevo general: qué está confirmado, qué no y qué falta construir (17 de noviembre de 2026, tarde)

Este es el relevo para la siguiente sesión. **Leerlo primero**, junto con `PLAN-CONSTRUCCION.md` (fases y marcas),
`LOG.md` (entradas 71 a 79), `RELEVO-17NOV-ANIMACIONES.md` (detalle de PowerPoint), `RELEVO-17NOV-LISTA-OFFICE.md`
(detalle de Word, Excel, contraseñas, tablas), `PARIDAD-OFFICE-REAL.md` (cómo se mide contra Office real) y
`PLAN-MEDIOS.md` (video y audio). Aquí no se repite el detalle técnico: se dice **qué está probado de verdad, qué solo
se escribió y qué falta**, para Word, Excel y PowerPoint, y cómo tiene que convivir cada función con lo que NOM ya
tenía.

## 0. Lo que pidió el usuario (vigente)

- **Todas las funciones de Word, Excel y PowerPoint**, las complejas y las de pago que LibreOffice, Google Docs u
  OpenOffice no tienen, «fully and proven»: comprobadas contra Office real en la Windows (`user@100.65.29.5`).
- **Cada función emparejada con lo que NOM ya tenía** (ver la sección 5): edición entre varios letra por letra,
  sin Internet, espacios cifrados, ramas, teléfono como servidor, app de Android, ruedita y paleta, ayuda, `.nom`,
  y la vuelta a Office (lo que sale de NOM se abre en Office sin «reparar» y se ve igual).
- Fuentes: **«we stay on the gray zone, with all the features no matter what»** (interletraje y ligaduras de las
  fuentes de Microsoft se quedan; no volver a preguntar).
- Video y audio: en la nube (la Pi tiene 2 TB), convertidos en el equipo a AV1 + Opus, 5 GB por cuenta, lo más ligero
  posible (memoria `nom-colab-medios-decisiones`, plan `PLAN-MEDIOS.md`).
- Mientras el usuario está fuera de casa: **nada de pruebas largas, de regresión ni en el teléfono**; se hacen al
  volver. Por eso casi todo lo del 17-nov está «hecho, sin confirmar» (sección 3).
- **Una sola sesión a la vez en `~/NOM`.** El 17-nov hubo dos sesiones editando la misma copia; se coordinaron con
  `ListAgents`/`SendMessage`, pero es fácil pisarse. Antes de empezar, `ListAgents`: si hay otra, acordar qué archivos
  toca cada una.

## 1. Estado del repositorio

- `main` va **24 commits adelante de `origin/main`** (de `6fadd42` a `71243e9`, más este relevo). **Nada de eso está
  publicado**: ni en GitHub (empujar a `main` despliega a Azure si toca `4-CODIGO/app/**`), ni en la Pi, ni en el APK.
- El código dice **0.15.4** (`Version.java`, `NOM.VERSION` en `nucleo.js`, `VERSION` en `sw.js`,
  `versionName`/`versionCode` en `android/app/build.gradle.kts`); la ayuda y `CAMBIOS.md` ya describen la **0.16.0**.
  Al publicar: subir los cuatro a 0.16.0 (código de APK 31), regresión completa antes.
- Sin rastrear a propósito: `7-PRESENTACION/`, `ENTREGA/5IM7_WACK_Casos_de_Uso.*`, `ENTREGA/5IM7_WACK_Presentacion.pdf`.
- `tmp-17-11-26/` (3 GB, en `.git/info/exclude`): lo útil ya está en `~/NOM/trabajo/`; **se puede borrar con permiso
  del usuario** (no se ha borrado).
- Servidor de pruebas local: `cd 4-CODIGO/app && SPRING_WEB_RESOURCES_STATIC_LOCATIONS=file:/home/user/NOM/4-CODIGO/app/src/main/resources/static/ NOM_DATOS=/home/user/nom-datos-pruebas NOM_PUERTO=8099 nohup java -jar target/nom-colab.jar &`
  (el JS y el CSS se sirven en vivo). Los 24 commits cambian también 6 archivos de Java (`CuentasApi`, `PaginasApi`…)
  y el jar es de las 03:15 del 17-nov: **`mvn package` antes de probar** lo del servidor.

## 2. Confirmado (medido de verdad)

«Confirmado» = medido contra Office real o probado de punta a punta. **Ojo:** casi todo se midió con la 0.15.x;
desde entonces hay 24 commits sin regresión, así que hay que **volver a medirlo** antes de publicar (sección 6).

### Word (contra Word 2024 real, `comparar_real.py`, renglones idénticos y páginas)
| Archivo | Renglones idénticos | Páginas Word / NOM | Medido |
|---|---|---|---|
| w01 letra | 85.4 % | 6 / 6 | 0.15.0 |
| w02 párrafo | 82.4 % | 9 / 9 | 0.15.0 |
| w03 estilos y listas | 96.6 % | 7 / 7 | 0.15.0 |
| w04 secciones | 38.4 % (después mejoró con secciones de otra hoja y columnas, `fa3cc00`, sin volver a medir la cifra) | 9 / 8 | 0.15.0 |
| w05 tablas | 85.7 % | 8 / 8 (tablas como Word, `03c64a3`) | 17-nov |
| w06 objetos | 77.0 % | 18 / 18 | 0.15.3 |
| w07 referencias | 91.2 % | 6 / 6 | 0.15.0 |
| w08 revisión (control de cambios, comentarios) | 94.3 % | 2 / 2 | 0.15.0 |
| w09 informe | 77.3 % | 9 / 9 | 17-nov |
| w10 raros | 18.1 % | 6 / 6 | 0.15.0 |
- Once de doce Word con las mismas páginas que Word (17-nov). La tabla editada se descarga solo con las celdas que
  cambiaron; Word la abre sin avisos y el validador de Microsoft no da errores nuevos.
- Contraseñas de Office (MS-OFFCRYPTO, Agile y Standard) y protección (restringir edición, marcar como final,
  contraseña para modificar): probadas en Office real (`a9e388a`, `1b0d8ea`).

### Excel (contra Excel 2024 real: texto de cada celda, rellenos, anchos y altos)
| Archivo | Texto | Rellenos | Anchos | Altos |
|---|---|---|---|---|
| x01 formato | 100 % | 100 % | 100 % | 98.2 % |
| x02 fórmulas | 98.3 % → 100 % en las matrices dinámicas (x02/x09) | 100 % | 100 % | 100 % |
| x03 formato condicional | 100 % | 100 % | 100 % | 100 % |
| x04 gráficas (37, contra `Chart.Export`) | 100 % | — | 100 % | 100 %; cada gráfica 92–98 % de parecido |
| x05 tablas y tablas dinámicas | 100 % | 100 % | 100 % | 100 % |
| x06 objetos | 100 % | — | 100 % | 100 % |
| x07 hojas | 100 % | — | 100 % | 100 % |
| x08 registro y presupuesto | 100 % | 100 % | 100 % | 100 % |
- Motor de fórmulas: matrices dinámicas, `LET`, `LAMBDA` (16-nov, x02/x09 100 %). Líneas de tendencia con ecuación y
  R², barras de error y tabla de datos de las gráficas (`6fadd42`).

### PowerPoint (contra PowerPoint 2024 real)
| Qué | Resultado |
|---|---|
| p01–p10 (parecido de imagen, 0.15.0) | p01 97.6 %, p02 99.0 %, p03 95.2 %, p04 98.0 %, p05 92.1 %, p06 94.0 %, p07 98.8 %, p08 99.5 %, p09 96.8 %, p10 97.9 % |
| Motor de animaciones, a01 (195 efectos) contra el video de PowerPoint | 97.0 % medio, 159 ≥ 98 % |
| a02 opciones (por letra, palabra, párrafo, repetir, suaves, rebote…) | letras y palabras 94–97 %; «Destello en negrita por letra» 84 %, «Volar por párrafo al revés» 85 % |
| Inicio/fin suaves (16 combinaciones) | error medio 0.12 % del recorrido |
| Rebote al final | error medio 0.17 % |
| La vuelta: los 195 efectos reescritos por NOM abiertos en PowerPoint real (`v01-nom.pptx`) | abre sin reparar; su video contra NOM 98.3 % medio |
| Validador de Microsoft sobre lo que exporta NOM (a01, a02, a03, p08) | 0 errores |
| Transición «Desvanecer» | 97 % (las demás sin medir) |

## 3. Hecho, pero **sin confirmar** (solo prueba de humo o ninguna)

Todo esto está en el código y en commits locales; **no** se probó en el teléfono, ni con la regresión, ni en Office
real salvo lo que se dice.

| Qué | Commit | Qué se probó | Qué falta probar |
|---|---|---|---|
| Panel de animaciones como PowerPoint: galería de 262 efectos, opciones, intervalos, panel con línea de tiempo, copiar, trayectorias dibujadas | `f3d3ea8`, `5ee8d97` | humo en el navegador con a02 (importar, galería, opciones, vista previa, panel, trayectoria, `timingXml`) | abrir en PowerPoint real un archivo armado en el panel (repetir, rebote, atenuar, ocultar, por párrafos, trayectoria propia) + validador; teléfono (barras con el dedo, trayectoria con el dedo, ruedita) |
| Panel de transiciones (151 variantes) | `250c800` | humo con a03 | las duraciones iniciales son «de memoria»; **calibrar** contra el video (`transiciones_real.py`); Morph sin medir |
| Vínculos y botones de acción | `619efac` | ida y vuelta (importar lo exportado) con los tres tipos | PowerPoint real; teléfono |
| Video y audio de PowerPoint al presentar (desde el original) | `f6f6835` | humo con 3 archivos (reproducir solo, clic, audio, Alt+Q, limpieza) | PowerPoint real (que `p14:trim end` es lo que se quita al final y `p14:fade` en ms), WebView del A55 |
| Gráficas por serie/categoría y SmartArt por nodo | `abfebb0` | humo (lo-tdf104792) e ida y vuelta | contra el video de PowerPoint |
| Sonidos de efectos y transiciones al presentar | `71243e9` | humo (ms-rendimiento-tipico, lo-tdf148478) | PowerPoint real; teléfono |
| Conversión de video/audio a AV1+Opus en el equipo (fase 1 de medios) | `9361285` | medido en Chromium de Playwright con contenido sintético | A55 y archivos reales; **todavía no se usa** (falta la fase 2) |
| Arreglo del motor: comandos de medios después del primero | `f6f6835` | humo | — |
| p:timing dentro de `mc:AlternateContent` (rebote) al importar, exportar y podar | `f3d3ea8` | humo | validador + PowerPoint real |
| «NOM deja de presentarse como proyecto escolar» | `613f4d6` | — | revisar con el usuario el texto visible |
| Ayuda, novedades y `CAMBIOS.md` de la 0.16.0 | varios | generadas con `generar.py` | leerlas completas antes de publicar |

## 4. Lo que se sabe que no está bien

- **Texto estático de PowerPoint ~2 px más arriba** que en PowerPoint (Arial 54 negrita: tinta en 284 contra 288 px a
  1280 de ancho). Sube también todos los efectos de texto (por eso los de texto quedan en 88–95 %).
- «Destello en negrita por letra» (84 %) y «Volar por párrafo al revés» (85 %) contra el video.
- Word w10 «raros» 18 % (falta la división automática de palabras) y w04 secciones por debajo del resto.
- Excel: un alto de fila con texto ajustado que Excel recalcula al abrir (aquí se respeta lo guardado; pasó una vez).
- Al exportar a .pptx **sin** el archivo original, el SmartArt sale como figuras sueltas y los videos, audios y sonidos
  se pierden (salen la portada o nada).
- `nom.qd.je` sin https (GitHub Pages no había emitido el certificado); su página de redirección aún dice NOM-COLAB.
- `.doc`, `.xls`, `.ppt`, `.odt`, `.ods`, `.odp` y `.rtf` se pueden elegir al importar, pero solo dan un aviso: **no
  se abren**.

## 5. Falta construir (paridad con Office), por programa

Cada punto lleva entre corchetes cómo se comprueba: **[R]** contra Office real (Windows), **[V]** la vuelta (exportar,
abrir en Office sin reparar, validador), **[T]** en el teléfono/app.

### Word
1. **División automática de palabras** (guiones, `w:autoHyphenation`) [R] — es lo que más baja w10.
2. **Tabla flotante** (`w:tblpPr`) con el texto alrededor, en el motor de «lo que flota» [R][V].
3. **Citas y bibliografía** (campos `CITATION`/`BIBLIOGRAPHY`, `customXml/Sources.xml`) unidas a `citas.js` [R][V].
4. **Combinar correspondencia** (`MERGEFIELD`, origen de datos de Excel/CSV, vista previa, generar documentos) [R][V].
5. Campos en general (`w:fldSimple`/`instrText`: hoy se leen; actualizar PAGE, NUMPAGES, DATE, REF, SEQ, TOC al
   editar) y puntos guía del índice [R].
6. **Pegar con formato** desde Word, Google Docs o una página web (decidido el 17-nov; `modo-texto.js` ya lee
   `text/html` en parte: completarlo con estilos, listas, tablas e imágenes) [T].
7. Controles de contenido y formularios (casillas, listas, fechas), marcas de agua editables, temas de documento.
8. Rellenos degradado/trama/imagen en formas de Word, puntas y guiones de líneas, lienzos (parte hecha en `fa3cc00`:
   volver a medir w06).
9. Imágenes **EMF/WMF** (hoy no se ven; también en Excel y PowerPoint) [R].
10. Abrir **.doc** (Word 97-2003): estimado 2-3 sesiones (texto, formato, tablas, imágenes, encabezados y pies).
11. Macros: conservar `vbaProject.bin` (.docm), ver y editar módulos (MS-OVBA), intérprete aislado.

### Excel
1. **Herramientas de datos**: minigráficos (sparklines: no existen), segmentaciones (slicers: no existen), quitar
   duplicados, texto en columnas, relleno rápido, consolidar, previsión, buscar objetivo, **Solver**, escenarios,
   tablas de datos (análisis Y si…) [R][V].
2. Tablas dinámicas: hoy se ven (x05 100 %); falta **crearlas y editarlas** (campos, filtros, actualizar) [R][V].
3. Comentarios con hilo (`threadedComment`) y notas editables [V].
4. Configuración de página e impresión de Excel (área de impresión, saltos, encabezados) [R] — la Windows tiene la
   cola de impresión apagada: activarla para medir.
5. **Power Query / Power Pivot**: conservar y mostrar consultas y el modelo; VBA (.xlsm) igual que en Word.
6. Gráficas 3D distintas de columnas y circular; formato detallado de gráficas en el editor de NOM.
7. Abrir **.xls**.

### PowerPoint
1. **Calibrar las transiciones** contra el video (`transiciones_real.py`; empezar por Empujar, Barrido, Dividir,
   Revelar, Barras, Forma, Cubrir, Descubrir, Destello, Cubo, Galería, Caída, Cortinas, Curvatura) y **Morph** por
   objeto, palabra y letra (p08 diap. 36, a03 152–154) [R].
2. Arreglar los ~2 px del texto estático (sección 4) [R].
3. **Medios, fases 2 a 5 de `PLAN-MEDIOS.md`**: almacén en el servidor (subir en partes y reanudar, `Range`, permisos
   por espacio, cuota de 5 GB, cifrado en reposo y extremo a extremo, borrar lo huérfano, teléfono como servidor y sin
   Internet), **insertar video y audio** desde NOM (presentaciones y documentos), al descargar .pptx convertir a MP4,
   al abrir un .pptx ofrecer guardar sus videos convertidos [R][V][T]. Video en la pantalla del proyector, medios en
   línea (YouTube), desencadenadores por marcador (`onMediaBookmark`; falta un archivo de PowerPoint que lo tenga).
4. Datos que faltan en la biblioteca de efectos (necesitan PowerPoint): direcciones de las salidas y algunas de
   entrada (Desplazar 3 y 12, Barras diagonales 3, Dividir vertical hacia dentro, Rueda por radios, Arrastrar 4,
   Rotar vertical, Zoom desde el centro de la diapositiva) [R].
5. Elegir un sonido para un efecto o transición desde NOM; sonido de una acción (`a:snd`); acción «al pasar el mouse»
   (`hlinkMouseOver`); vínculos dentro del texto (`a:rPr/a:hlinkClick`); abrir otro archivo o programa.
6. Editar los puntos de una trayectoria en el editor (hoy se dibuja de nuevo); crear SmartArt «uno por uno» desde NOM.
7. Efectos 3D en formas (`scene3d`/`sp3d`: no existen), modelos 3D (al menos su imagen de respaldo).
8. Grabar la presentación (narración, entintado, tiempos) y exportar a video; ensayar intervalos.
9. Abrir **.ppt**.

### Para los tres
- **.odt / .ods / .odp** (abrir y guardar), RTF.
- Terminar el **inventario por pestaña** (`docs/paridad/`: WORD.md casi completo, POWERPOINT.md solo Inicio e
  Insertar, EXCEL.md sin escribir) y convertirlo en `docs/PARIDAD-OFFICE.md`: una fila por comando de la cinta con
  «NOM: sí / a medias / no» y su prueba (T2/T3 de `RELEVO-REVISION-PROFUNDA.md`).

## 6. Emparejar cada función con lo que NOM ya tenía

Una función de Office no está terminada hasta que pasa por esta lista (es lo que el usuario llama «paired with all
the features we already had»):

| Lo de NOM | Qué hay que comprobar con cada función nueva |
|---|---|
| Edición entre varios (CRDT, letra por letra) | dos personas cambiando lo mismo a la vez (p. ej. una animación, una tabla dinámica, un campo) sin perder nada; el cambio se ve en el otro equipo |
| Sin Internet y al volver | se usa sin conexión y se sincroniza al volver (los medios esperan en el equipo hasta subirse) |
| Espacios cifrados (extremo a extremo) | el servidor no ve el contenido nuevo (medios incluidos: fase 2) |
| Ramas | la función sobrevive a crear rama y unirla |
| Teléfono como servidor | funciona con el teléfono haciendo de servidor |
| App de Android (WebView) | se ve y se usa con el dedo; **requiere recompilar el APK** (`assembleRelease` + `publicar-apk.sh`) |
| Ruedita y paleta de comandos | cada acción nueva está en la ruedita y en la paleta (reglas del plan) |
| Ayuda | explicada en `4-CODIGO/ayuda/contenido` y `generar.py` corrido |
| `.nom` | se guarda y se vuelve a abrir igual |
| La vuelta a Office | se exporta, Office la abre sin reparar, el validador no da errores nuevos, se ve igual |
| Apariencia | reglas de `DISENO-UI.md` §1 y §14 (variables de `nom.css`, un color de marca, sin degradados ni desenfoque); las hojas de estilo inyectadas llevan `data-nom` (si no, `nucleo.js` las quita) |
| Accesibilidad | teclado, lector de pantalla, alto contraste, «reducir movimiento» |

## 7. Pruebas pendientes (cuando el usuario esté en casa), en este orden

1. `pruebas/preparar.sh` y `pruebas/regresion.sh rapido`; luego la regresión completa.
2. `comparar_real.py` de los 28 archivos (w01–w10, x01–x08, p01–p10) para confirmar que nada bajó.
3. La vuelta en PowerPoint real de un archivo armado con el panel de animaciones, transiciones, vínculos, medios y
   sonidos + `validar_openxml.py`.
4. `transiciones_real.py` (calibración) y Morph.
5. A55: volver a vincular la depuración inalámbrica; `app_emulador.py`, `telefono_emulador.py`, `rama_emulador.py`;
   paneles con el dedo; reproducir medios en el WebView; convertir un video en el teléfono.
6. Publicar 0.16.0 (sección 1) solo si todo lo anterior pasa.

## 8. Orden recomendado para construir

1. Medios fase 2 y 3 (almacén + insertar), porque el usuario ya decidió y es lo que más se nota.
2. PowerPoint: los 2 px del texto, calibrar transiciones y Morph (necesitan la Windows: hacerlo estando el usuario).
3. Word: división de palabras, tabla flotante, pegar con formato, citas y bibliografía, combinar correspondencia.
4. Excel: herramientas de datos (minigráficos, segmentaciones, quitar duplicados, texto en columnas, buscar objetivo,
   Solver, escenarios), crear tablas dinámicas.
5. Comunes: EMF/WMF, efectos 3D, `.od*` y formatos viejos, VBA, Power Query.
6. Inventario por pestaña (`PARIDAD-OFFICE.md`) para no olvidar nada.

Herramientas y trucos (Windows por SSH, `office_real.py`, `comparar_real.py`, análisis de animaciones): sección 3 de
`RELEVO-17NOV-ANIMACIONES.md` y sección 4 de `RELEVO-17NOV-LISTA-OFFICE.md`.
