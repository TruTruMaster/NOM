# Relevo general: qué está confirmado, qué no y qué falta construir (17 de noviembre de 2026, noche)

Este es el relevo para la siguiente sesión. **Leerlo primero**, junto con `PLAN-CONSTRUCCION.md` (fases y marcas),
`LOG.md` (entradas 71 a 80), `RELEVO-17NOV-ANIMACIONES.md` (detalle de PowerPoint), `RELEVO-17NOV-LISTA-OFFICE.md`
(detalle de Word, Excel, contraseñas, tablas), `PARIDAD-OFFICE-REAL.md` (cómo se mide contra Office real) y
`PLAN-MEDIOS.md` (video y audio). Aquí no se repite el detalle técnico: se dice **qué está probado de verdad, qué solo
se escribió y qué falta**, para Word, Excel y PowerPoint, y cómo tiene que convivir cada función con lo que NOM ya
tenía. La sección 9 es la revisión completa de la noche del 17-nov (código, documentación, la Pi y GitHub).

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

## 1. Estado del repositorio (verificado la noche del 17-nov)

- `main` va **26 commits adelante de `origin/main`** (`b65438d`, la 0.15.4) más los de la revisión de esta noche
  (sección 9). **Nada de eso está publicado**: ni en GitHub (empujar a `main` despliega a Azure si toca
  `4-CODIGO/app/**` o `4-CODIGO/nucleo/**`), ni en la Pi, ni en el APK. Comprobado con `git fetch`.
- **Lo publicado:** la Pi sirve la **0.15.4** (interfaz 20, jar del 16-nov 16:57) en `https://www.nom.qd.je` y
  `https://nomcolab.dpdns.org`; la release «latest» de GitHub es **v0.15.4** (`NOM-COLAB.apk`, código 30); Azure,
  apagada (403). El repositorio es **público**, con licencia GPL-3.0 y «Website» `https://www.nom.qd.je`.
- El código dice **0.15.4** (`Version.java`, `NOM.VERSION` en `nucleo.js`, `VERSION` en `sw.js`,
  `versionName`/`versionCode` en `android/app/build.gradle.kts`); la ayuda y `CAMBIOS.md` ya describen la **0.16.0**.
  Al publicar: subir los cuatro a 0.16.0 (código de APK 31) e **INTERFAZ a 21** (la 0.16.0 trae campos nuevos en los
  documentos: `transOffice:`, árboles de animación, `oc`/`pos` del texto de Word…), regresión completa antes.
- Sin rastrear a propósito: `7-PRESENTACION/`, `ENTREGA/5IM7_WACK_Casos_de_Uso.*`, `ENTREGA/5IM7_WACK_Presentacion.pdf`.
- `tmp-17-11-26/` (3 GB, en `.git/info/exclude`): lo útil ya está en `~/NOM/trabajo/`; **se puede borrar con permiso
  del usuario** (no se ha borrado).
- Servidor de pruebas local: `cd 4-CODIGO/app && SPRING_WEB_RESOURCES_STATIC_LOCATIONS=file:/home/user/NOM/4-CODIGO/app/src/main/resources/static/ NOM_DATOS=/home/user/nom-datos-pruebas NOM_PUERTO=8099 nohup java -jar target/nom-colab.jar &`
  (el JS y el CSS se sirven en vivo; **`mvn package` antes** si cambió Java). Para medir contra Office sin estorbar a
  la regresión sirve un segundo servidor (copia del jar, otro `NOM_DATOS`, `NOM_PUERTO=8097`) con
  `NOM_BASE=http://localhost:8097` en `comparar_real.py`.

## 2. Confirmado (medido de verdad)

**La meta es 100 % en todo** (pedido del usuario, 17-nov): un renglón, una celda o un cuadro de animación «igual» o no;
lo que esté por debajo de 100 % es trabajo pendiente, no «suficiente». El parecido de imagen (PowerPoint) nunca da 100 %
exacto por el suavizado de bordes: para PowerPoint, la meta es diferencia de píxeles por debajo del umbral de suavizado
en cada diapositiva y cada cuadro, y cambiar `comparar_real.py` y `animaciones_real.py` para que lo reporten así.

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

Revisado en el código el 17-nov (búsqueda por nombres en español e inglés): lo que dice «no existe» no aparece en
`js/`; lo que existe a medias dice qué hay. Al empezar cada punto, volver a buscar por si otra sesión lo hizo.

Cada punto lleva entre corchetes cómo se comprueba: **[R]** contra Office real (Windows), **[V]** la vuelta (exportar,
abrir en Office sin reparar, validador), **[T]** en el teléfono/app.

### Word
1. **División automática de palabras** (guiones, `w:autoHyphenation`) [R] — es lo que más baja w10.
2. **Tabla flotante** (`w:tblpPr`) con el texto alrededor, en el motor de «lo que flota» [R][V].
3. **Citas y bibliografía de Word:** NOM ya tiene sus propias citas (`citas.js`: fuentes, APA, lista de referencias).
   Falta el puente con Word: leer y escribir los campos `CITATION`/`BIBLIOGRAPHY` y `customXml/Sources.xml` [R][V].
4. **Combinar correspondencia de Word:** NOM ya lo hace a su modo (`datos.js`: texto con `{{campos}}` + una hoja = un
   documento por fila). Falta el de Word: `MERGEFIELD`, el origen de datos guardado en el .docx y la vuelta [R][V].
5. Campos: PAGE/NUMPAGES ya se leen y se escriben, y la tabla de contenido se exporta. Falta actualizar al editar DATE,
   REF, SEQ (títulos de figura), referencias cruzadas e índice, y los puntos guía del índice [R].
6. **Pegar con formato:** ya existe a medias (`modo-texto.js` toma el HTML del portapapeles con negritas, listas,
   vínculos…). Falta completarlo desde Word, Google Docs y páginas web: estilos, tablas, imágenes, colores (decidido 17-nov) [T].
7. Controles de contenido: los de Word (`w:sdt`) ya se conservan y su texto se edita. Falta usarlos como formularios
   (casilla, lista desplegable, fecha) y crearlos desde NOM; marcas de agua editables; temas de documento.
8. Rellenos degradado/trama/imagen en formas de Word, puntas y guiones de líneas, lienzos (parte hecha en `fa3cc00`:
   volver a medir w06).
9. Imágenes **EMF/WMF**: se conservan dentro del original (salen al exportar), pero **no se ven** en NOM; dibujarlas
   (también en Excel y PowerPoint) [R].
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

0. **Llevar a 100 % lo que ya se mide** (sección 2): cada renglón de Word que no es idéntico, cada celda o alto de
   Excel distinto, cada diapositiva y cada cuadro de animación de PowerPoint con diferencia. Es la base: lo nuevo se
   construye sobre lo que ya da exacto.
1. Medios fase 2 y 3 (almacén + insertar), porque el usuario ya decidió y es lo que más se nota.
2. PowerPoint: los 2 px del texto, calibrar transiciones y Morph (necesitan la Windows: hacerlo estando el usuario).
3. Word: división de palabras, tabla flotante, pegar con formato, citas y bibliografía, combinar correspondencia.
4. Excel: herramientas de datos (minigráficos, segmentaciones, quitar duplicados, texto en columnas, buscar objetivo,
   Solver, escenarios), crear tablas dinámicas.
5. Comunes: EMF/WMF, efectos 3D, `.od*` y formatos viejos, VBA, Power Query.
6. Inventario por pestaña (`PARIDAD-OFFICE.md`) para no olvidar nada.

Herramientas y trucos (Windows por SSH, `office_real.py`, `comparar_real.py`, análisis de animaciones): sección 3 de
`RELEVO-17NOV-ANIMACIONES.md` y sección 4 de `RELEVO-17NOV-LISTA-OFFICE.md`.

## 9. Revisión completa (17 de noviembre de 2026, noche)

Pedido del usuario: auditar todo el código, corregir **toda** la información equivocada del repositorio (también lo
hecho que seguía marcado como pendiente), revisar la página, la Pi y GitHub, y después seguir construyendo (paso 0 de
la sección 8), «como un ingeniero muy cuidadoso que busca incongruencias entre la documentación y el código».

### 9.1 Lo que se revisó
- **Servidor** (Java, `app/` y `nucleo/`): filtros de sesión, CORS, cabeceras, origen de las peticiones, todas las
  rutas de `*Api.java`, servicios de cuentas, espacios, documentos (permisos, cuotas, proyección, ramas, sincronía,
  enlace para ver), canal en vivo y repositorios. Paridad de rutas con el teléfono (`ApiTelefono`): todas salvo
  `POST /api/cuentas/aceptar`, que allí no aplica (no hay cuentas con contraseña).
- **Interfaz**: `node --check` y ESLint (variables sin definir, claves repetidas, código inalcanzable…) de los 54
  archivos: limpio salvo lo de abajo; reglas del proyecto (`data-nom` en los estilos inyectados, variables CSS que
  existen); módulos sin probar del 17-nov (medios, transiciones, proyector).
- **App de Android**: WebView (solo `file:///android_asset/`, lo demás al navegador del sistema), puente `NomAndroid`
  (solo en la propia interfaz; la política de contenido impide marcos ajenos).
- **Afuera**: la Pi (`nomcolab-pi estado`, `/etc/nomcolab.conf`, túnel), las tres direcciones, GitHub (repositorio,
  releases, ramas, Pages de `nom-qd-je`), Azure.

### 9.2 Defectos encontrados y arreglados
| Dónde | Qué pasaba | Arreglo |
|---|---|---|
| `ServicioDocumentos.actividadExterna` | Un editor podía escribir filas de la bitácora **a nombre de otra cuenta** (bastaba su id y `nube`): falsificaba su recibo de proceso y sus aportes; además filas con hora en el futuro | Solo a nombre de quien sube o si sube el propietario del espacio (el que sincroniza su teléfono); horas limitadas al presente. `seguridad.py` 17.1 |
| Proyección (`/api/proyeccion/{código}`, sin sesión) | Daba **todo** el estado: comentarios, notas del orador, dueños y el Office original; el código seguía sirviendo hasta 8 h aunque a quien presenta le quitaran el acceso o se moviera el documento; los cambios en vivo también llevaban comentarios y notas | Foto y cambios sin comentarios, notas, dueños ni `_origen`; el código deja de servir sin acceso; al quitar el acceso o mover el documento se cierran sus pantallas. 17.2 |
| Enlace para ver (`/api/ver/…`) | Quitaba comentarios y notas, pero dejaba el `.docx`/`.pptx` original dentro (con los suyos) | Sin `_origen`. 17.3 |
| Alias (`/api/enlaces/{c}/alias` y `/invitado`) | Con un código vigente se creaban cuentas sin fin | Tope por red y en total por hora (`NOM_ALIAS_HORA_RED`, 300; `NOM_ALIAS_HORA`, 3000). 17.4 |
| `Origen` | Se creían las cabeceras de reenvío (CF-Connecting-IP…) desde direcciones `fe80::`/`169.254.x` también fuera de Azure: cualquier equipo de la misma red falseaba su IP y se saltaba los topes | Solo en Azure (`WEBSITE_SITE_NAME`) |
| `ManejoErrores` | Un `IllegalArgumentException` sin mensaje daba 500; los de Java mostraban nombres de clases («No enum constant mx.wack…») | 400 con el mensaje propio o «La petición no es válida». 17.5 |
| Memoria del servidor | Todo documento abierto se quedaba en memoria hasta reiniciar (con su Office original, hasta 14 MB): la Pi (512 MB para Java) se llenaba con unos cuantos | Salen de la memoria tras 20 min sin uso, guardados y sin canales ni versión pendiente, en el hilo de `persistir()` (`-Dnom.inactivoMs`). 17.6 |
| Mover un documento | Quien lo tenía abierto sin ser del espacio destino seguía recibiendo los cambios en vivo | Se le cierra el canal |
| `cambiosDesdeBase`, `actividadDesde` | Cortaban en la primera racha vieja (lo que ya se había corregido en `cambiosDesde`) | Sin cortar |
| `ecuaciones.js` | Una ecuación que terminaba en `\` (o `\` antes de un salto) tiraba el dibujo del documento y la exportación | Se escribe tal cual |
| Panel de transiciones vs. proyector y miniaturas | El panel guarda `transOffice:` y borra `trans:`; la pantalla del proyector y la marca «Con transición» de las miniaturas solo leían `trans:` | El proyector juega la transición de PowerPoint; la marca lee las dos |
| `union.js` | «Dejar entrar sin cuenta: cada quien **escribe un apodo**» — desde la fase 13 es un alias al azar | Texto corregido |
| Importar Office de más de 14 MB | No se guarda el original (`MAX_ORIGEN`) y no avisaba; el aviso de PowerPoint decía «se conservan» | Aviso al importar |
| Word, texto oculto (`w:vanish`) y elevado/bajado (`w:position`) | No se leían: NOM **mostraba lo oculto** («ESTO NO SE VE» de w01) | Se leen, se dibujan y se escriben; al parchar uno importado antes no se quitan |
| `--peligro` (editor.js, formatos.js) | La variable no existe en `nom.css`: en tema oscuro el rojo de respaldo no contrasta | `--error` |
| `ServicioFuentes.AGENTE` | Decía «proyecto escolar» (contra lo decidido el 17-nov) | «NOM/1.0 (sin fines de lucro)» |
| `azure/nomcolab local up` | Encendía el túnel de la laptop, que es **el mismo** de la Pi (869abbea…): Cloudflare habría repartido a la gente entre dos servidores con datos distintos | No lo enciende si la nube contesta desde otro equipo |
| `prueba_funciones.py` | Buscaba el menú viejo «Desvanecer»: la sección del lienzo se cortaba ahí (no corrían presentador ni .pptx) | Usa el panel nuevo y revisa `transOffice:` |

`rastreo.py`: 6 «MAL», revisados a mano: son de la prueba (toca «Bloquear» y luego «Cortar», que queda apagado;
«Pegar especial» depende de lo copiado; «Acción al hacer clic» abre una hoja que tapa la ruedita).

### 9.3 Documentación corregida
LEEME y README (direcciones, nombres de botones que ya no existen, «con apodo», cifras de pruebas, licencias de
Mediabunny y de las fuentes «NOM …», escuela), `PLAN-CONSTRUCCION.md` (casillas de lo ya hecho en las fases 10 y 22,
erratas `~/NOM` por `~/NOM-COLAB`, cifras viejas), `SERVIDOR-PROPIO.md` (`doas nomcolab-pi` no sirve sin la ruta
completa; dirección principal; aviso del túnel compartido), `DOMINIO-Y-BUSCADORES.md` (`nom.qd.je` todavía sin
https), `DISENO-UI.md`, `MARCA.md`, `PLAN-MEDIOS.md`, los relevos viejos marcados como históricos, la ayuda («Qué se ve
igual», abrir y descargar, preguntas 16 y 24, 230 funciones de hoja y no «casi 100») y las novedades de la 0.16.0 (le
faltaban tablas y secciones de Word, vínculos, sonidos, gráficas y SmartArt por partes).

### 9.4 Pendiente de decidir (el usuario)
- **`nom.qd.je` sin certificado**: según la documentación de GitHub Pages hay que quitar y volver a poner el dominio en
  el repositorio `nom-qd-je` para que lo emita (cambio en GitHub; no se hizo sin preguntar).
- En GitHub sigue una rama vieja del 28-oct (no es `main` ni `desarrollo`) con commits que rompen la regla de menciones
  del proyecto: borrarla solo si el usuario lo pide.
- Publicar la 0.16.0 (sección 7: falta la regresión completa, la vuelta en PowerPoint real y el A55).

