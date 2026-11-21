# Relevo del 21 de noviembre de 2026

**Leer este primero.** Sigue a `RELEVO-20NOV-NOCHE.md` (la sesión empezó el 20-nov a las 22:55). Al empezar estaba
publicada la **0.18.1**; esta sesión preparó la **0.19.0** (interfaz 24, APK 36, commit `ee19278`), que trae también lo
del 20-nov en la noche (guiones y números de línea como Word, contraseñas con versión). **Publicada la 0.19.0 el 21-nov a las 03:30** (§3).

Pedido vigente del dueño (20-nov, noche): seguir con todo lo de `~/NOM/prompt.md` hasta el jueves 26-nov.

## 1. Hecho en esta sesión

| Qué | Comprobado |
|---|---|
| **Menús que caben en el celular** (`50fdb50`): `prueba_funciones` daba 201/204 porque, con los 7 renglones de los números de línea, el «Más» del texto se salía de la pantalla del celular y no se llegaba a lo de abajo («Modo sugerencia», «Configurar página…»): el menú no tenía alto máximo. Ahora `.menu` mide a lo más la pantalla y se desplaza (`nom.css`), se acomoda arriba o pegado al borde si no cabe (`nucleo.js` `menu`), y los números de línea van en su propio menú («Más» → «Números de línea»: Ninguno, Continuos, Cada página, Cada sección, Suprimir en este párrafo, Opciones…), como el de Word | `prueba_funciones` 204/204 (dos veces); `prueba_acomodo` con 5 casos nuevos (el menú cabe, abre el suyo, continuos 1, 2, 3…, queda marcado, «Ninguno» los quita) |
| **La app de Android con su propio origen** (`5255bc0`): la interfaz ya no se abre con `file://`; `MainActivity` contesta `https://appassets.androidplatform.net/…` con sus archivos de `assets/` (cargador propio, sin dependencias; carpetas con su `index.html`; 404 sin salir a la red) y la red del salón (http) se deja con `MIXED_CONTENT_ALWAYS_ALLOW`. **Lo guardado por las versiones de antes pasa solo la primera vez**: `migrar.html` + `js/migrar-origen.js` corren en `file://` (leen `localStorage`, `nom-boveda` y `nom-binarios`: bytes, fechas, archivos, mapas; mensajes de ≤ 1 M caracteres por el puente, a un archivo) y luego en el origen nuevo (lo escriben con los mismos almacenes, índices y llaves; se comprueba la cuenta). Lo de `file://` no se borra; si falla, se reintenta al abrir (3 veces). Atrás no regresa a la página de paso. Spring (`FiltroCors.ORIGEN_APP`) y el teléfono (`ServidorHttp.cors`) aceptan el origen nuevo (las apps de antes siguen entrando igual). Si el equipo que comparte es de antes de la 0.19, la app dice «tiene una versión anterior de NOM» en vez de «sin conexión» (`api.js` `equipoViejo`: un pedido `no-cors`). La ayuda dentro de la app usa sus `index.html` (`nucleo.js`, `ayuda.js`) | `pruebas/reales/paso_origen.py` (nueva: instala la de antes, entra, escribe con y sin red, guarda 1.5 MB sin red, instala la nueva sin red y compara cada registro con SHA-256): **29/29 en el emulador (WebView 124) y 29/29 en el A55 (WebView 157)** con una copia de prueba (`NOM_PAQUETE=mx.wack.nom.prueba`, sin tocar la app del dueño); `origen_equipos.py` (nueva: equipo nuevo, de antes, sin respuesta, el teléfono como servidor, la ayuda con su búsqueda) **9/9 en los dos**; `app_emulador.py` 15/15; `seguridad.py` 450/450 (4 casos 13.3 nuevos) |
| **Sin red, lo pendiente sale solo** (`59ef373`): en el A55 (WebView 157) `navigator.onLine` sigue en `true` con el modo avión y nunca llega «online»: los cambios de un documento cerrado y los archivos pendientes no salían hasta que la persona hiciera algo (venía de antes). Mientras la última petición haya fallado por red, la página le pregunta al servidor cada 4 s (hasta cada 15 s); los binarios pendientes también se suben con el aviso de la API | A55: lo escrito sin red llega en 8 s y el archivo de 1.5 MB sube; `prueba_sin_internet` 32/32 |
| **`prueba_red` 24/24** (`c376921`): «navegador del celular → página y → app en vivo» fallaba desde la 0.17.0 y era de la prueba: el texto sí llegaba en vivo, pero con los renglones como Word se veía partido después de un guion («HOLA-» / «DESDE-CELULAR») y `innerText` no lo encontraba. Ahora lee el texto sin lo que es solo de pantalla y, si falla, dice en qué paso, si llegó al servidor y qué quedó pendiente | 24/24 |
| **PowerPoint: la pestaña Presentación y las secciones** (`1200b5c`, después de la 0.19.0): F5 desde el principio y Mayús+F5 desde la actual; ocultar diapositiva (atenuada y con el número tachado; no sale al presentar salvo si se empieza en ella o un vínculo lleva ahí); presentaciones personalizadas (hacer, copiar, editar qué diapositivas y en qué orden, presentar); ensayar intervalos (con pausa; si se conservan, avance automático); configurar presentación (repetir hasta Esc, usar intervalos, sin animación, qué muestra «desde el principio»); secciones (agregar, nombre, contraer, mover, quitar con o sin diapositivas). Se leen y se escriben como PowerPoint (`show="0"`, `p:custShowLst`, `p14:sectionLst`, `advTm`, `presProps.xml`); antes, borrar una diapositiva quitaba las secciones y las personalizadas del archivo | `presentacion_real.py` (nueva, contra PowerPoint 2024 con `windows/hacer_presentacion.py` y `presentacion_ver.py`) **31/31**: p13 leído igual; sin tocar, cambiado en NOM y hecho en NOM, PowerPoint dice lo esperado y abre los tres sin avisos; validador sin errores. `prueba_presentacion.py` 10/10. En Windows: 12 suites de PowerPoint y la ruedita (fidelidad, modelo, ediciones, rastreo y rastreo en el celular…) sin fallas |
| **Excel: «####» como Excel** (`9707662`): medido en Excel de píxel en píxel (`windows/medir_gatos2.py`, 88 números con y sin «_)»; con Calibri, con Aptos Narrow en la celda y con Aptos Narrow en el libro: los mismos umbrales) y los anchos de Windows (`windows/anchos_gdi.py`): a 11 puntos Excel suma dígito 7, «.» y «,» 4, «-» 5, «$» 7, paréntesis 5, espacio 3 (el hueco «_)» cuenta como espacio, el relleno «*» no ocupa) y el número cabe si eso más 5 px no pasa del ancho de la columna; la fuente de la celda no cambia nada. `js/anchos-excel.js`; con otros caracteres o tamaños, como antes | `comparar_real.py`: **los once libros (x01-x10, t02) con el 100 % del texto de las celdas** (x10 venía con 3 «####» de más). Ojo: con los anchos de Windows de Aptos Narrow (dígito 8) salían «####» de más en x02 y x09: Excel no usa esos |
| **Texto a los dos lados de un dibujo** (`3ade61e`): «Cuadrado, a ambos lados» es lo de siempre en Word y NOM ponía el texto de un solo lado (los flotantes invisibles de CSS no pueden partir un renglón). Medido con 41 dibujos hechos en Word (`windows/hacer_dos_lados.py` → `w17-dos-lados.docx`): cada renglón de la franja tiene un tramo a la izquierda (hasta 9 pt antes del dibujo) y otro a la derecha; Word llena primero el de la izquierda y un tramo donde no cabe la palabra que sigue se queda vacío. NOM: `franjaDe` marca la franja (`medio`), `apartarTexto` deja una marca sin tamaño (`.wd-medio`) en vez del flotante, `cortesWord` llena los dos tramos y `escribirAcomodo` pone el hueco de pantalla (`.wd-hueco`, y en justificado un relleno para que el navegador no estire de un tramo al otro); si el párrafo no se puede acomodar, de un solo lado como antes. `comparar_real.py` parte el renglón de NOM en el hueco | `comparar_real.py`: **w17 67.6 → 92.2 %, w12 86.1 → 90.7 %, t01 95.6 → 96.0 %**, los demás Word igual (w01-w04, w08, w09, w11, w13 100 %…). `prueba_acomodo` con 6 casos nuevos (tramos, el texto no cambia, escribir al final del tramo de la izquierda, se vuelve a partir, copiar a través del hueco, lo descargado) |
| **Excel: crear tablas de datos** (`54413d1`): Datos → Análisis de hipótesis → Tabla de datos… en Más, la ruedita y la paleta: el rango completo (la fórmula arriba a la izquierda, o arriba o a la izquierda en las de una variable) y la celda de entrada de fila, de columna o las dos; cada celda del interior lleva `=TABLA(fila;columna)`, que el importador ya calculaba y guarda como el `dataTable` de Excel. `herramientas.tablaDatos` | `datos_excel.py`: **27 de 27** (3 nuevos iguales a `Range.Table` de Excel 2024: una variable en columna, en fila y dos variables); `libro_excel.py` **7 de 7**: Excel abre lo que descarga NOM con `=TABLE(B1,B2)` en cada celda del interior y los mismos valores |
| **Excel: escenarios** (`cf892a2`): Administrador de escenarios (agregar, modificar, eliminar, mostrar, resumen) en Más, la ruedita y la paleta; reglas de Excel (nombre único, hasta 32 celdas cambiantes sin fórmula, «Creado por … el …», el usuario). Se leen y se guardan en `<scenarios>` (importador de Excel versión 11: `t.esc`). El informe resumen: hoja nueva antes de la de los escenarios, valores actuales y de cada escenario (los que no cambia una celda, con su valor de ahora, como Excel), el nombre definido en vez de $B$1, y el formato medido de Excel (título morado `#800080` con letra blanca de 12 y 10 pt, rótulos en gris `#C0C0C0` con «Celdas cambiantes:» en azul `#000080`, rayas finas y gruesas, la fila de comentarios oculta en un grupo, columnas de valores agrupadas, sin cuadrícula: `showGridLines="0"` al descargar). Hechos con Excel: `windows/hacer_escenarios.py` → `x11-escenarios.xlsx`; Excel los lee con `windows/escenarios_ver.py` | `escenarios_real.py`: **19 de 19** (lee los 4, «Mostrar» igual, los 30 números del resumen, sus 259 propiedades de formato iguales a las del resumen que hizo Excel —menos el ancho de las columnas de valores: NOM lo ajusta a «Valores actuales:», más largo que «Current Values:»—, cambiados y hechos en NOM Excel lee lo mismo y abre sin avisos, el formulario del administrador); `comparar_real.py`: los doce libros de Excel (x01-x11, t02) con el **100 %** del texto, x11 también los fondos |
| **0.19.0** (`ee19278`): `Version.java` 0.19.0 e INTERFAZ 24 (el idioma de cada párrafo, `numLineas` y `sinNumLin`), `nucleo.js`, `sw.js`, APK 36, novedades y `CAMBIOS.md` | ver §3 |

## 2. Cómo seguir

- Empezar por aquí; el orden sigue el de `~/NOM/prompt.md` (punto 1: cómo se VE un archivo; luego 2, 3 y 4).
- Servidor de desarrollo: 8096 con el jar de esta sesión (`trabajo/srv/nom-dev.jar`, ya acepta el origen nuevo). Un
  servidor 0.18.1 (el publicado) puede quedar en 8097 (`trabajo/21nov/srv-viejo/`) para probar «equipo de antes».
- **El A55 estuvo conectado por USB** (`adb devices`: `RZCX41FHPXX`; mismo Wi-Fi que la laptop, 192.168.0.x). Para
  probar sin tocar la app del dueño: compilar con `applicationId = "mx.wack.nom.prueba"` (solo para esa compilación;
  regresar `build.gradle.kts`), `NOM_PAQUETE=mx.wack.nom.prueba` en `paso_origen.py` y `origen_equipos.py`, y
  desinstalar la copia al terminar. `paso_origen.py` desinstala la app: en un teléfono real se niega con el paquete de
  siempre. Con la pantalla apagada el WebView se congela (CDP no contesta): `adb shell input keyevent KEYCODE_WAKEUP`.
  El A55 deja el Wi-Fi prendido en modo avión: `paso_origen.py` también apaga el Wi-Fi y lo regresa al terminar.
- Mediciones y registros de esta sesión: `~/NOM/trabajo/21nov/` (APK de prueba: `app-vieja.apk`, `app-nueva.apk`,
  `app-vieja-prueba.apk`, `app-prueba-a55.apk`; registros de cada prueba).

## 3. Publicación de la 0.19.0 (21-nov, 03:30): hecha

- **Regresión completa de Windows** (`--carriles 10`, 1359 s): 33 de 34 suites bien. La que no, `documento_grande`
  (16/17: el Word descargado sin partir del original), era de la prueba: descargaba con la exportación directa, que no
  espera el original del almacén, y con la computadora cargada a veces no estaba en la memoria. Al buscarlo salieron
  **dos defectos reales del mismo tipo**, arreglados: «Descargar todo» del espacio (`98e588d`: el Word salía sin sus
  imágenes, sin sus símbolos y con 182 letras menos) y «Descargar copia (.nom)» de un documento de este equipo
  (`5fb060b`: solo la referencia de la imagen). Prueba nueva `prueba_zip_espacio.py` (13/13, en la regresión). Después,
  en Windows, `zip_espacio`, `documento_grande`, `editar_real`, `funciones`, `red`, `sin_internet` y `office_ediciones`:
  todas bien.
- **La Pi**: `nomcolab pi actualizar` (respaldo `datos-2026-11-21_042734`), `humo.py` contra `www.nom.qd.je` 34/34;
  `/api/red/info` dice 0.19.0, interfaz 24.
- **GitHub**: `main` hasta `95e903d` (Azure sigue apagada).
- **APK 36**: release `v0.19.0` (SHA-256 `e462344e…`; el enlace fijo de la página baja este).
- **Buscadores**: IndexNow 200 (76 direcciones).
- **A55** (por USB): instalada la 0.19.0 encima de la 0.16.0; el paso al origen propio dio «bien, 12 datos y 35
  registros» (`adb logcat -s NOM:I`) y la app abrió con la sesión, el espacio y los documentos del dueño.

**Intentado y deshecho (contornos, 21-nov):** acomodar como Word junto a un contorno (estrecho, a través) con lo que el
polígono de Word (`wrapPolygon`) le quita a cada renglón, recortado a la caja como `shape-outside` y comprobado renglón por
renglón contra el navegador. w12 subía de 90.7 a 92.6 % pero w06 bajaba de 99.3 a 95.7 %: el párrafo del contorno de w06
quedaba con un renglón menos que en Word (Word aparta más de lo que da el contorno más la distancia) y el de abajo subía
junto a la imagen. Falta medir en Word la regla del contorno (un generador como `hacer_dos_lados.py` con formas de varias
siluetas y «estrecho»/«a través») antes de volver a intentarlo.

## 4. Lo que falta (de `PARIDAD-OFFICE.md`)

Igual que `RELEVO-20NOV-NOCHE.md` §3 y §4, menos lo de esta sesión: en PowerPoint las secciones y casi toda la pestaña
Presentación pasaron a «Sí» (quedan presentar en línea, grabar, subtítulos en vivo y Speaker Coach) y en Excel el «####»
de x10; también en Excel las tablas de datos (crearlas) y los escenarios. Unas 51 filas o partes de fila en «No» o «A medias»
(Word ~23, Excel ~12, PowerPoint ~13, comunes ~11). En la pestaña Datos de Excel quedan el relleno rápido y la previsión.
Cómo se ve (punto 1), lo que queda: Word w12 90.7 % (contornos estrecho y a través: hoy los acomoda el navegador con
`shape-outside` y no corta como Word), w17 92.2 % (cortes de página y un renglón al pie de la franja), t01 96.0 %, w05
98.5 %, w10 90.4 % (tablas dentro de un cuadro de texto y de una nota al pie, una figura en una celda), justificados o
centrados con tabuladores, cuadros de texto, w06; PowerPoint SmartArt (p11 89.5 %), transiciones y Transformar.
