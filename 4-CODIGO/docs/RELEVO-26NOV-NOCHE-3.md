# Relevo del 26 de noviembre de 2026 (noche, tercera sesión): preparar la expo

**Leer este primero.** Sigue a `RELEVO-26NOV-NOCHE-2.md`. Publicada al empezar: **0.23.6** (interfaz 27, APK 48).

## 0. Horario

- Expo: **viernes 27-nov a las 8:00**. Código congelado a las **5:30** como máximo (lo dijo el dueño). Después de
  congelar solo se arreglan errores vistos en el ensayo, cada uno con su prueba.

## 1. Estado (se actualiza en cada hito)

- 18:30 — inicio: relevo creado, suites en vivo lanzadas, barridos con los ayudantes.
- 18:45 — **en la copia de trabajo, sin publicar todavía (será la 0.23.7):**
  - Textos: 49 palabras del glosario que se veían en la interfaz (servidor, cifrado, solo lectura, red local, exportar,
    caracteres) cambiadas por las de todos los días; «En el servidor» del riel → «Tu almacenamiento». Medido con
    `pruebas/interfaz/legibilidad.py` (nuevo): interfaz INFLESZ 76.5, ayuda 80.9, 0 palabras prohibidas en la interfaz.
  - Celular: los íconos que no se entienden solos llevan un nombre corto debajo (`editor.js` CORTAS, `editor.css`);
    «Insertar» del texto con nombre en la computadora. Sin cambiar el orden. Falta verlo en vivo (después de publicar).
  - Importar Office en un espacio: el botón dice «Abriendo…» / «Subiendo… puede tardar un poco», no se puede tocar dos
    veces y sale la barra de carga (antes 8-22 s sin nada en pantalla).
  - Brechas §C.1 (revisadas línea por línea): `cssFuente`, `piezasXL`, `geometrias.js` (imágenes solo del almacén o `data:`),
    `calcularCombinadas` acotado, `CSS.escape` en `lienzo.js`, ciclos de `padre` en `smartart.js`, mapas sin prototipo
    en `ooxml.js`. Prueba nueva `app/src/test/js/office-armado.test.js`: 14 fallas antes, 0 después.
  - Teléfono (§P.3): H5 topes de cuentas por /64 solo con IPv6 de Internet (`Hosts.agruparRedPublica`, JUnit
    `HostsRedPublicaTest`), H6 el id del dueño ya no se distingue (401 igual que llave mala), H7 cabeceras ≤ 32 KB
    (431). Falta compilar el APK y probar H6/H7 en el emulador.
  - Guías: fuentes primarias, Zeigarnik fuera, contradicción de los íconos resuelta, `DISENO-UI.md` §16 y
    `GUIA-FACIL-DE-USAR.md` §4 con lo cambiado. Hoja para el dueño: `ENTREGA/Prueba_con_personas.docx`.

- 18:50 — **0.23.7 publicada en la Pi** (humo 39/39, con la compresión). Todas las suites en vivo en verde sobre la
  0.23.7 (ver §3). Commit `044f6ef` (sin subir a GitHub todavía).
- 19:25 — **0.23.8 publicada en la Pi**: al presentar, botones y contador se esconden a los 3 s sin mover el mouse
  (comprobado en vivo). Commit `ac29f75`.

### La presentación de la expo dentro de NOM (pedido del dueño, 26-nov 19:00)

Decisión del dueño: **nada de importar PDF a los editores** (son para crear y editar; un visor de PDF sería otro
módulo, después). En su lugar, la presentación del proyecto (que era HTML impreso a PDF) se pasó a una **presentación
de NOM editable** y se presenta desde el apartado de presentaciones.

- `7-PRESENTACION/fuentes/a_nom.py`: cada diapositiva = fondo sin texto (dibujado al doble) + cada texto como cuadro de
  texto de NOM en su lugar exacto (Nunito, tamaño, peso, color, espaciado, mismos cortes de línea). `comparar_nom.py`
  (NOM en modo Presentar contra el PDF), `techo.py` y `fidelidad.py` (tolerancia de 1 px) para medirlo;
  `descargar_nom.py` baja el `.nom` y el `.pptx` por la misma ruta que el botón «Descargar».
- Dónde está: cuenta `expo_profe`, espacio «Expo NOM», «Presentación del proyecto (NOM)». Archivos:
  `7-PRESENTACION/5IM7_WACK_Presentacion.nom` y `.pptx` (15 diapositivas; el .pptx abre con el texto editable).
- **Fidelidad medida** (15 diapositivas, 1280 × 720, tolerancia de 1 px): NOM contra el PDF **98.60 %** (peor 97.44 %);
  NOM contra el HTML original **98.92 %**; el techo (el HTML contra su propio PDF, dos programas distintos) es 99.67 %.
  Sin errores de JavaScript. Lo que falta para el 100 % son corrimientos de 1 px del texto.
- Para presentar: abrir el documento → «Presentar» (F5) → flechas, clic o deslizar; Esc sale.

- 19:50 — **0.23.9 publicada en la Pi** (lo descargado en HTML/Markdown/SVG y CSV largos; `prueba_armados.py` 9/0, antes
  3/6). Todas las suites en verde sobre la 0.23.9 (`trabajo/expo27/vivo4/`: funciones 202/205/202 en Chromium, celular y
  Firefox, e2e 33, sin Internet 32, red 24, presentación 10, diag1/85/pptx2/pres1 sin errores, ayuda 24, axe 28,
  accesibilidad 7, rastreo 632 acciones sin errores).
- 22:25 — **Fin de la sesión (límite de uso).** Estado exacto en §6. **La siguiente sesión empieza en §6.**

## 2. Matriz de funciones (A)

Inventario completo (332 funciones visibles, con archivo:línea y la prueba de cada una) en
`trabajo/expo27/barridos/ds-inventario.txt`; 31 con prueba de que hacen lo correcto, 183 solo con el rastreo (que no
truenen) y 118 que el inventario marcó «sin prueba» (varias sí están cubiertas: invitar con código, unirse, escribir a la
vez, comentar y sugerir están en `prueba_funciones`). Las más visibles que de verdad no tenían prueba quedaron en
`prueba_demo.py` (importar desde el botón, «Descargar» a Word/Excel/PowerPoint desde la barra, permisos desde la lista de
miembros, buscar en los documentos).

| Zona | Computadora (Chromium / Firefox) | Celular (navegador) | App Android | Papeles | Sin Internet | Cifrado | Office |
|---|---|---|---|---|---|---|---|
| Texto | funciones, rastreo / funciones | funciones --movil | red (interfaz del APK) | funciones (solo ver), demo | sin_internet | e2e | funciones (docx), demo, armados |
| Hoja | funciones, rastreo | funciones --movil | red | funciones | sin_internet | e2e | funciones (xlsx), demo |
| Presentación | funciones, presentacion, pres1, pptx2, rastreo | funciones --movil | red | funciones | — | e2e (proyector) | pptx2, demo, comparar_nom |
| Diagrama | funciones, diag1, rastreo | diag_movil | — | — | — | — | diag1 (pptx/svg), armados |
| Espacios, invitar, permisos | funciones, 85, demo | funciones --movil | red | demo (solo ver ↔ editar), lectores | — | e2e (llaves) | — |
| Cuentas y ajustes | funciones, cuentas_equipo | funciones --movil | red (registro desde la app) | — | — | e2e | — |
| Teléfono que comparte | emulador (H6/H7 hoy), telefono_emulador | — | — | — | — | — | — |
| Accesibilidad | axe 28/0, accesibilidad 7/7 | — | — | — | — | — | — |
| Tiempos | prueba_tiempos (normal y 3G) | — | emulador (arranque) | — | — | — | — |

Sin prueba automática todavía (para el ensayo de §4b): teléfono que comparte con dos equipos reales, el A55 con el
APK nuevo, y Office real (Word/Excel/PowerPoint en Windows) abriendo lo descargado hoy.

## 3. Resultados de las pruebas en vivo (B)

**Con la 0.23.7** (salidas en `trabajo/expo27/vivo2/`): humo 39/39 · funciones 202/0 · celular 205/0 · Firefox 202/0 ·
extremo a extremo 33/0 · sin Internet 32/0 · red 24/0 (después de poner al día el texto que busca) · presentación 10/0 ·
pres1, pptx2, diag1 y 85 sin errores de JavaScript · ayuda 24/0 · axe 28/0 · accesibilidad 7/7 · rastreo 632 acciones
bien, 0 mal (los «?» son paneles que cambian al usarlos).

Contra www.nom.qd.je con la 0.23.6 (antes de los cambios de esta sesión), salidas en `trabajo/expo27/vivo/`:
`humo.py` 36/36 · funciones 202/0 · funciones en celular 205/0 · extremo a extremo 33/0 · sin Internet 32/0 ·
red 24/0 · presentación 10/0 · pres1 y pptx2 sin errores de JavaScript. Rastreo: un «?» (Hoja › Ordenar y filtrar ›
Fijar hasta A: el menú no volvió a aparecer al reabrirlo; por revisar).

## 4. Tiempos medidos (D)

`pruebas/interfaz/prueba_tiempos.py` (nuevo) contra www.nom.qd.je, laptop en la red de la casa, Chromium:

| Qué (conexión normal) | Medido | Meta |
|---|---|---|
| Primera carga sin caché / con caché (hasta ver «Entrar») | 1.18 s / 0.30 s (4.2 MB bajados) | — |
| Entrar → inicio | 0.83 s | — |
| Abrir texto, hoja, presentación, diagrama | 0.20 / 0.22 / 0.25 / 0.19 s | < 1 s ✔ |
| Abrir Word 2 MB, PowerPoint 1.5 MB, Excel 0.4 MB importados | 0.37 / 0.45 / 0.54 s | — |
| Tecla → letra en pantalla (p75 / peor) | 16 / 24 ms | < 50 ms ✔ |
| Cambio de otra persona → visible (mediana de 5) | 180 ms | < 300 ms ✔ |
| Cambiar de diapositiva / presentar (F5) | 61 ms / 84 ms | ✔ |
| Volver de sin red → llega a la otra persona | 0.38 s | — |
| Importar en la página (Word 2 MB, PowerPoint 1.5 MB) | 0.14 / 0.27 s | — |
| **Subir lo importado** (Word 2 MB, PowerPoint 1.5 MB, Excel 0.4 MB) | **21.7 / 14.9 / 10.2 s** | depende de la subida de la red |
| Tarea más larga | 270 ms al abrir los Office grandes | **< 100 ms ✘** |

Lo que no cumple: (1) la subida va una por una (original + cada imagen) y está limitada por la subida de la casa;
ahora al menos se ve «Subiendo…». (2) Tareas de 150-270 ms al pintar los documentos grandes importados (no en los
normales).

3G lenta (perfil «Slow 3G» de Chromium: 2 s de latencia, 400 kb/s), **antes de comprimir**: primera carga 111 s
(4.2 MB), con caché 2.1 s, entrar 2.3 s, abrir un documento ~2.0 s (una ida y vuelta de 2 s), tecla → pantalla 24 ms,
cambio en vivo 1.3 s, ninguna tarea > 100 ms. Desde la 0.23.7 la página va comprimida (`ooxml.js` 611 → 188 KB):
volver a medir la primera carga.

## 4b. Listos para la expo (G) y ensayo al azar (B.3)

**La nube (Pi en casa) es un solo punto de falla.** Si en la escuela no carga www.nom.qd.je (se fue la luz o el
Internet de la casa): **plan B** = el teléfono comparte sin Internet. En la app: «¿Dónde trabaja tu equipo?» →
*Compartir desde este teléfono* → Encender; la laptop se conecta al punto de acceso del teléfono y abre la dirección
`http://IP:puerto` que muestra la app (o escanea el QR). Ensayarlo una vez antes (los documentos de la nube NO están en
el teléfono: llevar en la laptop `7-PRESENTACION/5IM7_WACK_Presentacion.nom` y los Office de ejemplo para importarlos
ahí).

**Cuentas de demostración** en el sitio real (contraseña en `trabajo/expo27/demo.txt`, fuera de git): `expo_profe`
(propietaria del espacio «Expo NOM»: texto, hoja, presentación, diagrama, un Word de 0.6 y 2 MB, un Excel de 0.4 MB, un
PowerPoint de 1.5 MB y **la presentación del proyecto**) y `expo_alumna` (puede editar). Crear además una cuenta en vivo.

**Antes de salir:** cargar laptop y teléfonos; instalar el APK 0.23.10 en el A55 (y probar que abre la nube); en la
laptop abrir www.nom.qd.je y recargar con **Ctrl+Mayús+R** (Cloudflare deja guardados los .js 4 h: sin recargar podría
quedar una versión de antes; lo arregla el «Browser Cache TTL» de Cloudflare, pendiente del dueño).

**Números para responder** (todos medidos hoy en el sitio real): suites 202 + 205 + 202 (Chromium, celular, Firefox)
funciones bien y 0 mal; extremo a extremo 33/0; sin Internet 32/0; red 24/0; accesibilidad (axe) 28/0; rastreo de 632
acciones sin errores. Tiempos: abrir un documento 0.2 s, tecla → pantalla 16 ms, cambio de otra persona 0.18 s, volver
de sin red 0.4 s. Lenguaje: INFLESZ 76.5 («bastante fácil»). Presentación en NOM: 98.6 % igual al PDF. Lo que falta,
dicho con honestidad: módulo de OCR (0 %), macros de Office, visor de PDF (después de la expo), prueba con personas.

**Ensayo al azar** (que alguien que no sea quien presenta elija números sin avisar; en la laptop, el A55 y otro
equipo, en el sitio real). Lo que falle es prioridad absoluta:

1. Recibo de mi trabajo
2. Historial de versiones y restaurar
3. Importar un Word, un Excel y un PowerPoint de verdad
4. Descargar en Word/Excel/PowerPoint y abrirlo en Office
5. Nueva diapositiva, imagen, animación y Presentar (F5)
6. Escribir los dos a la vez en el mismo texto
7. Buscar en los documentos del espacio por lo que dicen
8. Proyección en el salón con código de 6 letras
9. Cortar el Wi-Fi, escribir y volver (se junta solo)
10. Modo sugerencia: proponer y aceptar
11. Fórmulas en la hoja (=SUMA, =PROMEDIO, =SI) y una gráfica
12. Ajustes: tema oscuro y letra «Muy grande»
13. Ctrl+K: buscar un documento y una acción
14. Comentar y responder
15. Ayuda «?» de un editor
16. Diagrama listo (organigrama) y conector entre figuras
17. Invitar con código QR y entrar desde otro celular
18. Enlace para ver (sin cuenta) en otro navegador
19. Compartir desde el teléfono (sin Internet) y entrar desde la laptop
20. Cambiar un miembro a «Solo puede ver»
21. Crear un espacio y un documento de cada tipo
22. Crear un espacio cifrado de extremo a extremo
23. Ramas: copia sin red y fusionar
24. Chat del espacio con un adjunto

## 5. Para el dueño

- Prueba con personas: `ENTREGA/Prueba_con_personas.docx` (tareas de la guía §5 y SUS).
- Lo manual de antes: Browser Cache TTL en Cloudflare, `reposo.key` (copia), correo ARCO, DMARC, TLS 1.2.
- Decisiones de diseño de los espacios cifrados (plan privado §P.2): se recomiendan **después** de la expo.


## 6. Estado al cerrar y qué sigue (LEER PRIMERO)

### 6.1 Dónde quedó todo (26-nov, 22:25)

- **Pi: 0.23.9.** GitHub: `origin/main` sigue en el relevo anterior; **6 commits locales sin subir** (`3c051cf` …
  `ff07d18`). APK publicado: el de antes (0.23.6, APK 48); el A55 tiene uno viejo.
- **Sin publicar, en `main` (commit `ff07d18`, compila; JUnit 0 fallas):**
  1. `16ca8de` Buscar en los documentos encuentra el texto de un Word importado (`BuscarTextoImportadoTest`;
     en vivo `prueba_demo.py` da 9/1 con la 0.23.9 y debe dar 10/0 con la 0.23.10).
  2. `ff07d18` **Solo la misma versión trabaja con el servidor** (orden del dueño, 26-nov 20:00): `Version.misma` y
     `Version.sinVersion` (núcleo), `FiltroVersion` (computadora), la misma revisión al principio de
     `ApiTelefono.atender` (teléfono), cabecera `X-NOM-Version` en `api.js` (pedir, binario, sondeo), `acceso.js`,
     `editor.js` (al cerrar) y `?v=` en el canal en vivo; CORS de los dos permite la cabecera; la sincronía entre NOM la
     manda (`ServicioSincronia.abrir`). Con otra versión: **426** con el aviso; la página lo trata como «sin conexión»
     (lo hecho queda en el equipo y se manda al actualizar) y `api.alVersion` (app.js) muestra «Actualiza NOM» con
     «Recargar la página» (web) o «Descargar la app nueva» (app). Se deja sin versión: `GET /api/red/info`, salir,
     enlace para ver, proyección, verificar recibo, llave de pases. **Por qué** (para la presentación): una versión
     distinta reescribe al editar lo que no conoce (estilos, campos nuevos de Office, animaciones), dibuja distinto y la
     persona «arregla» lo que no estaba mal, o no trae arreglos del formato (reloj, llaves). Las páginas servidas por la
     nube o por el teléfono siempre traen la versión del servidor: solo les toca a pestañas abiertas durante una
     actualización (o archivos viejos en la caché del navegador) y a la app de Android.
  3. `ff07d18` **Dónde se guarda:** al crear o importar dentro de un espacio, «En «espacio»» o «Solo en este navegador /
     teléfono» (`vista-espacios.js nuevoDocumento`).
  4. Pruebas puestas al día para hablar con la versión del servidor (`nomprueba.version()`, `humo.py`, `prueba_e2e`);
     `humo.py` revisa el 426: **antes** (0.23.9) 40/42 con las 2 fallas del 426 (`trabajo/expo27/version-antes.txt`).

### 6.2 Lo primero (antes de congelar a las 5:30; expo a las 8:00)

1. **Riesgo del candado de versión con la caché de Cloudflare:** los `.js` llegan con `max-age=14400` (Browser Cache TTL
   de Cloudflare). Con service worker la página los vuelve a pedir (`cache: no-cache`), pero un navegador que entra por
   primera vez podría traer `.js` viejos → quedaría con «Actualiza NOM» aunque recargue. Antes de publicar: que el dueño
   ponga **Browser Cache TTL = «Respect existing headers»** en Cloudflare, o agregar `?v=VERSION` a los `<script>` y
   `<link>` en `PaginasApi.portada` (solo la página servida; el APK no pasa por ahí). Probarlo con un perfil de navegador
   nuevo después de publicar.
2. Publicar la 0.23.10: `mvn package` → `2-PAGINA-Servidor/nom.jar` → `4-CODIGO/servidor-propio/nom pi actualizar` →
   `humo.py` (debe dar 42/42), `prueba_demo.py` (10/0), `prueba_armados.py`, `prueba_funciones` (+ `--movil` y Firefox),
   `prueba_red` (ojo: la interfaz del APK en la prueba se carga de archivos con la MISMA versión), `prueba_e2e`,
   `prueba_sin_internet`, `rastreo`, `prueba_presentacion`. Probar a mano: una pestaña abierta antes de publicar → debe
   salir «Actualiza NOM» y «Recargar» la arregla.
3. **APK 0.23.10 (versionCode 52)** con `assembleRelease` (certificado `2ef146f6…`) → `3-APP-Android/NOM.apk` →
   `android/publicar-apk.sh` → **instalar en el A55** (conectado por USB: `adb -s RZCX41FHPXX install -r …`). Con el
   candado, **un APK viejo ya no trabaja con la nube**: el A55 y cualquier teléfono de la demo deben tener la 0.23.10.
   Probar el teléfono que comparte con el emulador o el A55 + la laptop (misma versión) y con versión distinta (426).
4. `git push` (antes: `git diff origin/main..HEAD` sin nombres de herramientas ni detalle de seguridad) y release.
5. **Office real** (PC con Windows por SSH, `user@100.65.29.5`, contesta): `pruebas/office/office_abre.py` falló con
   «resultado.json no existe» (¿sin sesión de escritorio en Windows?). Abrir ahí `7-PRESENTACION/5IM7_WACK_Presentacion.pptx`
   y `trabajo/expo27/demo-ui/descargado.{docx,xlsx,pptx}`.
6. Ensayo al azar (§4b) en la laptop, el A55 y otro equipo; lo que falle, primero.

### 6.3 Pedidos del dueño del 26-nov (22:20) — trabajo para la siguiente sesión

Avances ya lanzados (solo lectura, sin cambios): una revisión de los 6 archivos de prueba del dueño contra el importador →
`trabajo/expo27/`; otra busca la causa de BUSCARV, la imagen que se encoge en los bordes y los
botones que desaparecen en accesibilidad → `trabajo/expo27/barridos/ds-bugs-dueno.txt`. Confirmar cada hallazgo en el
código antes de arreglar.

**Antes de la expo si da tiempo (son errores que se ven):**
- Los **6 archivos del dueño** en `~/NOM` (fuera de git, no subirlos): `word_ooxml_stress_test_max.docx`,
  `test_features.docx`, `excel_ooxml_stress_test_max.xlsx`, `test_features.xlsx`,
  `powerpoint_ooxml_stress_test_max (1).pptx`, `test_features.pptx`. Importarlos en el sitio real, ver qué no se lee o no
  se dibuja, compararlos con Office real (Windows) y decir si el archivo mismo está mal armado.
- **BUSCARV** «da un resultado incorrecto cuando la condición es verdadera» (coincidencia aproximada) y **fórmulas que
  se duplican**; **imagen que al arrastrarla se encoge contra las orillas**; **botones que desaparecen al elegir
  opciones de accesibilidad** (contraste alto, letra grande, colores forzados). Cada uno con su prueba.
- **Imágenes que no cargan completas**: revisar con los archivos del dueño y con el almacén de binarios (precarga,
  reintento, `srcImgOk`).

**Prueba a fondo (después de lo anterior), web y en el A55 por adb (WebView con CDP: `pruebas/reales/cdp.mjs`):**
colocar imágenes y arrastrarlas a todos lados (orillas, fuera de la hoja, girar, recortar), combinaciones raras, cada
tipo de «Insertar» de los 4 editores, cada tamaño de página/diapositiva, letra, tabla y celda; deshacer/rehacer en
medio; dos personas a la vez; sin red a la mitad. Escribir lo que se encuentre como scripts reutilizables
(`pruebas/interfaz/exploracion_*.py`).

**Después de la expo (cambian formatos o la interfaz; planearlo con el dueño):**
- **Ajustes:** atajos de teclado personalizables; en el celular, vibración (háptica) y atajos; elegir **ruedita o la
  lista clásica** (menú contextual de siempre) como preferencia.
- **Idioma:** tomar el del sistema (hoy todo está en español); traducción de la interfaz (archivos de textos por idioma);
  **detectar el idioma del documento** y corregir la ortografía desde la ruedita (sugerencias al tocar una palabra).
- **Accesibilidad completa** (que no haya demanda): WCAG 2.2 AA en todas las pantallas (axe ya da 28/0 en las que se
  prueban: ampliar a editores con contenido, lector de pantalla NVDA/TalkBack, solo teclado, colores forzados, zoom
  400 %), y lo que pide la ley mexicana (Ley General para la Inclusión de las Personas con Discapacidad) — investigar con
  fuentes primarias antes de afirmar.
- **Archivos más ligeros:** medir cuánto pesa un `.nom` vacío y por qué (cabecera, diccionario, JSON del estado, ids de
  CRDT por letra); proponer un formato binario (varint, tablas de estilos por número, ids comprimidos) con su versión de
  formato y lectura de los viejos. Lo que pidió el dueño («cada letra un bit y cada formato un byte») no es posible tal
  cual (una letra necesita al menos ~5-8 bits y la edición en vivo necesita un id por letra), pero sí se puede acercar
  mucho: explicarle con números.
- **Imágenes lo más ligeras posible sin perderlas:** al importar, convertir a WebP/AVIF con calidad alta, tamaño según
  cómo se ve, original solo para descargar a Office; carga diferida y reintentos.
- **Lo que nos hace distintos de Google Docs** (para la presentación): sin Internet con un teléfono que comparte, espacios
  cifrados de extremo a extremo, recibo de quién escribió qué, fidelidad con Office medida contra Office real, gratis y
  libre (GPL-3.0), sin IA que escriba por ti, funciona en la red de la escuela. Preparar 1 diapositiva y una demo de 1 min.
- **Seguridad pedida:** qué datos y uuids ve cada quien (revisar cada respuesta de la API por papel), sanear cada campo
  que entra (barrido por ruta), pantallas de error propias para cada caso (404, 413, 426, 429, sin red, servidor caído),
  e índices en las consultas principales (H2/JDBC: `RepositorioUsuariosJdbc`, revisar `schema-h2.sql`).
- Tareas largas de 150-270 ms al abrir Office grandes; visor de PDF (módulo aparte, decisión del dueño); decisiones de
  §P.2 del plan privado.

### 6.4 Archivos de esta sesión

`trabajo/expo27/`: `vivo*/` (salidas de las suites por versión), `tiempos/`, `pres/cmp` (fidelidad de la presentación),
`barridos/` (barridos de código), `demo.txt` (contraseñas de demo, privado), `office/` (Office grandes), `herr/`
(generadores). Scripts nuevos en el repo: `pruebas/interfaz/{legibilidad,prueba_tiempos,prueba_armados,prueba_demo}.py`,
`app/src/test/js/office-armado.test.js`. Presentación: `7-PRESENTACION/fuentes/{a_nom,comparar_nom,techo,fidelidad,descargar_nom}.py`
(la carpeta 7-PRESENTACION no está en git, por decisión anterior).
