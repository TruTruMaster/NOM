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
