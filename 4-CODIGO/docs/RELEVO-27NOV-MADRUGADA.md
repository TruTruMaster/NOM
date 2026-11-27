# Relevo del 27-nov-2026 (madrugada, antes de la expo de las 8:00)

Orden del dueño (26-nov, 22:30): toda la lista de `~/NOM/prompt-expo-2.md` lista y publicada antes de las 5:00. Lo que a
las 4:15 no pase sus pruebas en vivo no se publica (queda en una rama `pendiente/<tema>`).

## 1. Bitácora

- 22:35 — 0.23.10: la portada pide cada programa y hoja de estilos con `?v=VERSION` (`PaginasLimpias.conVersion`): una
  caché (Cloudflare o el navegador) ya no puede entregar un `.js` viejo con la versión nueva. Publicada en la Pi.
  Humo 42/42 (el canal en vivo de la prueba no mandaba la versión: arreglado en `humo.py`).
- 22:40 — APK 0.23.10 (versionCode 52, certificado `2ef146f6…`) instalado en el A55 y en el emulador; el A55 entra a la
  nube con la cuenta de la expo.
- 22:45 — Presentación: diapositiva nueva «Lo que nos hace distintos de Google Docs» (16 diapositivas). En NOM
  (cuenta `expo_profe`, espacio «Expo NOM») mide 98.48 % igual al PDF (tolerancia de 1 px; peor 96.89 %). Office real
  abre el `.pptx` descargado sin avisos.
- 23:00 — Los 6 archivos del dueño contra Office 2024 real (Windows): los tres `*_stress_test_max` están dañados
  (Word solo abre el .docx reparándolo; Excel y PowerPoint no abren el .xlsx ni el .pptx). Los tres `test_features`
  abren bien en Office y en NOM; diferencias medidas en §3.
- 23:15 — Tanda 1 publicada en la Pi (compilada desde una copia limpia, `~/NOM-pub`): atajos de teclado que se cambian,
  vibración, ruedita o lista clásica, idioma (Automático/Español/English), BUSCARV/BUSCARH/COINCIDIR/BUSCARX como Excel,
  fórmulas que ya no se duplican, imágenes que no se encogen contra la orilla, mensajes de error propios, página propia
  de «no existe», índice de sesiones vencidas, Word sin hoja en blanco antes de un salto de sección.
- 23:10 — **El módem de la casa le cortó a la laptop el camino directo a Cloudflare** (ni ping a 1.1.1.1; el A55 y la
  Pi sí llegan). Las pruebas salen por el puente SOCKS a la Pi: `NOM_PROXY=socks5://127.0.0.1:1080` (nomprueba.py y
  humo.py lo entienden). En la escuela no aplica (otra red).

## 2. Qué falta (se actualiza en cada tanda)
- 23:20 — Tanda 1 en vivo (por el puente): humo 42/42 · demo 10/0 · armados 9/0 · funciones 202/0 · celular 205/0 ·
  Firefox 202/0 · extremo a extremo 33/0 · sin Internet 32/0 · red 24/0 · presentación 10/0 · ayuda 24/0 · pres1,
  pptx2, diag1 y 85 sin errores de JavaScript. `exploracion_novedades.py` 19/0 (lista clásica, idioma inglés, BUSCARV
  como Excel, barra de fórmulas sin duplicar, figura contra la orilla). `exploracion_errores.py` antes 3/9 (JSON crudo,
  «error 5xx» a secas). Office real abre sin avisos lo que descarga NOM (Word, Excel, PowerPoint y la presentación).
- 23:25 — Barrido de la API: qué ve cada papel en la API (13 puntos; detalle SOLO en el plan privado §Q.2). Arreglado hoy el
  chat (`Textos.limpioConRenglones`); lo demás, con su decisión, en §Q.2.
- 23:30 — Teléfono que comparte con el A55 real + la laptop (sin Internet, por el Wi-Fi): entrar con código, escribir
  en los dos sentidos ✔; candado de versión en el teléfono (otra versión o sin versión → 426) ✔. Encontrado y
  arreglado: el QR para unirse no salía al pasar de la nube a «compartir desde este teléfono» (`union.js`, `api.js`).
- 23:35 — Logotipo horizontal decía «NOM-COLAB» (el proyecto se llama NOM desde el 25-nov): ahora solo NOM; la
  presentación se volvió a generar y subir (cuenta `expo_profe`).
- 23:40 — Imágenes en los 4 editores (`exploracion_imagenes.py`): no cambian de tamaño al arrastrarlas contra ninguna
  orilla ni al deshacer/rehacer. Encontrado y arreglado: la ruedita cortaba los submenús a 8 opciones (en la hoja,
  «Más › Otras» perdía Imagen, Fijar, Formato condicional, Tabla dinámica…): ahora 7 + «Más».
- 23:42 — El asistente de edición se quedó sin cuota: se cambió a la otra cuenta (orden del dueño).
- 23:50 — Tanda 2 publicada: accesibilidad WCAG 2.2 AA (`docs/ACCESIBILIDAD.md`), Word importado con sus campos de
  pie de página, casillas, referencias e índice como Word, ruedita sin opciones perdidas, QR del teléfono, chat.
- **El `.nom` ya es compacto** (medido en vivo con `medir_nom.py`): vacío 41-59 bytes; texto real 0.61 bytes por letra
  (≈5 bits) y ≈3 bytes por cada cambio de formato; Office importado: Word 16 KB de 44, Excel 33 KB de 47, PowerPoint
  66 KB de 167. Guarda una foto comprimida (no los identificadores del CRDT), así que un formato v4 no ganaría casi
  nada y cambiar el formato hoy arriesgaría la expo: no se hizo (respuesta para el dueño en §3).
- 00:00-03:20 — **Pausa por el límite de uso** (la sesión se detuvo ~3 h). Siguieron en segundo plano: PowerPoint y Excel y el
  corrector de ortografía (B4).
- 03:20 — Revisado y guardado: PowerPoint/Excel como Office (líneas, punteados, sombras, radar, burbujas, área apilada,
  etiquetas de pastel/dona, nombres de una hoja); ortografía desde la ruedita (nspell + diccionarios libres, se bajan
  solo al usarlos; la opción aparece solo con una palabra mal escrita, para no mover los 8 lugares de la ruedita);
  imágenes que se vuelven a pedir al regresar la red o al llegar la llave (hallazgos del barrido en
  `trabajo/expo27/barridos/ds-imagenes-carga.txt`; quedan para después: verificar la huella al guardar, TIFF/EMF+
  que el navegador no dibuja, recorte de Word con la marca de espera de 1×1).
- 03:22 — **0.24.0 publicada** (versionCode 53). Tanda 2 en vivo (antes de la 0.24.0): todo en verde salvo axe 55/11
  (avatar con colores forzados + tema oscuro: regla CSS agregada para la versión final) y rastreo 638/4 (opciones de
  «Más › Otras» que antes eran inalcanzables y ahora el rastreo alcanza en el 4.º nivel: son submenús, no errores).
- `exploracion_errores.py` con la 0.24.0: **12/0** (antes 3/9).
- 03:28 — APK 0.24.0 (versionCode 53, certificado `2ef146f6…`) en el A55 y el emulador; release `v0.24.0` en GitHub
  (SHA-256 `419a36d8…`); `git push` hecho. El A55 quedó en NOM en Internet, cuenta `expo_profe`, idioma **Español**
  (su sistema está en inglés: con «Automático» la app salía en inglés) y abre la presentación, el PowerPoint del dueño
  y la hoja sin avisos (`reales/a55_expo.py` 3/0).
- Los 6 archivos del dueño quedaron importados en «Expo NOM» (`expo_profe`): los 3 `test_features` sin errores ni
  imágenes rotas (Word 9 páginas como Word, Excel 10 hojas visibles, PowerPoint 13 diapositivas).

## 2. Qué quedó fuera y por qué

- **Formato `.nom` v4:** no hace falta (ya pesa 0.61 bytes por letra) y cambiar el formato la noche antes era el mayor
  riesgo. Imágenes a WebP/AVIF al importar: no se hizo (el original de Office se guarda para descargar igual; convertir
  sin perder fidelidad necesita su propia prueba contra Office real).
- Imágenes (barrido): verificar la huella al guardar en el equipo, TIFF/EMF+ que el navegador no dibuja (sale vacío),
  recorte de Word calculado con la marca de espera (1×1).
- Seguridad (plan privado §Q.2): ids de otros miembros que ven lectores e invitados, `usoPropietario`, binarios de otro
  documento con el enlace en espacios cifrados, validar el formato de las llaves.
- Diferencias finas que quedan en el Word del dueño: índice con números de página guardados por Word, partición de
  la tabla con fila de alto exacto, «Distribuida». En PowerPoint: sombras con desenfoque aproximado.
- Rastreo: 4 opciones de «Más › Otras › Más» de la hoja se marcan porque el rastreo llega a un 4.º nivel (son
  submenús); `prueba_axe` ampliada: queda el texto de muestra «Escribe aquí» en las miniaturas con tema oscuro.
- Traducción al inglés: quedan textos sueltos («hace un momento», la barra de estado del texto, la ayuda y lo legal).

## 3. Para el dueño (expo)

- **Mostrar:** la presentación (`expo_profe` › «Expo NOM» › «Presentación del proyecto (NOM)», Presentar/F5) y los 3
  `test_features` (Word, Excel, PowerPoint). Los 3 `*_stress_test_max` están dañados según Office real (Word solo
  abre el .docx reparándolo; Excel y PowerPoint no abren los otros dos): NOM los abre, pero no son buen ejemplo.
- **La laptop de la casa no llega a www.nom.qd.je** desde el 26-nov 23:10 (el módem le cortó el camino a Cloudflare; el
  teléfono y la Pi sí llegan). En la escuela es otra red. Si en casa hace falta: `NOM_PROXY=socks5://127.0.0.1:1080`
  para las pruebas, o reiniciar el Wi-Fi/módem.
- **Archivos más ligeros («cada letra un bit, cada formato un byte»):** medido en vivo, un `.nom` vacío pesa 41 bytes;
  texto real, 0.61 bytes por letra (≈5 bits: no se puede 1 bit, un alfabeto de 60+ letras necesita al menos ~5-6 bits
  sin contexto) y ≈3 bytes por cada cambio de formato (cerca de «un byte»); un Word de 44 KB queda en 16 KB.
- **Idioma:** «Automático» toma el del equipo; en un teléfono en inglés la app sale en inglés. Para la demo, Ajustes ›
  Idioma › Español (ya está así en el A55).
- Antes de salir: recargar www.nom.qd.je con Ctrl+Mayús+R en la laptop de la expo.

## 4. Cacería de último momento (27-nov, 05:00-05:45; lo que reportó el equipo)

- **Ordenar A → Z con una tabla elegida revolvía toda la hoja** (y los números de fila): el menú ordenaba la vista
  de toda la hoja y además, al abrirse, cambiaba la selección a la columna entera. Ahora ordena SOLO la tabla con
  formato o el rango que estaba elegido, moviendo las celdas como Excel (encabezado detectado, fila de totales fuera,
  fórmulas ajustadas a su fila, Ctrl+Z lo deshace). Con una sola celda o columna sigue el orden de la vista (solo para
  ti). `exploracion_orden_tabla.py`: antes 5/3, ahora 8/0.
- **Imágenes flotantes en el enlace para ver** («visita web»): salían en línea, sin su lugar ni su giro. La exportación
  a HTML (que usa el enlace para ver) ahora las pone libres desde su párrafo y giradas.
  `exploracion_imagen_flotante.py` (dos personas a la vez, vista sin páginas, girar, deshacer, quien abre después y el
  enlace para ver): 13/0 (una corrida de cuatro tardó más de 2.5 s en llegarle a la otra persona; no se repitió).
- **Fotos giradas o recortadas deformadas** cuando llegaban del almacén después de la marca de espera: se recalcula su
  proporción al cargar.
- La portada pide los archivos con la versión y una marca del arranque (`0.24.0-xxxx`): un arreglo publicado con el
  mismo número ya no se queda en la caché del navegador.
- Pista del lienzo que se salía de la pantalla en el celular: ahora se parte en renglones.
- Barridos sin hallazgos: `exploracion_layout.py` (28/0, computadora y celular) y `exploracion_fotos.py` (los 15
  documentos de «Expo NOM» con todas sus fotos cargadas, 30/0).
- APK 0.24.0 vuelto a compilar con todo esto (mismo número; release actualizada) e instalado en el A55.
