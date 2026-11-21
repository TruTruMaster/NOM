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

## 4. Lo que falta (de `PARIDAD-OFFICE.md`)

Igual que `RELEVO-20NOV-NOCHE.md` §3 y §4 (unas 56 filas o partes de fila en «No» o «A medias»: Word ~23, Excel ~14,
PowerPoint ~15, comunes ~11). Esta sesión no cambió filas de la paridad: lo de arriba es de la app, la red y las pruebas.
