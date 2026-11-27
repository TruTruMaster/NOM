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
