# Plan: video y audio en NOM (ligeros, en la nube)

Decidido por el usuario el 17 de noviembre de 2026 (lo que dijo, resumido):

- Los archivos del usuario **tienen que estar en la nube** para abrirlos desde cualquier lugar.
- Convertir los formatos comunes (MP4, MP3…) a lo **más ligero posible sin perder calidad visible**, que se
  descomprima en tiempo real al reproducir, tanto video como audio, con archivos propios de NOM.
- **Límites de subida que no castiguen**; la Raspberry Pi recibe un disco de **2 TB** (entre el 17 y el 19 de noviembre;
  la noche del 17 todavía tenía solo su tarjeta de 58 GB): «decide lo mejor para MUCHOS usuarios».
- **Azure ya no cuenta** (no se diseña para su cuota de 50 MB).
- Al descargar: **lo más ligero para el ancho de banda** del servidor.

## Decisiones técnicas

| Tema | Decisión | Por qué |
|---|---|---|
| Códecs | Video **AV1**, audio **Opus**, en un contenedor WebM (Matroska) que NOM guarda como archivo propio | Son los más eficientes que los navegadores decodifican solos en tiempo real (por hardware en la mayoría de los teléfonos Android recientes). Un códec propio escrito en JavaScript sería más lento y en la práctica más pesado. |
| Calidad | «Visualmente sin pérdida» con tasa variable calculada (ver «Medido»), Opus 128 kb/s estéreo (64 kb/s en mono) o ~75 % de la tasa de un MP3/AAC si era menor | Volver a comprimir siempre pierde algo; la meta es que no se note. Comprimir sin ninguna pérdida casi no ahorra. |
| Dónde se convierte | **En el equipo de quien sube**, con los codificadores del navegador (WebCodecs, biblioteca Mediabunny, MPL-2.0) | El servidor (la Pi) no gasta CPU, y por la conexión de la casa solo sube lo ya convertido (ligero). Si el equipo no tiene codificador AV1, VP9; si no tiene WebCodecs, se sube el original dentro del límite y se marca «sin convertir». |
| Dónde se guarda | En el servidor, **aparte del documento** (archivo binario por medio, no dentro del JSON), por espacio, con permisos del espacio, cifrado en reposo como lo demás, y cifrado en el equipo si el espacio es de extremo a extremo (como los adjuntos del chat) | Así se abre desde cualquier lugar, el documento sigue ligero y la sincronía no carga videos. |
| Reproducir | El servidor entrega el archivo por partes (`Range`), con caché permanente en el navegador (el nombre es la huella del contenido) | Ver un video otra vez no lo vuelve a bajar. |
| PowerPoint | Al descargar el .pptx, **el navegador** convierte el video a MP4 H.264/AAC (por hardware, rápido) para que PowerPoint lo reproduzca. Los medios que venían en un .pptx salen con su archivo original | El servidor solo manda la versión ligera: el ancho de banda no cambia. |

## Medido (17-nov-2026, Chromium de Playwright, contenido sintético testsrc2)

- El AV1 que trae el navegador (libaom en «tiempo real») **no** es más eficiente que H.264 a la misma tasa: en 720p,
  1 Mb/s, 36.8 dB contra 37.9 de x264 «medium» (SVT-AV1 de ffmpeg sí: 40.0). Su modo de calidad constante
  (cuantizador) no baja de ~6.4 Mb/s en 720p aunque se pida la menor calidad: no sirve para aligerar.
- Por eso lo ligero sale de **no gastar de más**: tasa variable = 60 % de la de un H.264 (90 % HEVC/VP9), con techo de
  0.1 bits por píxel y cuadro (1080p30 ≈ 6.2 Mb/s) y piso de 0.025. Resultado: un video de teléfono 1080p a 17 Mb/s
  pasó de 13.7 a 4.3 MB (**69 % menos**, PSNR 42.5 dB, SSIM 0.992); un MP4 ya comprimido (8.3 Mb/s) de 6.2 a 3.4 MB
  (**45 % menos**, SSIM 0.990). Un MP3 de 192 kb/s a Opus: 31 % menos. Si no se ahorra al menos 3 %, se sube el original.
- Falta medir con videos reales de cámara y en el A55 (en casa), y la vuelta a MP4: el Chromium de pruebas no trae
  codificador H.264 (Chrome y Android sí).
- Si algún día el servidor tiene CPU de sobra, recodificar en él con SVT-AV1 daría ~20-30 % menos a la misma calidad.

## Límites (pensados para muchas cuentas en 2 TB; se cambian con propiedades del servidor)

- **Por cuenta: 5 GB de medios** (`nom.medios.cuota.mb`, por omisión 5120). 2 TB alcanzan para 400 cuentas llenas;
  la mayoría usa mucho menos. Se ve en la cuenta cuánto queda.
- **Por archivo, ya convertido:** video hasta **500 MB**, audio hasta **200 MB**. El archivo de origen puede pesar hasta
  4 GB (se convierte en el equipo antes de subir).
- **Video:** hasta 1080p (lo más grande se reduce sin deformar), los cuadros por segundo del original hasta 60,
  hasta 60 minutos. **Audio:** hasta 3 horas.
- Lo que no se usa en ningún documento (ni en sus versiones) durante 30 días se borra y libera la cuota.

## Fases

1. **(Hecho el 17-nov: `js/medios-convertir.js`, `js/vendor/mediabunny.min.js` 1.65.1.)** **Convertir en el navegador** (`js/medios-convertir.js` + `js/vendor/mediabunny`): de MP4, MOV, M4V, WebM, MKV,
   MP3, M4A, AAC, WAV, OGG, FLAC a WebM AV1/Opus con barra de avance, cancelable, con los límites de arriba y las
   alternativas (VP9; original). También la vuelta a MP4 H.264/AAC para PowerPoint.
2. **Almacén de medios en el servidor**: subir (en partes, se puede reanudar), entregar con `Range` y caché, permisos
   por espacio, cuota por cuenta, cifrado en reposo por bloques (para poder pedir partes), extremo a extremo en los
   espacios cifrados, borrar lo que nadie usa. También en el teléfono como servidor y sin Internet (lo subido espera
   en el equipo hasta que haya conexión).
3. **Insertar video y audio** en presentaciones (y en documentos de texto), reproducirlos en el editor y al presentar
   (lo de `medios-diapositiva.js`, que hoy toma los del .pptx original, pasa a usar también el almacén), con
   recorte, inicio automático o al clic, volumen, repetir y marcadores.
4. **Office**: al descargar el .pptx, la conversión a MP4; al abrir un .pptx con videos, ofrecer guardarlos en el
   almacén ya convertidos (más ligeros).
5. Pruebas: conversión en la computadora y en el A55, reproducción en el WebView, subir con la conexión de la casa,
   límites y permisos (`pruebas/seguridad.py`). **Las pruebas largas y en el teléfono esperan a que el usuario esté en
   casa.**

## Riesgos conocidos

- Equipos de Apple viejos (sin AV1 por hardware) no reproducen AV1 en Safari: ahí se usaría el original o VP9.
- Convertir un video largo en un teléfono lento tarda (se muestra el avance y se puede seguir trabajando).
