# Relevo del 26 de noviembre de 2026 (noche)

**Leer este primero.** Sigue a `RELEVO-26NOV-TARDE.md`. Publicada al empezar: **0.23.4** (interfaz 27, APK 46).
Publicada al cerrar: **0.23.5** (interfaz 27, APK 47) — ver §2 para qué parte ya está en cada lugar.

## 1. De qué trató esta sesión

Revisión de la página real (www.nom.qd.je) con un modelo de amenazas propio: para cada hipótesis se siguió el flujo
completo (página → servidor → almacenamiento → lo que reciben los demás) y se comprobó en el sitio real con cuentas
desechables, borradas al terminar. El detalle vive solo en los archivos privados de seguridad (plan §O); aquí va la
cuenta neutra.

## 2. Hecho

- **Los 66 archivos sin commit de la tarde** quedaron revisados según el riesgo (servidor, teléfono, scripts de la Pi,
  página) y publicados en la 0.23.5. Decisión tomada: en `api.js` y `acceso.js` queda `nomcolab.dpdns.org` (el dominio
  que sí es nuestro); el nombre que dejó el renombre del 25-nov no lo era. No se muestra en pantalla.
- **Revisión con hipótesis:** 23 hipótesis escritas antes de mirar el código. **5 hallazgos confirmados y cerrados**
  (tres comprobados en el sitio real antes y después del arreglo), **18 descartadas o acotadas con su evidencia** y una
  decisión sugerida para el dueño (abajo). Quedan sin barrer a fondo: el HTML de las vistas grandes y de los
  importadores de Office (contenido por la política de seguridad de la página, pero falta revisarlo archivo por
  archivo), la comparación ruta por ruta del teléfono contra el servidor y las llaves de los espacios cifrados.
- Cambios visibles (novedades de la 0.23.5): quien solo mira un documento (enlace para ver, proyección) recibe solo lo
  visible; «Investigar» no lee páginas de la red del propio servidor; en espacios cifrados no hay biblioteca compartida
  ni búsqueda en el servidor; una imagen pendiente ya no cuenta como intento fallido en la proyección; al cerrar sesión
  o cambiar de servidor se olvidan las llaves e imágenes de la sesión anterior.
- Pruebas nuevas de Java: `FuentesRedPropiaTest`, `SecuenciaSinOcultosTest` (`mvn -o -q test` sin fallas).
- Publicación: ver §4.

## 3. Para el dueño

- **Decisión sugerida:** hoy quien tiene «Solo puede ver» dentro de un espacio recibe el documento completo (también lo
  borrado) y el historial de versiones, igual que quien edita. En Google Docs quien solo ve no ve el historial. ¿Se
  iguala? (si sí: los lectores reciben la misma copia limpia que el enlace para ver y no ven versiones).
- Lo manual que sigue de antes: Browser Cache TTL en Cloudflare, `reposo.key` (copia), correo ARCO, DMARC, TLS 1.2.

## 4. Estado de la publicación

- **0.23.5 en la Pi** (dos subidas el 26-nov; la segunda con todo), `humo.py` 36/36 contra www.nom.qd.je.
- Suites en vivo: funciones 202/0, extremo a extremo 33/0 (dos veces), sin Internet 32/0, ayuda 24/0,
  presentación 10/0, red 24/0, biblioteca y citas sin errores. `prueba_mudanza` no se corrió: necesita un servidor
  local (prohibido por la orden del 25-nov).
- El teléfono que comparte, probado en el emulador con la versión de depuración (mismo código): página, API, varias
  peticiones por conexión, cuerpos, peticiones mal formadas, entrada del dueño con su llave y edición de un documento.
- **GitHub** `main` al día y **release v0.23.5** con el APK 47 (mismo certificado de siempre). Buscadores avisados.
- Falta: instalar en el A55 (no estaba conectado).

## 5. Cómo seguir

1. Si el dueño contesta lo de §3, hacerlo (es poco código: los lectores reciben `paraMirar` y no ven versiones).
2. Los ayudantes se quedaron sin cuota a las 15:15 (DeepSeek: límite gratuito; agy: Gemini y respaldo). Cuando
   vuelvan, mandarles los barridos que quedaron pendientes (las tareas están escritas en el plan privado §O y en
   `trabajo/auditoria-26nov-noche/`): importadores de Office y vistas grandes de la página, y la comparación ruta por
   ruta del teléfono contra el servidor.
3. Lo de paridad con Office sigue en `RELEVO-21NOV.md`.
