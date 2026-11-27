# Relevo del 26 de noviembre de 2026 (noche, segunda sesión)

**Leer este primero.** Sigue a `RELEVO-26NOV-NOCHE.md`. Publicada al empezar: **0.23.5** (interfaz 27, APK 47).
Versión de esta sesión: **0.23.6** (interfaz 27, APK 48) — ver §4 para qué parte ya está en cada lugar.

## 1. De qué trató esta sesión

Lo pendiente del relevo anterior: la decisión del dueño sobre quien «Solo puede ver» y los barridos que los ayudantes no
pudieron hacer por falta de cuota (comparación ruta por ruta del teléfono con el servidor, llaves de los espacios
cifrados y el HTML de los importadores de Office y de las vistas grandes). El detalle de seguridad vive solo en el plan
privado (§P); aquí va la cuenta neutra.

## 2. Hecho

- **Decisión del dueño (26-nov): sí se iguala.** Quien tiene «Solo puede ver» en un espacio recibe la misma copia que el
  enlace para ver (sin lo borrado, comentarios, notas del orador ni el original de Office), no ve versiones, copia lo
  que ve y en vivo recibe lo mismo que una pantalla de proyección (si algo borrado vuelve a verse, su página vuelve a
  pedir la copia). Cuenta el papel en el espacio, no la fecha de entrega. La página le esconde Historial y Comentarios.
  Prueba nueva: `DifusorLectoresTest`.
- **Espacios cifrados:** la llave de un espacio solo se usa si abre su nombre cifrado; dar la llave a otros es de quien
  edita y ya la tiene; cada quien puede quitar la suya; el reloj de una pestaña ya no regresa; importar un `.nom` con
  varios textos ya no repite identidades. Prueba nueva: `app/src/test/js/e2e-llaves.test.js` (falla con el código de
  antes en los 6 casos de arreglo).
- **Teléfono y servidor parejos:** el teléfono exige `application/json` y no aparta memoria por lo que dice que traerá;
  el servidor topa los códigos inexistentes del enlace para ver y limita a 256 KB lo que llega a cuentas y enlaces.
- **Ayudantes:** `nom-ds` salta solo al siguiente modelo gratis cuando uno se queda sin cuota (lista en
  `~/.config/nom-ds/modelos`; hoy sirve `hy4-flash-free`, que hay que exigirle el formato del informe). agy en un
  worktree fuera de `~/NOM` no puede leer `~/NOM/GEMINI.md` ni correr comandos: se le hace trabajar en `~/NOM` sobre
  archivos que nadie más toca, después de hacer commit de lo propio.

## 3. Para el dueño

- Decisiones de diseño de los espacios cifrados (detalle en el plan privado §P.2): la contraseña de entrar también abre
  la llave privada (el servidor la ve); no se cambian las llaves al quitar a alguien; enlaces de los textos cifrados en
  claro; sin huellas de llaves para comparar.
- Lo manual de antes: Browser Cache TTL en Cloudflare, `reposo.key` (copia), correo ARCO, DMARC, TLS 1.2.

- **Diapositivas, hojas e importadores de Office:** los valores del documento (medidas, colores, imágenes, nombres de
  efectos) se validan donde se arma la página; los archivos con partes enormes o anidadas sin fin se rechazan sin trabar
  la página. `formulaExcel` da lo mismo que antes en 156 casos y ya no es cuadrática.

## 4. Estado de la publicación

- **0.23.6 en la Pi** (varias subidas el 26-nov; la última con todo). En vivo contra www.nom.qd.je: `humo.py` 36/36,
  funciones 202/0, extremo a extremo 33/0, sin Internet 32/0 (primera subida), red 24/0, presentación, importar
  .pptx, diagramas y la prueba 85 sin errores de JavaScript; lectores 18/18 y llaves 13/13 (scripts privados).
- APK 48 firmado con el certificado de siempre (`2ef146f6…`). GitHub `main` y release v0.23.6: ver el commit del relevo.
- Falta: instalar en el A55; pruebas en el navegador de los casos armados de CSS (están cubiertos con pruebas en Node).

## 5. Cómo seguir

1. Lo que quedó bajo del plan privado §P.3 (H5–H7) y las decisiones de diseño de §P.2 (para el dueño).
2. Lo de paridad con Office sigue en `RELEVO-21NOV.md`.
