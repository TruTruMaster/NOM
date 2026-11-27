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

## 2. Matriz de funciones (A)

(en construcción)

## 3. Resultados de las pruebas en vivo (B)

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
normales). 3G lenta: pendiente.

## 5. Para el dueño

- Prueba con personas: `ENTREGA/Prueba_con_personas.docx` (tareas de la guía §5 y SUS).
- Lo manual de antes: Browser Cache TTL en Cloudflare, `reposo.key` (copia), correo ARCO, DMARC, TLS 1.2.
- Decisiones de diseño de los espacios cifrados (plan privado §P.2): se recomiendan **después** de la expo.
