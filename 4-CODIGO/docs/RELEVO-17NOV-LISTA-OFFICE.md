# Relevo del 17 de noviembre de 2026: la lista grande de Office, NOM y lo que sigue

> **Histórico** (se conserva como registro de ese día; lo que dice del estado puede ya no ser cierto). Lo vigente
> —qué está confirmado, qué falta y en qué orden— está en `RELEVO-ESTADO-Y-PENDIENTES.md`.

Leer al empezar la siguiente sesión, junto con `PLAN-CONSTRUCCION.md` y `LOG.md` (entradas 71 a 76).

## 1. Estado del repositorio

- **Nueve commits locales sin empujar ni publicar** (de `6fadd42` a `e6e3406`; `git log origin/main..HEAD`).
  Nada de esto está en la Raspberry, en GitHub ni en el APK. Azure sigue apagada a propósito.
- El código todavía dice **0.15.4** (`Version.java`, `NOM.VERSION` en `nucleo.js`, `VERSION` en `sw.js`,
  `versionName`/`versionCode` en `android/app/build.gradle.kts`). Las novedades de la ayuda y `CAMBIOS.md` ya
  describen la **0.16.0**: al publicar, subir esos cuatro lugares a 0.16.0 (código de APK 31).
- Sin correr todavía: `pruebas/regresion.sh rapido`, ediciones/fidelidad completas y la prueba en el A55. El usuario
  pidió «todo menos las pruebas» en esta tanda; correrlas antes de publicar.
- Sin rastrear a propósito: `7-PRESENTACION/`, `ENTREGA/5IM7_WACK_Casos_de_Uso.*`, `ENTREGA/5IM7_WACK_Presentacion.pdf`
  (entregables del equipo; no se tocaron).
- `tmp-17-11-26/` (3 GB) es una copia de `/tmp` que guardó el usuario; está en `.git/info/exclude` (no se sube).
  Ahí quedó el material de trabajo de esta sesión (ver la sección 5).

## 2. Lo que se hizo (detalle en LOG 71 a 76)

| Commit | Qué |
|---|---|
| `fa3cc00` | Matrices dinámicas, LET y LAMBDA; Word con secciones de otra hoja, columnas, rellenos y grupos |
| `3d48460` | SmartArt sin dibujo (`js/smartart.js`); la vuelta en Office real (`vuelta_office.py`) |
| `b3432b4` | Word: lo que flota en su lugar con el texto alrededor; espacio entre párrafos y páginas como Word |
| `a9e388a` | Archivos de Office con contraseña (`js/cifrado-office.js`, MS-OFFCRYPTO, Agile y Standard) |
| `1b0d8ea` | Protección como en Office (restringir edición, final, contraseña para modificar, hoja y libro) |
| `1918156` | **Se llama NOM** + revisión externa (privacidad, Cloudflare, Azure en EE. UU., `security.txt`, huella del APK, accesibilidad, ayuda sin contradicciones, `CAMBIOS.md`, LEEME) |
| `03c64a3` | **Tablas de Word como Word** (ver abajo) |
| `e6e3406` | LOG 71 a 76 |

### Renombre a NOM (1918156)
- Todo lo visible dice NOM. Se conservan a propósito: el paquete `mx.wack.nom`, `nom-colab.jar`, el repositorio
  `TruTruMaster/NOM-COLAB`, el asset `NOM-COLAB.apk` (el botón de descarga apunta a ese nombre), el canario cifrado
  de `ArranqueCifrado.java` (`"nom":"NOM-COLAB"`) y `NOM_NUBE=https://nomcolab.dpdns.org` en la Pi (los pases
  firmados de los teléfonos dependen de esa identidad). Los buscadores tienen los nombres viejos como alternativos.
- La app usa **www.nom.qd.je** y trata `nomcolab.dpdns.org`, `nom.qd.je` y `www.nom.qd.je` como la misma nube
  (`api.js`: `NUBE_PRINCIPAL`, `mismoServidor`, migración de lo guardado). El enlace de invitación de www.nom.qd.je
  abre la app (`AndroidManifest.xml`). **Requiere recompilar el APK.**
- `nom.qd.je` (raíz, GitHub Pages `TruTruMaster/nom-qd-je`): http funciona, **https aún no** (GitHub no había
  emitido el certificado; se volvió a registrar el dominio por la API). Revisar
  `GET /repos/TruTruMaster/nom-qd-je/pages` y, con certificado, activar `https_enforced`. La página de redirección
  todavía dice «NOM-COLAB» en su título.
- Pendiente de la revisión: descripción, sitio web y temas del repositorio en GitHub (R3), camino de recuperación de
  cuenta sin correo ni administrador (E1), limpieza de tumbas del CRDT (E2), probar restaurar un respaldo (E6),
  .doc/.xls/.ppt/.odt que el selector deja elegir pero no abre (C8).

### Tablas de Word (03c64a3)
- `ooxml.js`: `leerTabla(c, prof)` (antes estaba dentro de `bloque`), `ricoCelda`, `bordesDirectosWord`,
  `firmaCelda`; campos nuevos del bloque tabla: `fh:N` (alto «e»xacto/«m»ínimo en px), `fp:N` (no partir),
  `cabRep`, `cmar`, `cli`/`clx`, `cesp`, `tw`/`tj`/`ti`, `va:`, `td:`, `bxd:`, `z:`, `h:` (contenido rico como en
  Word, con la huella del texto) y `o:` (`filaOriginal|firma`). Al descargar, `parcharTabla` saca la tabla original
  con solo las celdas que cambiaron.
- `modo-texto.js`: `htmlTabla`, `htmlRico`, `ricoDe`, `partirTabla` (filas de relleno `.wd-fila-corte`, encabezado
  repetido `.wd-cab-rep`, recorte con `clip-path`, `.wd-borde-corte`) llamada desde `calcularPaginas` y desde el
  paginador con secciones; `parteMinima` de una tabla = su primera fila. Doble clic en una celda rica la vuelve texto.
- Medido contra Word real: w05 8/8 páginas, w09 9/9; once de doce documentos con las páginas de Word. Sin tocar, el
  `document.xml` sale idéntico; editado, Word lo abre sin avisos y el validador de Microsoft no da errores nuevos.
- **Falta:** tabla flotante (`tblpPr`) con el texto alrededor (la hoja es `display:flex`: hay que tratarla como un
  flotante del motor de «lo que flota»).

## 3. Lo que sigue (decidido, en este orden)

1. **Motor de animaciones de PowerPoint** (empezado solo en análisis; nada escrito en el código).
   - PowerPoint guarda cada efecto como comportamientos en `p:timing` (ejemplos reales en
     la carpeta de trabajo de la sección 5, `anim/slide37.xml.timing.xml` y `slide38…`, sacados de
     `~/.cache/nom-pruebas/office-real/archivos/p08-animaciones.pptx`): `set` (visibilidad y estilos),
     `animEffect` (filtros fade, wipe(dir), barn, blinds, box, checkerboard, circle, diamond, dissolve, plus,
     randombar, strips, wedge, wheel(n)…), `anim` con fórmulas (`#ppt_x`, `ppt_w`, `1+#ppt_h/2`, `$`, `sin`, `pi`,
     `fmla` por `tav`, `tmFilter`), `animScale`, `animRot`, `animMotion` (ruta), `animClr`, `iterate` (por letra o
     palabra), disparadores (`interactiveSeq`), `cmd`/`audio`/`video`.
   - Plan: módulo nuevo `js/animaciones-office.js` con un motor por cuadros (requestAnimationFrame) que evalúe esos
     comportamientos sobre los bloques (y sobre párrafos, palabras y letras del texto). Importar `p:timing` a
     `anim:<página>` (cada efecto con su id de bloque, inicio clic/con/después, retraso, duración y `ppt`: sus
     comportamientos). El mapa `spid → bloque` sale de `ox` (índice del hijo del árbol de la diapositiva) y del `id`
     de su `cNvPr`. Reproducir en `presentar` (`lienzo.js`, `htmlDiapositiva`/`avanzar`) y en el proyector.
   - Hoy: NOM tiene 17 efectos propios (CSS `a-<tipo>-<efecto>`), las animaciones importadas **no se ven** (solo se
     conservan al exportar, DEF-45) y las animaciones creadas en NOM **no se exportan** a PowerPoint.
   - Después: crear con los presets de PowerPoint (entrada, énfasis con parámetros —aumentar/encoger, girar,
     transparencia, colores, negrita, subrayado—, salida, rutas) y escribir `p:timing` al exportar.
2. **Transiciones de PowerPoint y Morph.** p08 trae 36 (fade, push, wipe, split, reveal, randomBar, circle, pull,
   cover, fallOver, drape, curtains, wind, prestige, fracture, crush, peelOff, pageCurlDouble, airplane, origami,
   dissolve, checker, blinds, wheel, ripple, honeycomb, glitter, vortex, shred, switch, flip, gallery, prism, doors,
   morph…); hoy se importan 6 y NOM dibuja 4. Morph: emparejar objetos entre diapositivas (por nombre `!!` o por
   contenido) e interpolar lugar, tamaño, giro y color.
3. Herramientas de datos de Excel: minigráficos, segmentaciones, buscar objetivo, Solver, escenarios, tablas de
   datos, quitar duplicados, texto en columnas, relleno rápido, consolidar, previsión.
4. Imágenes EMF/WMF (hoy no se ven), citas y bibliografía de Word (campos CITATION/BIBLIOGRAPHY y `Sources.xml`
   unidos a `citas.js`), combinar correspondencia (MERGEFIELD), medios (video y audio) en diapositivas.
5. VBA: conservar `vbaProject.bin` (.xlsm/.docm/.pptm), ver y editar los módulos (MS-OVBA) y un intérprete aislado.
   Power Query y Power Pivot: conservar y mostrar consultas/modelo. Modelos 3D (glb…): al menos su imagen de respaldo.
   Efectos 3D en formas (scene3d/sp3d). Archivos .doc/.xls/.ppt y .odt/.ods/.odp.
6. **Decisión pendiente del usuario (fuentes y GPL-3.0):** los dibujos son OFL (bien), pero el interletraje y las
   ligaduras se copiaron de las fuentes de Microsoft (zona gris legal). Quitarlos cuesta de 1 a 3 puntos de renglones
   idénticos contra Word. Presentarle las dos opciones con esos números antes de publicar.

## 4. Cómo trabajar (lo que costó descubrir)

- Windows con Office: `user@100.65.29.5` (el alias `flast-desktop-win` falla por la llave del equipo).
  `python3 pruebas/office/office_real.py correr ID script.py args`; para abrir archivos en Office real, subirlos a una
  carpeta propia (`nomoffice\entrada\<nombre>`) y correr `abrir_office.py entrada\<nombre>`: en `entrada` hay archivos
  cifrados que traban a Office con su ventana de contraseña.
- Medir contra Office: `PLAYWRIGHT_BROWSERS_PATH=~/.cache/nom-pruebas/navegadores
  ~/.cache/nom-pruebas/venv/bin/python pruebas/office/comparar_real.py ARCHIVOS --salida DIR` con el servidor local
  en el puerto 8099 (`NOM_DATOS=/home/user/nom-datos-pruebas NOM_PUERTO=8099 java -jar target/nom-colab.jar`).
  Detenerlo con `kill $(pgrep -f "^java -jar target/nom-colab.jar")` (un `pkill -f` con ese texto se mata a sí mismo).
- La lista de renglones de Word no incluye las celdas de tabla: las tablas se juzgan por páginas e imagen.
- Validador de Microsoft: `python3 pruebas/office/validar_openxml.py ARCHIVO --corpus ~/.cache/nom-pruebas/office-real/archivos`.
- Después de cambiar la ayuda (`4-CODIGO/ayuda/contenido`), correr `python3 4-CODIGO/ayuda/generar.py` (también
  escribe `CAMBIOS.md`).
- Publicar (cuando toque y las pruebas pasen): jar a la Pi, push a `main` (despliega a Azure si toca `4-CODIGO/app`),
  APK con `assembleRelease` + `android/publicar-apk.sh` (ahora publica SHA-256 y huella del certificado).

## 5. Material de esta sesión en `tmp-17-11-26/`

La carpeta de trabajo (encontrarla con `find tmp-17-11-26 -maxdepth 4 -type d -name scratchpad`; es la de la sesión `1bbc013b…`):

- `revision_externa.txt`: la revisión externa completa que pegó el usuario.
- `tablas/`: `abrir_doc.py` (abre un archivo de Office en NOM y corre un JS de medición), `vuelta_tablas.py`
  (importar, editar y exportar w05 y comprobar), `w05-vuelta-*.docx`.
- `anim/`: el `p:timing` real de las diapositivas 37, 38 y 40 de p08.
- `tab0` a `tab3`, `todosw`: comparaciones contra Word (imágenes lado a lado y `reporte.json`) antes y después de
  las tablas.
- `renombrar.py`: el script del renombre (excluye LOG, entregables e historia).
