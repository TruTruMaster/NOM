# Relevo del 17 de noviembre de 2026 (tarde): animaciones y transiciones de PowerPoint

Leer al empezar la siguiente sesión, junto con `RELEVO-17NOV-LISTA-OFFICE.md` (lo de antes), `PLAN-CONSTRUCCION.md` y
`LOG.md`. Este relevo se actualiza mientras se trabaja: lo último está al final de cada sección.

## 0. Pedido y decisiones del usuario (17-nov)

- Pedido (textual, resumido): seguir el proyecto con todas las funciones de Office, las más complejas y las que la gente
  paga y que LibreOffice, Google Docs u OpenOffice no tienen, «fully and proven», comprobado en la Windows con Office real.
- **Decisión de fuentes: «we stay on the gray zone, with all the features no matter what»**: se quedan el interletraje y
  las ligaduras copiados de las fuentes de Microsoft (la decisión pendiente del punto 6 del relevo anterior). No quitarlos.
- Al acercarse el límite de uso: guardar lo de `/tmp` en un lugar que no se borre y dejar el relevo. **Hecho así:** las
  herramientas van al repositorio (`4-CODIGO/pruebas/office/analisis/`) y los datos chicos a `~/NOM/trabajo/` (está en
  `.git/info/exclude`). Lo útil de la copia de 3 GB `tmp-17-11-26/` quedó en `~/NOM/trabajo/sesion-16nov/` (revisión
  externa, scripts de tablas, parches, `renombrar.py`, reportes): **esa copia de 3 GB ya se puede borrar** (pedirle permiso
  al usuario antes de borrarla; no se borró).

## 1. Estado del repositorio

- Commits locales de esta sesión (sin empujar ni publicar; se suman a los nueve del relevo anterior): `198e7c8` (motor de
  animaciones), `f4bd7c1` (fidelidad, letras, transiciones) y los que sigan. Nada está en la Pi, en GitHub ni en el APK.
- La versión en el código sigue en 0.15.4; las novedades de la ayuda dicen 0.16.0 (ver el relevo anterior). Antes de
  publicar: regresión completa, subir versión en los cuatro lugares y escribir las novedades de animaciones y transiciones
  en la ayuda (`4-CODIGO/ayuda/contenido`, luego `python3 4-CODIGO/ayuda/generar.py`).
- El servidor de pruebas: `cd 4-CODIGO/app && SPRING_WEB_RESOURCES_STATIC_LOCATIONS=file:/home/user/NOM/4-CODIGO/app/src/main/resources/static/ NOM_DATOS=/home/user/nom-datos-pruebas NOM_PUERTO=8099 nohup java -jar target/nom-colab.jar &`
  (con esa variable el JS y el CSS se sirven en vivo, sin recompilar). Detenerlo: `kill $(pgrep -f "^java -jar target/nom-colab.jar")`.

## 2. Lo que se hizo

### Motor de animaciones (`js/animaciones-office.js`)
- Interpreta el árbol de tiempos de PowerPoint (ECMA-376 §19.5) tal cual: `set`, `anim` (con fórmulas `#ppt_x`, `ppt_w`,
  `$`, `sin`, `pi`…, `tavLst` con `fmla`, `from/to/by`, aditivo), `animEffect` (filtros), `animScale`, `animRot`,
  `animMotion` (caminos M/L/C/Z/E y relativos, `rAng`), `animClr` (rgb/hsl, by), `cmd`; repetir, autorreversa, velocidad,
  `tmFilter`, `iterate` por letra y palabra (con `tmPct`/`tmAbs`, al revés), disparadores (`interactiveSeq`), grupos por clic
  como los arma PowerPoint, composición tipo SMIL (lo que empezó después manda; lo aditivo se suma).
- Importa `p:timing` del .pptx a `anim:<página>` (cada efecto con su árbol en `ppt.c`, destinos con `bid`), `animMedios:`,
  `animBld:`; los colores del tema se resuelven al importar (`_c`) y el esquema queda en `_doc.esquemaOffice`.
- Exporta: `timingXml()` escribe el `p:timing` desde la lista (nuevo en `pptxArmar`; en la ruta con origen se reescribe solo
  si la lista cambió, con el número de figura que quedó en el archivo). Los 17 efectos de antes de NOM se reproducen y se
  escriben con su efecto de PowerPoint (`LEGADO`).
- Biblioteca: `js/animaciones-office-datos.js` (262 efectos tal como los escribe PowerPoint 2024: 118 de entrada con
  direcciones, 51 de salida, 29 de énfasis, 64 trayectorias; nombres en español y categorías). Se arma con
  `4-CODIGO/herramientas/animaciones/armar_presets.py` desde los archivos `a01`/`a02` que hace PowerPoint.

### Lo que se midió de PowerPoint (no volver a adivinar)
- **Filtros (barrido, persianas, cuadros, cuadro, círculo, rombo, más, dividir, cuña, rueda, barras diagonales):** cada
  punto se descubre según un campo D(x,y) ∈ [0,1] con borde LINEAL: alfa = recortar((s·(1+G) − D)/G), **G = 0.145** (el
  borde del barrido coincide al píxel). `wipe(down)` descubre desde ABAJO; `wipe(up)` desde arriba. La salida («out») es lo
  contrario de la entrada en barrido, cuadro, círculo, rombo, más, dividir, cuña y rueda; es la entrada al revés en
  persianas, cuadros bicolores, barras diagonales (y aleatorios). Cuadros bicolores: 6 filas, periodo = ancho/6, filas
  alternas corridas medio periodo. Persianas: 6. Diagonales: celdas de 10 pt. Disolver y barras: al azar a 1 px.
- **Inicio/fin suaves (accel/decel)**: NO es la aceleración constante de SMIL. Velocidad máxima R = 1/(1 − a/2 − d/2) igual,
  pero la rampa de velocidad es suave (3u² − 2u³): recorrido en la subida = R·a·(u³ − u⁴/2). Medido con 16 combinaciones
  (`a04-aceleracion`): error medio 0.12 % del recorrido (SMIL: hasta 4 %). **Se aplica a cada comportamiento por separado**
  (no al efecto entero): medido en «Silbido» al medio píxel.
- **Rebote al final** (`p14:bounceEnd` = b): llega en 1 − b a velocidad pareja y rebota como resorte amortiguado con el
  tiempo escalado por √b (K = 6.25, W = 15.75): error medio 0.17 %.
- **`ppt_x` sin #** en una fórmula = el valor al EMPEZAR ese comportamiento (lo que dejaron, en ese instante, los que
  empezaron antes, aunque terminen justo ahí). Arregló Voltear, Rebote, Flotar, Giro central, Cremallera, Bumerán.
- Comportamientos con `rctx="IE"` (para el reproductor de Internet Explorer): PowerPoint no los usa (Transparencia).
- Por letra: **los espacios no cuentan**; la caja de PowerPoint de cada letra (para ppt_x/y/w/h) **no depende de la letra**:
  centro a 0.6 del alto del renglón desde arriba, alto = 2.02 × tamaño + 5.2 pt (Arial, Calibri, Times; 18 a 72 pt;
  interlineado 1 a 2; `a05-letras`). Si una letra cambia de tamaño, PowerPoint vuelve a acomodar el texto.
- Una figura sin relleno anima su color de relleno desde el blanco del tema (bg1).
- Transiciones: Desvanecer (cruzado) coincide 97 %; Empujar va con curva suave; «Desvanecer a negro» en el video de
  PowerPoint cambia de golpe en ~0.6 s (rareza del video: no copiar a ciegas).
- **Video de PowerPoint (`CreateVideo`) es de cuadros variables**: hay huecos de ~0.96 s. Decodificar siempre con
  `ffmpeg -fps_mode passthrough` y usar los tiempos de `ffprobe` (si no, los tiempos se corren).

### Resultados (NOM contra el video de PowerPoint, cuadro por cuadro, 640×360)
- `a01-animaciones` (195 efectos: todos los de entrada, salida, énfasis y trayectorias): **97.0 % medio, 159 ≥ 98 %**.
  Lo flojo que queda es texto (88–95 %), limitado por el dibujo estático del texto (ver pendientes).
- `a02-opciones` (por letra, palabra, párrafo, repetir, rebobinar, suaves, rebote, atenuar, ocultar, direcciones):
  letras y palabras 94–97 %. Flojos: «Destello en negrita por letra» (84 %), «Volar por párrafo al revés» (85 %).
- La vuelta (`animaciones_vuelta.py`): importar, cambiar duración y retraso de todo, exportar y volver a importar: los
  efectos y transiciones de a01, a02, a03 y p08 salen iguales (±1 ms de redondeo); el validador de Microsoft: **0 errores**
  en los 8 archivos. **Comprobado en PowerPoint de verdad:** `v01-nom.pptx` (el catálogo con los 195 efectos reescritos por
  NOM, cada duración ×1.5 y +100 ms) abre sin avisos ni reparar (195 diapositivas, 2535 formas) y su video de PowerPoint
  contra NOM reproduciendo el mismo archivo: **98.3 % medio, 159 de 195 ≥ 98 %** (lo más bajo, efectos de texto 88–93 %).

### Transiciones (`js/transiciones-office.js`)
- Las 149 variantes de Office 2024 se importan (`transOffice:<página>`: tipo, atributos, duración `p14:dur`, `advTm`,
  `advClick`, sonido) y se reproducen al presentar (2D con el mismo modelo de borde; 3D con CSS: cubo, caja, girar,
  órbita, puertas, ventana, voltear, cambiar, galería, cinta, noria, desplazamiento, atravesar; en piezas: vórtice, onda,
  panal, brillo, triturar, deformar, fractura, aplastar; Transformar empareja objetos por nombre «!!», tipo y texto).
  El avance automático («después de n s») se respeta. Se exportan como PowerPoint (`xml()`, con `mc:AlternateContent` y
  respaldo para las de 2010 en adelante).
- **Falta calibrarlas contra el video** (`transiciones_real.py` ya mide con los inicios exactos): hoy solo Desvanecer
  está medida (97 %). Lo de 3D y en piezas es aproximado.

## 3. Herramientas (todas en el repositorio)

- Windows (`user@100.65.29.5`, Tailscale; Office ProPlus 2024 16.0.17932): `pruebas/office/windows/animaciones_office.py`
  `constantes | catalogo | opciones | transiciones | aceleracion | letras | video NOMBRE`. Se corre con
  `python3 pruebas/office/office_real.py enviar` y `office_real.py correr ID animaciones_office.py catalogo`, y
  `office_real.py traer "a0*"`. Los archivos y videos quedan en `~/.cache/nom-pruebas/office-real/archivos/`
  (a01-animaciones, a02-opciones, a03-transiciones, a04-aceleracion, a05-letras: .pptx, .json, .mp4).
  Cada diapositiva lleva una marca binaria de 12 cuadritos (número + comprobación) para saber en qué cuadro empieza.
- Linux: `pruebas/office/animaciones_real.py ARCHIVO [--diapos 1-40] [--cuadros 8] [--desfase ms]` (efectos),
  `transiciones_real.py [--diapos]`, `animaciones_vuelta.py` (la vuelta), `validar_openxml.py CARPETA`. Con
  `PLAYWRIGHT_BROWSERS_PATH=~/.cache/nom-pruebas/navegadores ~/.cache/nom-pruebas/venv/bin/python` y el servidor en 8099.
  Reportes y lado a lado en `~/.cache/nom-pruebas/animaciones/<archivo>/`.
- Análisis (`pruebas/office/analisis/`): `mosaico.py` (cuadros de PowerPoint en mosaico), `perfil.py` (ancho del borde),
  `trayecto.py`/`ajuste*.py` (curvas de movimiento), `letras.py`/`cajas_letras.py` (caja de cada letra),
  `depurar_anim.py` (abre un .pptx en NOM y corre un JS en el instante t), `resumen_efectos.py`/`arbol.py` (XML de un efecto).
  Sus datos de esta sesión están en `~/NOM/trabajo/animaciones-17nov/`.

## 4. Lo que sigue (en este orden)

1. (Hecho: la vuelta en PowerPoint real, ver arriba. Ojo: por SSH a Windows no escribir `2>nul` en un comando de
   bash: se vuelve `/dev/null` y cmd se cuelga; usar `office_real.ssh(...)` desde Python.)
2. **Crear y editar animaciones como en PowerPoint** (hoy el menú de NOM solo ofrece sus 17 efectos): galería con los 262
   (Entrada/Énfasis/Salida/Trayectorias, categorías Básico/Sutil/Moderado/Llamativo), opciones de efecto (dirección,
   secuencia: como un objeto, todo a la vez, por párrafo; texto por letra o palabra con % de retraso), intervalos (inicio,
   duración, retraso, repetir, rebobinar, desencadenador al hacer clic en una figura), inicio y fin suaves, rebote,
   después de la animación (atenuar a un color, ocultar), panel de animación con línea de tiempo, copiar animación,
   trayectorias propias dibujadas. Todo con su árbol de la biblioteca (`presetDe`) y modificándolo (iterate, repeatCount…).
   En la ruedita y la paleta de comandos también (reglas del plan).
3. Calibrar transiciones con `transiciones_real.py` (empezar por Empujar, Barrido, Dividir, Revelar, Barras, Forma, Cubrir,
   Descubrir, Destello, Cubo, Galería, Caída, Cortinas, Curvatura) y el panel de transiciones (galería, opciones, duración,
   aplicar a todas, avance automático). **Transformar (morph)**: emparejar por forma y texto como PowerPoint (por objeto,
   palabra y letra) y medirlo (p08 diapositiva 36 y a03 152–154).
4. Medios en diapositivas (video y audio con `cmd` playFrom/togglePause, marcadores), compilación de gráficas y SmartArt
   por serie/categoría/nivel, botones de acción y vínculos a diapositivas, sonidos de efectos y transiciones.
5. Texto estático de PowerPoint ~2 px más arriba que en PowerPoint (medido en un cuadro de texto Arial 54 negrita: tinta
   en 284 contra 288 px a 1280 de ancho): revisar la posición del primer renglón (esto sube todos los efectos de texto).
6. Luego la lista del relevo anterior: herramientas de datos de Excel, EMF/WMF, citas y bibliografía, combinar
   correspondencia, VBA, Power Query/Pivot, 3D, formatos viejos y .od*; pruebas completas y publicar.
