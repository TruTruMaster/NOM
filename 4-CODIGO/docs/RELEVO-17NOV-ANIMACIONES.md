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
2. **(Hecho el 17-nov, noche; sin probar en el teléfono ni regresión: el usuario estaba fuera de casa.)** Crear y editar
   animaciones como en PowerPoint, en `js/animaciones-panel.js` (nuevo) con el contexto `ctxAnim` de `lienzo.js`:
   - Galería con los 262 efectos (Entrada/Énfasis/Salida/Trayectorias; categorías de PowerPoint; las trayectorias por
     Básico 1-18, Especial 19-34, Líneas y curvas 35-64), buscador en español o con el nombre en inglés. En el menú
     «Animación» hay seis rápidos a un toque.
   - Opciones del efecto (`opcionesDe`/`conOpciones` en `animaciones-office.js`): dirección (`variantes`, `cambiarEfecto`
     conserva las opciones), cantidad de Girar, tamaño de Aumentar o encoger (en `to`, como PowerPoint), color de los
     énfasis de color, inicio y fin suaves, rebote al final, autorreversa, después de la animación (atenuar a un color,
     ocultar al terminar, ocultar con el siguiente clic), texto por letra o palabra con % y al revés, por párrafos (un
     efecto por párrafo con `txEl/pRg`, `grupoP` y `par` en la lista; `build="p"` y `rev="1"` en `bldLst`). Intervalos:
     inicio, retraso, duración (deshabilitada en los instantáneos, como Aparecer), repetir (n, hasta el siguiente clic,
     hasta el final), rebobinar, desencadenador (clic en una figura). **Solo se reescribe lo que la persona cambió**
     (Balancear y Onda giran distinto en cada tramo; los colores del tema y un `tmAbs` se quedan como vinieron).
   - Panel de animación: vista previa con el mismo motor, reproducir todo o desde la elegida, una fila por efecto (los
     de por párrafos juntos), barras de tiempo que se arrastran (retraso) y su borde (duración), subir/bajar, opciones,
     quitar, desencadenadores agrupados.
   - Copiar y pegar animación (`paraBloque`), trayectoria dibujada a mano (`trayectoriaPropia`: presetClass="path"
     presetID="0", camino relativo `M 0 0 L … E`, `Z` si se cierra), en la ruedita (Más › Otras › Animación) y en la
     paleta (panel, agregar, trayectoria, copiar, pegar).
   - Reproductor: `subTnLst` (atenuar `ppt_c` al siguiente clic, ocultar al terminar o al siguiente clic), «hasta el
     siguiente clic» se detiene al hacer clic, grupos marcados con su secuencia.
   - Importar/exportar: el `p:timing` dentro de `mc:AlternateContent` (el que escribe PowerPoint con rebote) se lee (Choice) y
     se escribe igual (Choice con p14 y Fallback sin p14:*); en la ruta con origen también se encuentra, se poda y, si
     queda vacío, se quita el AlternateContent entero. `<p:tn val>` del «ocultar al terminar» se guarda como `@ef` y se
     escribe con el número nuevo del efecto. Los efectos por párrafos importados se marcan como grupo.
   - Prueba de humo hecha (a02-opciones en el navegador: importar, galería, opciones, vista previa, panel, trayectoria,
     `timingXml`): sin errores. **Falta** (cuando haya tiempo y PowerPoint): abrir en PowerPoint real un archivo con
     efectos armados en el panel (repetir, rebote, atenuar, ocultar, por párrafos, trayectoria propia) y pasar el validador
     de Microsoft; probar en el teléfono (barras con el dedo, trayectoria con el dedo, ruedita).
   - También hecho: «Animar la forma adjunta» al ir por párrafos (primer efecto con `<p:spTgt><p:bg/></p:spTgt>`, `par: -1`,
     `bldP build="p" animBg="1" autoUpdateAnimBg="0"`; se reconoce al importar), copiar un efecto por párrafos lo vuelve a
     partir en el bloque nuevo, y las trayectorias del bloque elegido se dibujan en el editor (punteadas, inicio verde y
     fin rojo, `.ao-tray`). Un efecto que anima un solo párrafo conserva su párrafo al editar sus opciones.
   - Pendiente de esto: sonidos de efectos; mover o editar los puntos de una trayectoria en el editor (hoy se dibuja de
     nuevo). Datos que faltan en la biblioteca (necesitan
     PowerPoint): las salidas solo traen su dirección de siempre y faltan direcciones de entrada (Desplazar 3 y 12,
     Barras diagonales 3, Dividir vertical hacia dentro, Rueda por radios, Arrastrar 4, Rotar vertical, Zoom desde el
     centro de la diapositiva); «Desplazar hacia fuera» se puede derivar de la entrada (el tav 0 con # de la entrada es
     el tav 1 sin # de la salida).
3. **Panel de transiciones hecho (17-nov, noche)**: `NOM.animPanel.transiciones(ctxTrans, p)` con el catálogo
   `NOM.transOffice.CATALOGO` (151 variantes tal como las escribió PowerPoint en a03, en Sutil / Llamativo / Contenido
   dinámico; `buscar` y `nombreCompleto` reconocen una transición importada), opciones del efecto, duración, avanzar al
   hacer clic y/o después de n s, aplicar a todas, vista previa desde la diapositiva anterior; en el botón Transición,
   la ruedita de la miniatura y la paleta. Prueba de humo en el navegador con a03: bien. Las duraciones con que
   empieza cada transición son las de PowerPoint de memoria, sin medir. **Sigue de este punto:** calibrar transiciones con `transiciones_real.py` (empezar por Empujar, Barrido, Dividir, Revelar, Barras, Forma, Cubrir,
   Descubrir, Destello, Cubo, Galería, Caída, Cortinas, Curvatura) y el panel de transiciones (galería, opciones, duración,
   aplicar a todas, avance automático). **Transformar (morph)**: emparejar por forma y texto como PowerPoint (por objeto,
   palabra y letra) y medirlo (p08 diapositiva 36 y a03 152–154).
4. **Vínculos y botones de acción hechos (17-nov, noche)**: `a:hlinkClick` del cNvPr de cada figura → `bloque.accion`
   (`sig`, `ant`, `primera`, `ultima`, `volver`, `fin`, `diapo` con su página, `url` http/https/mailto); al presentar,
   tocar la figura lo hace; se escribe como PowerPoint en las dos rutas (desde cero y con el original: `conVinculo`
   cambia el vínculo de la figura reescrita, sin quitar los que NOM no entiende, como abrir un archivo o una macro).
   Cuadro «Acción al hacer clic» en el menú Animación, la ruedita (Animación › Vínculo) y la paleta. Probado ida y vuelta
   (importar lo exportado) con los tres tipos. Arreglo de paso: en la ruta con origen, las relaciones externas nuevas
   salían sin `TargetMode="External"`. Falta: vínculos dentro del texto (a:rPr/hlinkClick), «al pasar el mouse»
   (hlinkMouseOver), abrir otro archivo o programa, y sonidos de acción. **Compilación de gráficas y SmartArt hecha (17-nov,
   noche):** SmartArt: cada figura importada guarda su `modelId` en `bloque.dgm` y el motor anima solo las figuras de
   `<p:graphicEl><a:dgm id/>` (probado con lo-tdf104792: tras el primer clic se ve el primer nodo y el segundo no).
   Gráficas: `graficas-office.js` marca cada barra, línea, punto, rebanada, rótulo, barra de error y tendencia con
   `data-s`/`data-c`; el motor anima el fondo (gridLegend), una serie, una categoría o un punto (`a:chart`); el cuadro
   de opciones tiene la pestaña «Gráfica» (agrupar por serie, categoría, elemento de la serie o de la categoría y
   «empezar dibujando el fondo») con `porGrafica`; se exporta con `bldGraphic/bldChart` y se reconoce al importar
   (probado ida y vuelta en las dos rutas). Falta: crear «uno por uno» de SmartArt desde NOM (al exportar sin el
   original el SmartArt sale como figuras sueltas), gráficas 3D y las nuevas de Excel sin marcar, y la máscara de un
   barrido por serie va por cada elemento, no por toda la serie. **Sonidos hechos (17-nov, noche):** el sonido de un
   efecto (`p:audio` con `sndTgt`, en `childTnLst` o en `subTnLst` con `evt="begin|end"`) y el de una transición
   (`snd`, en bucle si `loop`, y «detener el sonido anterior» con `sndFin`) suenan al presentar, sacados del original
   (`urlSonido` en lienzo.js, por la relación de la diapositiva); se callan al salir. Probado con ms-rendimiento-tipico
   (efecto) y lo-tdf148478 (transición). Al exportar sin el original se quitan (la relación no existe y PowerPoint
   pediría reparar); con el original se conservan. Falta: elegir un sonido para un efecto o transición desde NOM (los
   sonidos de fábrica de PowerPoint no vienen en el archivo) y el sonido de una acción (`a:snd` del vínculo).
   - **Video y audio de PowerPoint al presentar (hecho el 17-nov, tarde, en otra sesión; sin pruebas largas: el usuario
     estaba fuera de casa).** Al importar, la imagen de un `p:pic` con `a:videoFile`/`a:audioFile`/`p14:media` queda con
     `bloque.medio` = { tipo, ruta (la parte del .pptx), recorte {ini, fin} de `p14:trim`, fundido {ent, sal} de
     `p14:fade`, marcadores de `p14:bmkLst` } (`medioDePic` en `ooxml.js`). **El medio no se copia al documento:** al
     presentar, `js/medios-diapositiva.js` lo saca del original que ya guarda el documento (`NOM.origen.leer(doc).zip`) como
     URL de blob. Los ajustes salen del nodo `p:video`/`p:audio` de `animMedios:<p>` (volumen, silencio, repetir hasta
     detener = `repeatCount="indefinite"`, ocultar mientras no se reproduce = `showWhenStopped="0"`, seguir en las
     siguientes = `numSld`, pantalla completa = `fullScrn`) y los comandos llegan del motor por `ctx.alMedio`:
     `playFrom(s)`, `play`, `togglePause`, `pause`, `resume`, `stop`, `seek(s)`. Clic en un medio sin desencadenador:
     reproducir o pausa. Barra al pasar el puntero (reproducir/pausa, posición con marcadores, tiempo, sonido) y teclas de
     PowerPoint (Alt+P, Alt+Q, Alt+↑/↓, Alt+Inicio/Fin). En el editor, la marca «Video»/«Audio» sobre el bloque. Al
     exportar con el original, la figura reescrita conserva su `p:nvPicPr` (DEF-45), así que sigue siendo video.
   - **Arreglo en el motor (`animaciones-office.js`):** el comando de medio usaba `bh.ordenUnico`, que no existía: todos
     compartían la clave «undefined» y después del primero (por ejemplo, reproducir) ya no llegaba ninguno (pausa, volver
     a reproducir). Ahora la clave es comportamiento + momento en que empieza; al saltar a un paso (`irAPaso`, `cuadro`) se
     anotan como dados sin reproducir nada (regresar a una diapositiva ya no arranca sus videos) y `fotoDe` no los cuenta.
   - Prueba de humo (Chromium sin pantalla, tres archivos: `lo-video-loop-mute`, `lo-tdf120573`, `p08` diap. 40):
     reproducir solo y en silencio al 80 %, clic = pausa/sigue, audio al hacer clic desde 0, Alt+Q, nada queda al salir;
     sin errores. Nota: `nucleo.js` quita las hojas de estilo sin `data-nom` (la de los medios ya lo lleva).
   - **Falta de medios:** desencadenadores por marcador (`onMediaBookmark`: no hay archivo de PowerPoint con uno para ver
     cómo lo escribe; los marcadores ya se ven en la barra), confirmar en PowerPoint que `p14:trim end` es lo que se quita
     al final (así se tomó) y que `p14:fade` va en ms, insertar un video o audio desde NOM (necesita que el usuario decida
     dónde se guarda: dentro del documento pesa mucho), video en la pantalla del proyector, medios en línea (YouTube) y
     al exportar sin original (hoy sale solo la portada). Probar en el teléfono y en PowerPoint real.
5. Texto estático de PowerPoint ~2 px más arriba que en PowerPoint (medido en un cuadro de texto Arial 54 negrita: tinta
   en 284 contra 288 px a 1280 de ancho): revisar la posición del primer renglón (esto sube todos los efectos de texto).
6. Luego la lista del relevo anterior: herramientas de datos de Excel, EMF/WMF, citas y bibliografía, combinar
   correspondencia, VBA, Power Query/Pivot, 3D, formatos viejos y .od*; pruebas completas y publicar.
