# NOM-COLAB — Diseño de la interfaz

Equipo 6 "WACK" · CECyT 9 · 5IM7 · octubre 2026 (apariencia revisada en noviembre de 2026, versión 0.14.0)

Este documento explica **por qué** la interfaz se ve y se comporta como lo hace.
Cada decisión sale de un principio de psicología cognitiva o de diseño de
interacción, y cada una dice en qué archivo vive. La investigación detrás de la apariencia actual
(qué hace que una interfaz parezca de plantilla, primera impresión, cómo se encuentran las cosas)
está en el documento «UI and Psychology» del equipo; aquí queda lo que se decidió.

## 1. Estilo: plano, neutro y con un solo color de marca

**Qué es.** Superficies planas con un borde de 1 px, grises neutros y un solo color de marca: el
azul noche del ícono. El color se guarda para lo que dice algo (el tipo de documento y los
estados). Esquinas redondeadas pequeñas, en una escala fija.

**Por qué se dejó la «arcilla» con «vidrio» (noviembre de 2026).** La versión anterior juntaba
catorce de las veinte señales por las que hoy se reconoce una interfaz armada con plantilla: acento
índigo con degradado y resplandor, desenfoque en todo lo que flota, radios de 16 a 28 px, luz bajo
el cursor, grano en el fondo, todo animado en cascada. Cada una tenía su razón; juntas daban una
primera impresión genérica, y esa impresión se forma en unos 50 ms (Lindgaard et al., 2006) y pesa
en la credibilidad (Fogg et al., 2003). Lo simple y familiar es lo mejor calificado (Tuch et al.,
2012), y los sistemas de diseño de las empresas grandes (Material 3, Fluent 2, Carbon) coinciden
en lo mismo: base neutra, un color de marca, pocos radios y pocas sombras definidos como variables.

**Decisión.**

| Regla | Dónde |
|---|---|
| **Un color de marca**, el azul noche del ícono (`#2B2952`, 13.6:1 con blanco), para **la** acción principal de cada pantalla. En tema oscuro la acción principal es clara (`--sobre-acento` oscuro). | `css/nom.css` (`--acento`, `--sobre-acento`) |
| El naranja de la llama (`--llama`) solo para avance y marca (barra de carga, anillo de primeros pasos): no alcanza contraste para texto. | `--llama` |
| Lo elegido **dentro de un documento** (celdas, marcos, tiradores) va en un azul propio, para no confundirlo con la acción principal. | `--seleccion`, `css/editor.css` |
| Toda superficie lleva **borde de 1 px**; la sombra solo dice qué está encima de qué (tres niveles). Los campos tienen un borde de 3.3:1: su límite no depende de una sombra. | `.caja`, `--sombra-1`, `--sombra-2`, `--sombra-flotante`, `--borde-campo` |
| Lo que flota (menús, ventanas, avisos, ruedita) es **opaco**: sin desenfoque. | `--sup-flota`, `--sup-inversa` |
| **Radios** 4, 6, 8, 10 y 12 px; nada más. | `--r-xs` … `--r-xl` |
| **Letra** Nunito (la de la marca) con tres pesos: 400, 600 y 700 (800 solo en el nombre de la app). Tres tamaños de título: 26, 18 y 15 px. | `css/nom.css` (base) |
| Sin degradados, resplandores, texturas ni luces que sigan al cursor. | — |
| Contraste medido contra WCAG 2.x AA: tinta 15.5:1, texto secundario 6:1 (5.5:1 sobre lo hundido), el menor de los textos de color 5.3:1. | `--tinta`, `--tinta-suave` |
| Modo **calidad baja**: quita las sombras (hardware mínimo de los requerimientos). | `:root[data-calidad="baja"]`, CU-27 |
| El **lienzo de trabajo** es papel blanco sobre gris neutro y no compite con el contenido. | `css/editor.css` (`.papel`, `.wd-pagina`) |

## 2. Captar la atención rápido (sin robarla)

| Principio | Qué dice | Cómo se aplica |
|---|---|---|
| **Atributos preatentivos** | El color, la forma y la posición se procesan antes de la atención consciente (≈ 200–250 ms). | Cada modo tiene un color fijo en todo el sistema: Texto azul, Hoja verde, Presentación naranja, Diagrama violeta. Se reconoce el tipo de documento sin leer. `--m-texto`, `--m-hoja`… |
| **Efecto Von Restorff** | Lo que es distinto a su entorno se recuerda y se encuentra primero ([Laws of UX](https://lawsofux.com/)). | Un solo color de marca (azul noche) y se usa para **una** acción principal por pantalla: "Entrar", "Nuevo espacio", "Nuevo documento", "Presentar". Todo lo demás es neutro. |
| **Efecto estético-usabilidad** | Lo atractivo se percibe como más fácil de usar y la gente tolera mejor fallas menores ([Laws of UX](https://lawsofux.com/)). | Una sola tipografía (Nunito) con pocos pesos, espaciado en múltiplos de 4/8 y una escala fija de radios. |
| **Gestalt: región común y proximidad** | Lo que comparte un contorno se lee como un grupo. | Cada espacio, documento y ajuste es una tarjeta; los controles relacionados van juntos (segmentados de 2–4 opciones). |
| **Efecto Zeigarnik** | Las tareas inconclusas se recuerdan más y la gente quiere cerrarlas. | En Inicio, lo primero es **Recientes**: los últimos documentos abiertos. `vista-espacios.js · inicio()` |

## 3. Familiar: que no haya que aprender nada nuevo

**Ley de Jakob:** la gente pasa la mayor parte de su tiempo en *otros*
programas y espera que el tuyo funcione igual
([Laws of UX](https://lawsofux.com/)). Por eso cada modo imita al programa que
ya conocen los docentes:

| Modo | Se parece a | Detalles que lo hacen familiar |
|---|---|---|
| Texto | Google Docs / Word | Hoja de papel centrada; Ctrl+B/I/U; Enter crea párrafo; atajos `#`, `-`, `[]`, `>` y menú `/` como en Notion. |
| Hoja | Excel / Sheets | Barra de fórmulas `fx`, letras y números de fila, flechas, Tab, F2, Supr, copiar y pegar; clic en otra celda mientras escribes una fórmula inserta su referencia; funciones en **español y en inglés** (`=SUMA` o `=SUM`). |
| Presentación | PowerPoint / Slides | Miniaturas a la izquierda, diapositiva 16:9 al centro, F5 para presentar, clic o → para avanzar. |
| Diagrama | Miro / draw.io | Pizarrón con puntos, rueda para moverse, Ctrl+rueda o pellizco para zoom, flecha que sale de la figura. |

Atajos universales: Ctrl+Z / Ctrl+Y, Ctrl+S (guarda versión), Esc (cierra),
Supr (borra).

## 4. Decidir rápido y sin cansarse

| Principio | Qué dice | Cómo se aplica |
|---|---|---|
| **Ley de Hick** | El tiempo de decisión crece con el número de opciones ([Laws of UX](https://lawsofux.com/)). | La barra de herramientas muestra **solo las del modo activo**, en el orden en que se usan; las que no se explican con el ícono llevan su nombre a la vista (rastro de información, NN/g). En la computadora nada queda escondido: si no cabe, la barra baja a un segundo renglón. Opciones raras van a un menú "…". Los segmentados nunca pasan de 4 opciones. |
| **Carga cognitiva / memoria de trabajo** | La memoria de trabajo maneja ≈ 4 bloques a la vez. | Cuatro modos, tres secciones en móvil (Inicio, Unirse, Ajustes), un solo campo obligatorio extra al registrarse. Entrar es una sola columna con un botón principal. |
| **Divulgación progresiva** | Mostrar lo avanzado solo cuando se necesita. | La dirección manual del servidor está plegada bajo "¿No aparece?"; las funciones de hoja, en un menú. |
| **Ley de Fitts** | Un objetivo grande y cercano se alcanza más rápido. | Objetivos táctiles ≥ 44 px. En móvil la navegación y las herramientas están **abajo, al alcance del pulgar**, y el botón flotante "+" en la esquina. |
| **Prevención de errores** | Es mejor evitar el error que explicarlo. | Validación al momento con el mensaje junto al campo; medidor de fuerza de contraseña; aviso de Bloq Mayús; confirmación antes de borrar o depurar. |

## 5. Fluido: respuesta y movimiento

| Principio | Qué dice | Cómo se aplica |
|---|---|---|
| **Umbrales de respuesta (Nielsen)** | 0.1 s se siente instantáneo; 1 s es el límite antes de perder el hilo ([SAP Fiori, motion](https://www.sap.com/design-system/fiori-design-web/v1-120/foundations/visual/motion-design)). | Cada cambio se **aplica en local primero** y se envía después (CRDT); escribir nunca espera a la red. Medido: 42 ms entre que un equipo envía y otro recibe. |
| **Umbral de Doherty** | La productividad sube cuando la respuesta baja de ≈ 400 ms ([Laws of UX](https://lawsofux.com/)). | Esqueletos de carga en lugar de pantallas en blanco; si la red tarda más de 2.5 s, el indicador ya dice "Sin conexión" en vez de dejar esperando. |
| **Duración del movimiento** | Microinteracciones ≤ 150 ms; movimientos pequeños 150–250 ms ([SAP Fiori](https://www.sap.com/design-system/fiori-design-web/v1-120/foundations/visual/motion-design), [Val Head](https://valhead.com/?p=2978)). | Tokens `--t-micro` 90 ms, `--t-corto` 140 ms, `--t-largo` 200 ms. Una sola curva (`--suave`), sin rebote. |
| **Respeto al movimiento reducido** | Las animaciones pueden marear a algunas personas. | La preferencia de la app "Reducir movimiento" (Ajustes o Ctrl+K) pone todas las duraciones en 1 ms. Ya no se toma la del sistema: muchos teléfonos tienen las animaciones del sistema apagadas (opciones de desarrollador, ahorro de batería) y la app se veía distinta en el celular que en la computadora. |
| **Regla pico-final** | Se recuerda el momento más intenso y el final. | Al crear cuenta aparece una palomita que se traza; al guardar, "Versión N guardada". |
| **El movimiento explica** | Si quitar una animación no cambia lo que se entiende, sobra. | Las pantallas aparecen ya armadas (sin cascadas). Solo se anima lo que cambia de estado: una ventana que se abre, un aviso que llega, el cajón de versiones. Al cambiar de tema los colores se funden en un instante. No hay pantalla de saludo al entrar: se llega directo al trabajo. |

## 6. Un entorno enfocado, sin distracciones

**Lo que dice la investigación.** Cambiar de tarea deja *residuo atencional*:
parte de la mente se queda en lo anterior, sobre todo si quedó incompleto
(Leroy, 2009). Tras una interrupción, la gente tarda en volver a la tarea
original y suele pasar por otras dos antes; el número exacto de minutos está en
discusión entre fuentes, pero el costo en estrés y esfuerzo sí se repite
([resumen 1](https://blog.saner.ai/context-switching-statistics/),
[resumen 2](https://concepts.dsebastien.net/concept/task-switching-cost/)).
NOM-COLAB nació justo para quitar el cambio entre programas.

| Decisión | Dónde |
|---|---|
| **Un tipo por documento** (cambio oct. 2026): el tipo se elige al crear y queda a la vista junto al nombre; no se puede cambiar para no romper el formato. Las tablas de un texto se editan ahí mismo. | `editor.js` (etiqueta `.ed-tipo`) · `modo-texto.js` (tablas) |
| **Nadie te cambia la pantalla.** Si un compañero cambia el modo del documento, solo aparece un aviso con "Ir"; tú decides. | `editor.js · alCambiar` |
| **Modo enfoque** (Ctrl+.): oculta barras, pasa a pantalla completa y atenúa todo menos el párrafo que escribes. | `.editor.enfoque.atenuar` |
| Avisos **abajo, sin sonido, sin bloquear**, máximo 3, se van solos. Nunca hay ventanas emergentes por algo que no pediste. | `nucleo.js · aviso` |
| La presencia de otros se ve **en el borde del bloque** que editan (color + nombre), no como notificación. | `editor.js · pintarPresencia` |
| Longitud de línea de 50–75 caracteres (`max-width: 70ch`): la lectura más cómoda. | `.b-texto` |
| El editor de texto en escritorio se abre **ya con el cursor puesto** en el primer párrafo vacío. | `modo-texto.js` |

## 7. Micro-textos (CU-22)

Todo ícono tiene un micro-texto: **título + qué hace + atajo**. Aparece al dejar
el cursor 350 ms (para no parpadear al pasar), o con una **pulsación prolongada
de 450 ms** en pantalla táctil (con una vibración corta). Tras la pulsación
prolongada no se ejecuta la acción, así explorar nunca causa un cambio.
`nucleo.js · micro-textos`, atributo `data-micro="Título|explicación|atajo"`.

## 8. Escritorio y móvil, una sola interfaz

| | Navegador (escritorio) | Móvil (navegador o app Android) |
|---|---|---|
| Navegación | Riel lateral: Inicio, tus espacios y Unirse arriba; lo guardado (en este equipo, en el servidor) abajo | Barra superior + pestañas pegadas al borde de abajo + botón "+" |
| Hojas (modales) | Centradas | Suben desde abajo, con agarradera |
| Herramientas del editor | Arriba, con nombre en las que lo necesitan; botones de 32 px (cursor) | **Abajo**, sobre el teclado (la app usa `adjustResize`); solo íconos de 44 px |
| Tipo de documento | Etiqueta de color junto al nombre (no se cambia después de crearlo) | Igual |
| Diapositivas | Miniaturas a la izquierda | Tira horizontal abajo |

La app Android (`android/`) empaqueta **los mismos archivos** de la interfaz
web: no hay dos interfaces que mantener. Solo agrega lo que un navegador no
puede hacer: escuchar el anuncio UDP de la red, guardar en Descargas y abrir el
selector de archivos.

## 9. Accesibilidad

- Foco visible en todo (anillo de 2 px del color de marca con separación).
- Navegación completa con teclado: Tab en hojas atrapa el foco; flechas en menús.
- `aria-label` en todos los botones de ícono; `role="status"` en el indicador de sincronización.
- Tema claro, oscuro o automático (sigue al sistema).
- Campos de 16 px para que iOS no haga zoom al escribir.

## 10. Texto como en Word (y en tiempo real como Google Docs)

| Decisión | Por qué | Dónde |
|---|---|---|
| **Una sola hoja editable** con páginas reales (Carta, A4, Oficio; vertical u horizontal; márgenes de Word), en vez de un cuadro por párrafo. | Es el modelo mental de Word: seleccionar varios párrafos, arrastrar, pegar o borrar a través de párrafos funciona como siempre. | `modo-texto.js` |
| Estilos con los nombres y medidas de Word: Normal (Calibri 11, interlineado 1.15, 8 pt después), Título, Subtítulo, Título 1-3, Cita, Código. | Familiaridad (ley de Jakob) y exportar a .docx con estilos **reales** de Word: el panel de navegación y el índice de Word los reconocen. | `modo-texto.js`, `ooxml.js` |
| Barra de herramientas en el orden de Google Docs: estilo · fuente · tamaño · B I U S · color · resaltado · enlace · alinear · interlineado · listas · sangría · insertar · copiar formato. | Lo que más se usa, a un clic; lo demás en "Más" e "Insertar" (ley de Hick). | `modo-texto.js` |
| Atajos de Word/Docs (lo que en Word hace la minibarra al seleccionar aquí lo hace la ruedita, sección 11): Ctrl+B/I/U, Ctrl+Mayús+L/E/R/J, Ctrl+Alt+1/2/3, Ctrl+Mayús+7/8/9, Ctrl+K, Ctrl+Enter, Mayús+F3… | Las manos no dejan el teclado; quien viene de Word no aprende nada nuevo. | `modo-texto.js · dialogoAtajos` |
| Atajos al escribir: `# `, `## `, `- `, `1. `, `[] `, `> ` al inicio del párrafo. | Autoformato como el de Word y Notion. | `atajosMarkdown` |
| **Cursores de los demás** con su nombre y color, y su selección sombreada. | Se ve dónde escribe cada quien antes de chocar (Google Docs). | `pintarCursores` |
| Deshacer por palabra; Enter, pegar o mover el cursor empiezan un paso nuevo. Solo deshace **tus** cambios. | Como Word; deshacer nunca borra lo que escribió otra persona. | `crdt.js · _inversoSeq` |
| Barra de estado con página actual, conteo de palabras (también de la selección), vista con o sin páginas y zoom. | Información siempre a la vista, sin abrir diálogos. | `pintarEstado` |

## 11. Ruedita de acciones (menú radial)

Inspirada en las ruedas de "ping" de los juegos (Apex Legends, League of Legends,
la rueda de comunicación de Overwatch): un anillo de 8 acciones que se elige **por
dirección**.

**Sale al mantener presionado, centrada en el dedo o el cursor.** Igual en computadora,
celular y la app (probado en un Samsung A55 real). Que salga sola al terminar de seleccionar
resultó molesto, así que solo aparece cuando la persona la pide sosteniendo:

| Dónde | Cómo |
|---|---|
| Texto, con el dedo | La pulsación larga del sistema elige la palabra y la rueda sale en ese momento bajo el dedo. Si se sostiene sobre una selección que ya existía (o una imagen), se conserva esa selección. En la app se quita la barra nativa de Cortar · Copiar · Seleccionar todo (esas acciones están en la rueda; las manijas para ajustar la selección se quedan). En Chrome esa barra es del navegador y no se puede quitar. |
| Texto y celdas, con el mouse | Mantener el clic quieto ~0.55 s, también como pausa mientras se arrastra para seleccionar. |
| Figuras y celdas, con el dedo | Mantener ~0.55 s; se puede deslizar hacia la acción y soltar. |
| Siempre | Clic derecho o tecla Menú: al instante. |

Al soltar sin moverse la rueda se queda abierta hasta que se elige; al hacer la acción, se va. Reemplaza a la barra
pequeña que salía encima del texto. Es un disco opaco con borde y se abre de una vez, sin giro.

| Principio | Qué dice la investigación | Cómo se aplica |
|---|---|---|
| Menús circulares (pie menus) | Callahan, Hopkins, Weiser y Shneiderman (CHI 1988): elegir en un menú circular es más rápido y con menos errores que en uno en lista, porque todas las opciones quedan a la misma distancia del puntero. | 8 sectores iguales; el puntero empieza en el centro. |
| Ley de Fitts | El tiempo para alcanzar un objetivo depende de la distancia y el tamaño. | Cada sector es una cuña que crece hacia afuera: el objetivo es "infinito" en su dirección. |
| Menús de marcas (marking menus) | Kurtenbach y Buxton (CHI 1993/1994): con práctica la gente elige por el gesto, sin mirar el menú. | **El mismo orden en todos los editores** (en el sentido del reloj, desde arriba): 0 Negrita · 1 Cursiva · 2 Subrayado · 3 Color · 4 Eliminar o cortar (abajo) · 5 Copiar o duplicar · 6 Estilo, orden o fórmulas · 7 Más. Lo que no aplica se muestra apagado en su lugar, sin recorrer las demás. Lo que ya está puesto (negrita activa, ancho actual) lleva una marca. |
| Selección que no estorba | La herramienta debe aparecer cuando se necesita y no antes (divulgación progresiva). | Solo sale si se sostiene: seleccionar, tocar o desplazarse nunca la saca. |
| Zona muerta y cancelar | Soltar sin moverse no debe ejecutar nada. | Soltar en el centro cancela; un clic derecho rápido deja la rueda abierta para elegir con clic. |
| Submenús | Mantener la rueda corta (8 máximo). | "Color", "Resaltar", "Estilo" y "Más" abren una segunda rueda en el mismo lugar. |
| Accesibilidad | Debe poder usarse sin ratón. | Tecla Menú o Mayús+F10 la abre en el cursor; flechas, 1-8 y Enter para elegir; Esc cierra. Mayús+clic derecho deja el menú normal del navegador (corrector ortográfico). |

| Contexto | Rueda |
|---|---|
| Texto | negrita, cursiva, subrayado, color (y resaltar), cortar, copiar, estilo (títulos, listas, cita), más (enlace, pegar, copiar y borrar formato, tachado, buscar, imagen, salto de página) |
| Imagen en el texto | ancho completo, mitad, un cuarto, libre / en línea, eliminar, copiar, alinear, más (cortar, descargar) |
| Figuras, cuadros de texto e imágenes del lienzo | negrita, cursiva, subrayado, color (relleno y texto), eliminar, duplicar, orden (frente, fondo, girar, voltear, alinear), más (agrupar, bloquear, copiar, cortar, pegar, editar texto, elegir todo) |
| Celdas | negrita, cursiva, subrayado, color, borrar, copiar, fórmulas (autosuma y funciones), más (cortar, pegar, editar, agregar filas o columnas) |
| Lienzo vacío (clic derecho) | texto, rectángulo, elipse, nota, pegar, imagen, más figuras, ver todo |

`rueda.js` (`enlazar()` decide cuándo sale: al sostener o con clic derecho; cada editor dice qué hay seleccionado).

## 12. Figuras como en Canva

| Función | Cómo se usa |
|---|---|
| Más de 20 figuras (galería con vista previa real) | Botón "Figuras": rectángulo, elipse, rombo, triángulos, pentágono, hexágono, octágono, estrella, corazón, nube, globo de diálogo, flechas, chevrón, cruz, paralelogramo, trapecio, cilindro, documento, nota, líneas. |
| **Barra contextual** fija arriba del lienzo (la ruedita es la que va junto a la selección) | Relleno (o sin relleno), borde (color, grosor, sólido/guiones/puntos), texto (tamaño, color, negrita, alineación, fuente), transparencia, esquinas redondeadas, sombra, posición, girar, voltear, agrupar, bloquear, duplicar, eliminar. |
| Girar | Tirador redondo debajo de la figura; Mayús gira de 15° en 15° y se "pega" a 0°, 90°, 180° y 270°. |
| Cambiar tamaño | 8 tiradores; funcionan aunque la figura esté girada (la esquina opuesta se queda fija). Mayús conserva la proporción. |
| **Guías inteligentes** | Al mover, la selección se alinea sola con los bordes y centros de las demás figuras (y con el centro de la diapositiva); se ven líneas rosas. Alt para mover libre. |
| Selección múltiple | Recuadro arrastrando en el fondo, Mayús+clic, Ctrl+A. Con el dedo: **sostener una figura y tocar otras** con otro dedo, o en la ruedita "Más → Elegir varias" y luego tocar las demás; un aviso abajo cuenta cuántas van y "Listo" termina. Tocar una ya elegida la quita; arrastrar mueve todas. En la hoja, igual: sostener una celda y tocar otra elige el rango; en el texto, sostener la selección y tocar más allá la extiende. |
| Agrupar, alinear y distribuir | Ctrl+G / Ctrl+Mayús+G; alinear a la izquierda/centro/derecha/arriba/medio/abajo; distribuir con espacios iguales. |
| Orden y bloqueo | Ctrl+] / Ctrl+[ (Mayús: al frente / al fondo); bloquear evita moverlo por accidente. |
| Copiar y pegar | Ctrl+C / Ctrl+V entre diapositivas y entre pestañas; Ctrl+D duplica. |

Todo se exporta a PowerPoint con geometrías **nativas** (`star5`, `heart`, `cloud`,
`rightArrow`, `can`…), giro, volteo, borde, transparencia y sombra: se ve igual en
PowerPoint, Google Slides o LibreOffice.

## 13. Estilo protegido

La página se ve como se diseñó, aunque el navegador o el sistema quieran cambiarla:

| Qué podría cambiarla | Cómo se evita |
|---|---|
| Modo oscuro automático de Chrome, Samsung Internet y otros | `<meta name="color-scheme" content="light dark">`: la app declara sus propios temas claro y oscuro. |
| Extensiones de modo oscuro (Dark Reader) | `<meta name="darkreader-lock">`. |
| Extensiones de temas o estilos (Stylus y similares), traductores | Las hojas de estilo que no son de NOM-COLAB (las nuestras llevan `data-nom`) se quitan en cuanto aparecen (`nucleo.js`, `protegerEstilo()`); `notranslate` evita que un traductor reescriba el texto del editor. |
| Alto contraste forzado de Windows y colores forzados del navegador | `forced-color-adjust: none` (y la app ya cumple contraste AA por sí misma). |
| Letra inflada en el celular, tamaño de letra del sistema | `text-size-adjust: 100%`; en la app Android `setTextZoom(100)`. |
| Modo oscuro forzado de Android o de la marca del teléfono | `setAlgorithmicDarkeningAllowed(false)` / `FORCE_DARK_OFF` y `setForceDarkAllowed(false)` en el WebView. |
| Impresión que quita fondos | `print-color-adjust: exact`. |

No se puede impedir el zoom del navegador ni el tamaño mínimo de letra que elige la persona
(son ajustes de accesibilidad del propio navegador), ni lo que una extensión haga con
permisos mayores que los de la página.

## 14. Acabado: lo que hace que se sienta cuidado

Lo que separa una app "que funciona" de una que se siente cuidada no son los efectos, sino cómo
responde. Tres ideas:

1. **La estética cambia la confianza**, y el juicio se forma en los primeros instantes (efecto
   estética-usabilidad: Kurosu y Kashimura 1995, Tractinsky et al. 2000). El efecto se diluye con
   el uso, así que lo que sostiene la impresión es la calidad de la interacción.
2. **Densidad de comportamiento, no de pixeles.** Cada elemento tiene sus estados (normal, cursor,
   foco, presionado, deshabilitado, cargando), los vacíos y errores son específicos y el movimiento
   es corto.
3. **Que la persona sepa qué hacer.** Vacíos que invitan a la primera acción, una lista corta de
   primeros pasos y una paleta (Ctrl+K) que hace descubrible todo y enseña los atajos.

| Detalle | Cómo se aplica | Dónde |
|---|---|---|
| Buscar o ir a | Ctrl+K (Alt+/ dentro del texto, donde Ctrl+K es enlace): documentos de todos los espacios, recientes, espacios, crear, ir a, tema, movimiento, atajos y las acciones del documento abierto. Coincidencia sin acentos y por iniciales; se muestra al instante y los documentos llegan después. | `experiencia.js`, `.paleta` |
| Barra de carga | Arriba, de 2 px, del naranja de la marca; solo si algo tarda > 160 ms y, si sale, dura ≥ 350 ms (no parpadea). | `NOM.cargando`, `.carga` |
| Inicio | Primero **Recientes** y **Tus espacios**; después cuánto ocupa la cuenta y los primeros pasos. | `vista-espacios.js · inicio()` |
| Primeros pasos | Tarjeta con anillo de progreso, debajo de los espacios: crear espacio, primer documento, probar la ruedita, invitar al equipo, ajustar la app. Se marca sola; desaparece al completarla o al ocultarla. En el celular se desliza de lado. | `NOM.primerosPasos`, `.pasos` |
| Consejo de la ruedita | La primera vez en cada tipo de documento, un recuadro explica cómo sale la ruedita; se va al usarla. | `NOM.alAbrirEditor`, `.consejo` |
| Atajos | La tecla `?` (o la paleta) abre todos los atajos agrupados. | `NOM.atajos` |
| Tipografía | Títulos balanceados (`text-wrap: balance`), párrafos sin palabras huérfanas, cifras tabulares en contadores, fechas y celdas. | `nom.css` |
| Avisos | Una línea muestra cuánto les queda; si el cursor se mueve encima se detienen para leerlos. | `aviso()` |
| Miniaturas de documentos | Cada tarjeta muestra una hoja de papel con la forma de su tipo (renglones, cuadrícula, diapositiva o figuras unidas); el nombre va en una línea para que todas queden alineadas. | `NOM.miniatura`, `.dmini` |
| Solo lectura | Con ese permiso la barra del editor solo trae lo que no cambia el documento (buscar, descargar) y dice "Solo puedes ver: también buscar y descargar"; buscar no ofrece reemplazar. | `editor.js · pintarHerramientas` |
| Barras de desplazamiento | Delgadas, del color del borde. | `nom.css` |

### Reglas para cada cambio de interfaz

1. **Solo variables**: ningún color, radio, sombra o duración escrito a mano en un componente.
2. **Un botón principal por pantalla**; los demás, con borde o planos.
3. **El color dice algo o no va**: tipo de documento, estado o marca; nunca decoración.
4. **Sin degradados, desenfoques, resplandores ni texturas** en la interfaz.
5. **Tres tamaños de título, tres pesos**; lo que deba destacar más recibe espacio.
6. **El movimiento explica**: si quitarlo no cambia lo que se entiende, se quita.
7. **Primero lo propio y lo frecuente**; la ayuda y los avisos, después.
8. **Nombre visible** en toda herramienta cuyo ícono no sea universal.
9. **Palabras concretas**: si una frase serviría para cualquier otra app, se reescribe.

## 15. El teléfono como servidor

- **Dónde:** en la app, "Conecta con tu red" muestra primero *Usar este teléfono como servidor*
  (tarjeta con acento, botón Encender). Encendido, la app se conecta a sí misma (127.0.0.1) y al
  pie de Entrar dice "Este teléfono es el servidor · los demás entran en IP:puerto" con
  *Compartir*. En *Este dispositivo* hay un panel con el estado (Encendido/Apagado), cada
  dirección con Copiar y Compartir, *Usarlo en esta app* y *Apagar*. La paleta (Ctrl K) lleva ahí.
- **Arquitectura:** la lógica del servidor vive en `4-CODIGO/nucleo` sin Spring (solo `java.*`).
  El servidor de la computadora la envuelve con Spring (`Configuracion`, `Tareas`, `CanalSse`);
  el teléfono con un servidor HTTP propio (`ServidorHttp`, `ApiTelefono`, `ServicioAnfitrion`) y
  repositorios en archivos JSON con el mismo formato. Las respuestas son idénticas: la misma
  prueba de API (41 casos) pasa en ambos y la forma de cada respuesta coincide.
- **Reglas:** Android 8+ (el cifrado de contraseñas lo necesita); servicio en primer plano con
  notificación y botón Detener, Wi-Fi y procesador despiertos mientras sirve; si Android cierra
  la app, al abrirla se vuelve a encender si la app usaba este teléfono.

## Fuentes

- Lindgaard, Fernandes, Dudek y Brown — *Attention web designers: You have 50 milliseconds to make a good first impression!* (Behaviour & Information Technology, 2006).
- Tuch, Presslaber, Stöcklin, Opwis y Bargas-Avila — *The role of visual complexity and prototypicality regarding first impression of websites* (IJHCS, 2012): https://research.google/blog/users-love-simple-and-familiar-designs-why-websites-need-to-make-a-great-first-impression/
- Fogg et al. — *How do people evaluate a web site's credibility?* (Stanford, 2,684 participantes): https://advocacy.consumerreports.org/research/how-do-people-evaluate-a-web-sites-credibility
- NN/g — Information Scent, Vertical Navigation y 5 Principles of Visual Design: https://www.nngroup.com/articles/information-scent/ · https://www.nngroup.com/articles/vertical-nav/ · https://www.nngroup.com/articles/principles-visual-design/

- Laws of UX (Jakob, Hick, Fitts, Doherty, Von Restorff, estético-usabilidad): https://lawsofux.com/
- SAP Fiori — Motion design (umbrales de 0.1 s y 1 s, duraciones por tamaño): https://www.sap.com/design-system/fiori-design-web/v1-120/foundations/visual/motion-design
- Val Head — How fast should your UI animations be?: https://valhead.com/?p=2978
- Costo del cambio de tarea y residuo atencional (Leroy 2009; Mark, Gudith y Klocke 2008): https://blog.saner.ai/context-switching-statistics/ · https://concepts.dsebastien.net/concept/task-switching-cost/
- Callahan, Hopkins, Weiser y Shneiderman — *An empirical comparison of pie vs. linear menus* (CHI 1988).
- Kurtenbach y Buxton — *The limits of expert performance using hierarchic marking menus* (CHI 1993) y *User learning and performance with marking menus* (CHI 1994).
- Roh, Jeon, Kim y Lee — *Replicated abstract data types: Building blocks for collaborative applications* (JPDC, 2011): el algoritmo RGA.
- Ink & Switch — *Peritext: A CRDT for rich-text collaboration*: https://www.inkandswitch.com/peritext/
