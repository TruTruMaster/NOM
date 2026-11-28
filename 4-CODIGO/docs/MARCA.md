# NOM — Marca: "la llama compartida"

WACK · octubre 2026

Archivos: `4-CODIGO/marca/` (fuentes SVG y PNG) · en la app: `static/img/icono*.{svg,png}`,
el símbolo `#logo` de `index.html` y `android/app/src/main/res/drawable/ic_llama*.xml`.

## 1. La idea

Dos lenguas de fuego, cada una de un color, que **se inclinan una hacia la otra** y juntas
forman **una sola llama**. Cada lengua es una persona; la llama es el documento que hacen
juntas. La separación en S entre las dos es el punto donde se encuentran sus cursores.

- Recoge el lema del equipo, *We carry the flame*, y el color `--llama` (#F06A3D) que la
  interfaz ya reservaba para la marca.
- La otra lengua es índigo. Desde la 0.14.0 el color de la interfaz (`--acento`) es el azul noche
  del fondo del ícono (`#2B2952`), así que el ícono y la app se ven como una sola cosa; el naranja
  queda para el avance (barra de carga) y la marca.
- El fondo es una noche índigo: el fuego se ve de noche, y es donde la gente se reúne a su
  alrededor.

## 2. Por qué funciona (investigación)

| Principio | Qué dice | Cómo se aplica |
|---|---|---|
| **Forma redonda vs. angular** | Jiang, Gorn, Galli y Chattopadhyay (2016, *Journal of Consumer Research* 42(5)) muestran en cinco experimentos que un logo circular activa la idea de suavidad y comodidad, y uno angular la de dureza y fuerza. | La base es un círculo (amable, fácil de usar) y las puntas son agudas (energía, avance). La marca dice las dos cosas que la app quiere decir. |
| **Natural, armónico, moderadamente elaborado** | Henderson y Cote (1998, *Journal of Marketing*, 195 logos): los logos que mejor se reconocen y más agradan son naturales, armónicos y de elaboración moderada. | Una llama es una forma natural que todos reconocen; dos piezas en equilibrio son armonía; dos colores y una separación son elaboración moderada, sin detalles que se pierdan. |
| **El fuego reúne** | Wiessner (2014, *PNAS* 111(39)): de día la gente habla de trabajo; junto al fuego, por la noche, cuenta historias, y esa conversación construye confianza y cooperación. | La metáfora del producto: el lugar donde el equipo se junta a crear. |
| **Un solo elemento dominante** | Las guías de íconos de app coinciden: una forma fuerte, 2 o 3 colores y que se lea a 16 px. Los íconos de Google Workspace de 2020 se criticaron porque todos usaban los mismos cuatro colores y no se distinguían. | El ícono anterior (cuatro cuadros de colores) tenía ese mismo problema y además recordaba al logo de Microsoft. La llama es una sola silueta que se reconoce sin leer. |
| **Colores complementarios** | Azul-violeta y naranja están casi opuestos en el círculo cromático: máximo contraste de tono sin chocar. El azul se asocia a confianza y calma; el naranja a energía y cercanía. | Índigo (confianza, la interfaz) + naranja (energía, el equipo). |
| **Efecto Von Restorff** | Lo distinto a su entorno se encuentra primero. | Casi todos los íconos de una pantalla de inicio son claros; un mosaico oscuro con una llama encendida se encuentra al instante. |
| **Gestalt: cierre** | El ojo completa una figura aunque esté partida. | Se ve una llama, no dos piezas; la separación solo aparece al mirar de cerca. |

## 3. Construcción

Rejilla de 100 × 100 (las coordenadas exactas están en `marca/marca.svg`):

- **Círculo base**: centro (50, 62), radio 24. Las dos lenguas terminan en él.
- **Punta principal** (naranja): (55, 9). **Punta secundaria** (índigo): (34, 29). Las dos
  se inclinan hacia adentro.
- **Muesca**: (44.5, 49). De ahí baja la separación en S hasta la base.
- **Separación**: 3.5 unidades (1.75 a cada lado de la curva).
- La caja de la marca va de x 26 a 75 y de y 9 a 86; su centro, (50.5, 47.5), es el que se
  centra en cada aplicación.
- **En el ícono**: mosaico de 92 con esquinas de 24, marca al 84 %. **Android**: marca a 0.68
  dentro de 108 dp (cabe en la zona segura de 66 dp). **Maskable**: al 67 %.
- **Área de respeto** alrededor del logotipo: la mitad del radio del círculo base del ícono.

## 4. Color

| Nombre | Valor | Uso |
|---|---|---|
| Noche | `#2B2952` | Fondo del ícono |
| Lengua clara (sobre oscuro) | `#E4E4EE` | La lengua izquierda en el ícono |
| Lengua tinta (sobre claro) | `#2B2952` | La lengua izquierda en la marca sola sobre papel |
| Fuego | `#F06A3D` | La lengua naranja en todas las versiones (es el color `--llama` de la interfaz) |
| Tinta / Papel | `#2B2A33` / `#F5F3EF` | Texto del logotipo y fondo, como en la interfaz |

Desde el 27-nov-2026 todos los colores son planos: sin degradados, sin el resplandor de brasa y sin
el brillo de arriba. El degradado morado (`#A9A9FF` → `#6767E8`) era el color que más se asocia con
los diseños hechos en serie, y era lo primero que se veía en cada pantalla. Para volver a aplanar
cualquier SVG del logo: `herramientas/marca/logo_plano.py ARCHIVO...`.

## 5. Logotipo

`NOM` en Nunito 900 junto al ícono de la llama (desde el 27-nov-2026 el logotipo ya no lleva «-COLAB»: el proyecto
se llama solo NOM desde el 25-nov). La interfaz usa la letra del sistema
desde el 27-nov-2026; Nunito queda solo en el logotipo. Las letras están convertidas a trazos en
`marca/logo-horizontal*.svg`, no dependen de tener la fuente instalada.

## 6. Versiones y tamaños

| Archivo | Para |
|---|---|
| `icono.svg`, `icono-512.png`, `icono-1024.png` | App, favicon, presentaciones |
| `compartir.svg` → `static/img/compartir.png` | La imagen que se ve al compartir el enlace |
| `icono-maskable.svg` / `static/img/icono-maskable-512.png` | Android al instalar la página y `apple-touch-icon` |
| `marca.svg` | La llama sola sobre fondo claro |
| `marca-mono.svg` | Un solo color (`currentColor`): sellos, impresión a una tinta, notificaciones |
| `logo-horizontal.svg` / `-oscuro.svg` | Portadas, encabezados, documentos |
| `drawable/ic_llama.xml`, `ic_fondo.xml`, `ic_llama_mono.xml` | Ícono adaptable de Android, con capa monocromática para los íconos con tema (Android 13+) |
| `drawable/ic_notificacion.xml` | Notificación del servidor en el teléfono |

Tamaño mínimo: ícono a 16 px; logotipo horizontal a 120 px de ancho.

## 7. Qué no hacer

- No girar, estirar ni reflejar la llama: la inclinación hacia adentro es el mensaje.
- No pintar las dos lenguas del mismo color sobre el mosaico (para un solo color está
  `marca-mono.svg`).
- No cerrar la separación ni ponerle contorno, sombra, brillo ni degradado.
- No poner la llama naranja sobre fondos naranjas o rojos, ni la índigo sobre índigo.
- No cambiar la letra del logotipo.
