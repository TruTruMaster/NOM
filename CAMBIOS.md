# Cambios de NOM

Lo que cambió en cada versión, en corto (es la página «Novedades» de la ayuda; la escribe `4-CODIGO/ayuda/generar.py`). El diario completo, día por día, está en `LOG.md`.

## 0.16.0 — noviembre de 2026

- **Se llama NOM.** Antes NOM-COLAB. La dirección principal es **www.nom.qd.je**; la de antes
  (nomcolab.dpdns.org) sigue funcionando con las mismas cuentas, y la app las trata como el mismo lugar
  (un QR o un enlace con el nombre de antes ya no la cambia de servidor ni cierra tu sesión).
- **Archivos de Office con contraseña:** se abren pidiendo la contraseña (el cifrado de Office 2010 en
  adelante y el de Office 2007) y se pueden descargar con contraseña. Office los abre igual.
- **Protección como en Office:** «Restringir edición» de Word (solo lectura, solo comentarios, solo
  cambios con control o solo formularios), «Marcar como final», «Contraseña para modificar», proteger
  hojas de Excel con sus permisos (celdas bloqueadas, fórmulas ocultas) y proteger el libro. Lo que trae
  el archivo se respeta y se puede poner o quitar desde **… → Proteger (como en Office)**.
- **Word:** el espacio entre párrafos como Word (el mayor entre el de después y el de antes, sin
  duplicarse), «Conservar con el siguiente», «Conservar líneas juntas», «Salto de página anterior»,
  control de líneas viudas y huérfanas, letra capital, y lo que flota (imágenes, formas, cuadros de
  texto y gráficos) en su lugar de la hoja con el texto alrededor como lo acomoda Word.
- **Word, tablas como en Word:** se parten entre hojas (con las filas de encabezado repetidas arriba), cada fila y
  celda con sus medidas, bordes y texto girado, y celdas con imágenes, listas y otras tablas dentro. Al descargar
  solo se escriben de nuevo las celdas que cambiaron.
- **Word, secciones:** cada sección con su tamaño, orientación, márgenes, columnas (con línea y equilibradas),
  encabezado y pie, y las formas con degradado, trama, imagen, transparencia, puntas y grupos.
- **Word, texto oculto y elevado:** lo que el autor ocultó (formato «Oculto») ya no se ve, como en Word, y se
  conserva al descargar; la letra elevada o bajada (Fuente → Avanzado → Posición) se ve en su lugar.
- **SmartArt** con el dibujo y los colores de PowerPoint (antes algunos salían con letra blanca sobre
  blanco) y los que no traen dibujo, acomodados con su diseño.
- **Animaciones y transiciones como en PowerPoint:** se reproducen como en PowerPoint (medido contra sus
  videos), y ahora también se crean y se editan igual: los mismos efectos con sus opciones, el panel de
  animación con línea de tiempo, copiar animación, trayectorias dibujadas y el panel de transiciones.
  Lo que armes se guarda en el .pptx como lo escribe PowerPoint.
- **Videos y audios de un .pptx** se reproducen al presentar, como en PowerPoint: solos o al tocarlos,
  con su volumen, en bucle y recortados, con una barra para pausar o ir a otra parte. También los **sonidos**
  de los efectos y de las transiciones. (Salen del archivo original que NOM guarda dentro del documento: si el
  .pptx pesaba más de 14 MB no se guarda y se quedan como imagen.)
- **Vínculos y botones de acción** de PowerPoint: ir a otra diapositiva, a la siguiente o la anterior, terminar o
  abrir una página web al tocar la figura; se ponen en **Animación → Acción al hacer clic** y salen así al .pptx.
- **Animar gráficas y SmartArt por partes:** una gráfica por serie, por categoría o por elemento, y un SmartArt
  por figura, como en PowerPoint.
- **Gráficas:** líneas de tendencia con su ecuación y R², barras de error y tabla de datos.
- **Fórmulas:** matrices dinámicas (los resultados se derraman a las celdas vecinas), `LET` y `LAMBDA`.
- **Accesibilidad:** el alto contraste del sistema cambia los colores de la interfaz, «reducir
  movimiento» del sistema se respeta siempre y, si usas una extensión de colores, puedes dejarla
  actuar en **Ajustes**.
- **Privacidad, más clara:** qué puede ver Cloudflare, dónde guarda sus datos la versión en Azure y un
  canal privado para pedir lo de tus datos.
- **App de Android:** cada versión publica su SHA-256 y la huella del certificado con que se firma.
- **Arreglos de una revisión completa:** el código para dejar entrar sin cuenta ya dice que se entra con un
  alias al azar (no con un apodo escrito); la pantalla del proyector ya no recibe comentarios, notas del orador
  ni el archivo original, se mueve con las transiciones de PowerPoint y deja de servir si a quien presenta le
  quitan el acceso; una ecuación que terminaba en barra invertida («\») ya no impide abrir el documento; un archivo de Office de
  más de 14 MB avisa que no se guarda su original.
- **Probado en PowerPoint de verdad:** lo que armas con los paneles de animaciones, transiciones y vínculos se
  abre en PowerPoint sin pedir reparar y se ve igual (medido con el video de PowerPoint). De paso se arregló:
  al cambiar una transición que traía **sonido**, el sonido ya no se pierde; «**Atenuar**» y «Ocultar con el
  siguiente clic» ocurren cuando empieza la animación que sigue, como en PowerPoint; y un archivo con nombres
  internos con acentos (algunos que guarda LibreOffice) sale de NOM como Office lo pide.
- **Excel:** una suma entre varias hojas cuyos nombres llevan espacios (por ejemplo `'Enero 2026:Marzo 2026'!B2`)
  ya no da `#REF!`.

## 0.15.4 — noviembre de 2026

- **Gráficas 3D de Office** (columnas y circular) con la perspectiva y la luz de Excel, la circular
  separada, los estilos de gráfica con degradados y transparencia, y el eje secundario en las gráficas
  combinadas. Las 37 gráficas del libro de prueba quedan entre 92 y 98 % iguales a la imagen que exporta
  Excel.

## 0.15.3 — noviembre de 2026

- **Formas y cuadros de texto de Word** con su geometría, relleno, contorno, giro y texto, en su lugar.
- **Excel alineado como Excel:** vertical, sangría, texto girado y apilado, «Reducir hasta ajustar»,
  «Centrar en la selección», el texto que se desborda a las celdas vecinas y los altos de fila exactos.
- **Rellenos de Excel:** las 18 tramas, degradados y bordes en diagonal.
- **Gráficas nuevas de Excel:** histograma, Pareto, cajas y bigotes, cascada, rectángulos, proyección
  solar y embudo, con los datos vivos de la hoja.

## 0.15.2 — noviembre de 2026

- **Estilos de tabla de Word** con sus franjas, primera y última fila y columna.
- **Numeración como la escribe Word** («Chapter 1», «1.2», letras y romanos por nivel).
- **Imágenes flotantes de Word** con su ajuste (cuadrado, estrecho, arriba y abajo, delante o detrás).

## 0.15.1 — noviembre de 2026

- **Gráficas de Office** dibujadas con las reglas de Office (escala, barras, leyenda, rótulos) en Excel,
  Word y PowerPoint.
- **Lo que va encima de la hoja de Excel:** imágenes, formas, WordArt, SmartArt, conectores, controles de
  formulario y notas.
- **Tablas de PowerPoint** con su estilo, celdas combinadas, rellenos y bordes.

## 0.15.0 — noviembre de 2026

- **Paridad con Office real**, medida contra Word, Excel y PowerPoint 2024 de verdad: 32 familias de
  letra con las medidas exactas de las de Office (sobre dibujos libres), interlineado y justificado como
  Word, formato de números y formato condicional completos de Excel, los estilos de tabla de Excel, y las
  187 formas, rellenos, fondos y líneas de PowerPoint.

## 0.14.2 — noviembre de 2026

- **Documentos largos de Word** (una tesis de 180 páginas): escribir ya no congela la página unos
  segundos cada vez que se acomodan las hojas. Ahora tarda una fracción de segundo y espera a que dejes
  de teclear.
- **Hojas en blanco de más**: cuando un párrafo con formato (negritas, otra letra, un subíndice) caía
  justo entre dos hojas, a veces se cortaba muchas veces y aparecían páginas vacías (un informe de 78
  páginas se veía de 312). Ya no.
- Las **notas al pie** se cuentan en la página donde está su número, aunque el número vaya en
  superíndice.
- **Excel** hecho por otros programas, con rutas internas escritas al revés (`xl\workbook.xml`) o en
  minúsculas: ahora se abre y se descarga bien.
- Los archivos **.doc, .xls, .ppt** (Office 97-2003) y los de LibreOffice (.odt, .ods, .odp) ya se
  pueden elegir al importar, y el aviso dice qué hacer: guardarlos como .docx, .xlsx o .pptx.
- En documentos muy grandes, la copia para abrirlos sin Internet se guarda más espaciada y el conteo
  de palabras se actualiza al dejar de escribir, para que la página no se trabe.
- El servidor responde con el error correcto cuando una petición está mal hecha (antes decía «error
  inesperado»).

## 0.14.1 — noviembre de 2026

- **Direcciones limpias** en la ayuda y en lo legal: ahora son como `/legal/terminos` o
  `/ayuda/preguntas-frecuentes`, sin `index.html`. Las direcciones anteriores siguen funcionando y
  llevan a la nueva.
- **Entrar y crear cuenta**: abajo hay un pie con Ayuda, Preguntas frecuentes, Privacidad, Términos
  y todo lo legal, para consultarlos sin iniciar sesión.

## 0.14.0 — noviembre de 2026

- **Apariencia nueva**, más sobria: fondos lisos, bordes finos, letra menos gruesa y un solo color de
  marca (el azul noche del ícono). El color queda para lo que dice algo: el tipo de documento y los avisos.
- **Entrar** ahora es una sola columna: usuario, contraseña y, debajo, crear cuenta, el código de
  invitación y usar sin cuenta.
- **Inicio**: primero lo último que abriste y tus espacios; cuánto ocupas y los primeros pasos van después.
- **Barra lateral**: tus espacios quedan arriba y lo guardado (en este equipo y en el servidor), abajo.
- **Presentaciones**: las herramientas van en el orden en que se arma una presentación (diapositiva,
  texto y figuras, diseño, transición y animación) con su nombre a la vista, y **Presentar** resalta.
- **Hoja y diagrama**: Autosuma, Funciones, Ordenar y filtrar, Gráfica, Figuras, Conector y Diagramas
  listos muestran su nombre junto al ícono.
- En la computadora ya no quedan herramientas escondidas a la derecha: si no caben, bajan a un
  segundo renglón.
- Menos movimiento: las pantallas aparecen ya armadas y la app entra directo, sin pantalla de saludo.

## 0.13.2 — noviembre de 2026

- **Ramas**: lo que escribes sin Internet durante más de 2 minutos ya no se junta solo con el
  original cuando otra persona también lo cambió: se abre la rama, como debe ser.
- **PowerPoint**: al cambiar el formato o el texto de una diapositiva importada se conservan sus
  **animaciones**, y los títulos siguen siendo títulos al descargar el archivo.
- **Word**: al editar un párrafo de un archivo importado ya no se pierde lo que NOM no muestra
  (marcadores, comentarios, campos, idioma del texto, controles de las plantillas).
- **Códigos de invitación**: se pueden escribir con o sin guion, y también **pegar el enlace
  completo**; en la app, tocar un enlace de invitación abre la app.
- En el celular: **toca un cuadro ya elegido para escribir** en él; la tecla Atrás cierra primero el
  menú abierto; en la hoja, **Enter después de Tab** regresa a la columna donde empezó el renglón.
- **Hoja de cálculo: 89 funciones nuevas** (ya son 185): financieras (`PAGO`, `VA`, `VF`, `TASA`,
  `TIR`, `VNA`), de fecha y hora (`HORA`, `NUM.DE.SEMANA`, `DIA.LAB`, `DIAS360`), estadística
  (`PERCENTIL`, `CUARTIL`, `PENDIENTE`, `PRONOSTICO`, `COEF.DE.CORREL`), texto (`TEXTOANTES`,
  `TEXTODESPUES`), lógicas (`CAMBIAR`, `SI.ND`, `XO`) y de referencia (`FILA`, `COLUMNA`,
  `INDIRECTO`, `DESREF`). Ver la lista de fórmulas.
- La **versión** de NOM ahora se ve en *Ajustes y preferencias*.

## 0.13.1 — noviembre de 2026

- **Hasta 50 personas a la vez** en cada espacio, también en el mismo documento (antes 5). Ver
  Hasta 50 personas.
- El tope de **cuentas nuevas por hora** ahora es **por dispositivo**: una escuela entera que sale a
  Internet con la misma dirección ya no se queda sin poder registrarse.
- Los **archivos del chat** cuentan en el almacenamiento del dueño del espacio, y no se pueden
  subir si ya está lleno.

## 0.13.0 — noviembre de 2026

- **Sin Internet no hay cuentas**: en el teléfono que comparte se entra con un
  alias al azar, sin escribir nada; quien tiene Internet conserva su
  cuenta con un pase.
- **Ramas**: si editas sin Internet y otros cambian el documento, lo tuyo va a una
  rama y el original no se toca; se compara y se fusiona.
- **Crear documentos dentro de un espacio sin Internet**, y botón
  Instalar en esta computadora.
- **Cifrado en el servidor**: documentos, versiones, bitácoras, cuentas y respaldos van cifrados
  en el disco; y **espacios cifrados de extremo a extremo** con **chat** del espacio.
- Este sitio de **ayuda** y las páginas **legales**, también sin Internet.
- El código ahora es software libre bajo la **GPL-3.0**.

## 0.12.3

- Si el navegador o la app se cerraban de golpe, al volver sin Internet ya no pide iniciar sesión.

## 0.12.2

- Revisión completa en equipos reales: 17 arreglos (tope de 25 MB revisado antes de guardar,
  pegar en la red del salón, filas y columnas en el celular, notas al pie en el celular y más).

## 0.12.1

- Más seguridad en las cuentas: bloqueo por red, tope de cuentas nuevas, cambiar la contraseña,
  borrar la cuenta. Se ve qué está en el servidor y cuánto ocupa tu cuenta.

## 0.12.0

- Conversión con Office mucho más fiel: el archivo original va dentro del documento y lo que no
  tocas sale idéntico. Formato .nom comprimido.

## 0.11

- Hojas, textos, presentaciones y diagramas con lo que tienen Word, Excel, PowerPoint y Canva;
  investigar y citar en APA; papelera, carpetas y plantillas.

## Antes

- El teléfono puede ser el servidor; la ruedita; unirse con QR; el recibo de tu trabajo; trabajar
  sin conexión y juntar al volver.
