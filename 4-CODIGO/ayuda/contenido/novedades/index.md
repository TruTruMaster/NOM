---
titulo: Novedades
descripcion: Qué cambió en cada versión, en palabras sencillas.
orden: 1
aplica: Navegador, App Android, Teléfono que comparte
actualizado: 2026-11-21
version: 0.19.0
---
## 0.19.0 — noviembre de 2026

- **Guiones automáticos como Word:** NOM parte las palabras donde las parte Word 2024 (con su diccionario de español y
  de inglés, cada párrafo con su idioma) y solo cuando Word las parte. Un documento de prueba hecho en Word quedó con el
  100 % de sus renglones iguales.
- **Números de línea de Word:** se ven como en Word (cada página, cada sección o continuos, cada cuántos renglones y a
  qué distancia) y se ponen desde «Más» → «Números de línea». Se guardan al descargar el Word.
- **App de Android con su propio lugar dentro del teléfono:** la primera vez que abras esta versión pasa sola todo lo
  que tenías guardado (la sesión, los documentos y los cambios sin enviar), en unos segundos. Para conectarte a otro
  teléfono o computadora que comparte, ese equipo también necesita la 0.19: si tiene una versión anterior, la app lo
  dice.
- **Sin red, más confiable:** al volver la conexión, lo que quedó pendiente (cambios y archivos) se manda solo, también
  en los teléfonos que no avisan que la red volvió.
- **Descargas completas:** «Descargar todo» de un espacio (el .zip con Word, Excel y PowerPoint) y la copia .nom de un
  documento de este equipo traen antes sus imágenes y su archivo de Office original; antes, si no se habían abierto en
  ese momento, podían salir sin ellos.
- **Menús:** un menú largo cabe en la pantalla del celular y se desplaza (antes no se llegaba a lo de abajo, como
  «Modo sugerencia»).
- **Cuentas:** las contraseñas se guardan con un cálculo más fuerte; las de antes siguen sirviendo y se pasan solas al
  formato nuevo al entrar.

## 0.18.1 — noviembre de 2026

- **La página abre mucho más rápido la primera vez** (y después de cada actualización): el navegador ya baja sus
  archivos a la vez y no uno por uno. Medido en `www.nom.qd.je`: de 8.5 s a cerca de 1 s hasta ver la pantalla de entrar.

## 0.18.0 — noviembre de 2026

- **Herramientas de datos de Excel, como en Excel:** Buscar objetivo, nombres definidos y su administrador (la fórmula
  conserva el nombre: `=B2*IVA`), cálculo manual con F9 y Mayús+F9, agrupar y desagrupar filas y columnas (el esquema,
  con sus niveles), mostrar fórmulas (Ctrl+`), rastrear precedentes y dependientes, evaluar una fórmula paso a paso,
  consolidar y tablas de datos. Comprobadas contra Excel 2024: dan lo mismo, y Excel abre lo que descarga NOM con todo eso.
- **Bonos con el primer o el último período irregular** (`PRECIO.PER.IRREGULAR.1` y `.2`, `RENDTO.PER.IRREGULAR.1` y
  `.2`): ya son 458 funciones.
- **Los renglones de Word, donde los pone Word,** también en las celdas de las tablas, en los títulos, junto a lo que
  flota y con versalitas, escala de caracteres, superíndices, sangrías negativas, recuadros y columnas de distinto ancho
  (seis documentos de prueba con el 100 % de los renglones iguales a Word 2024).
- **PowerPoint:** los renglones donde los pone PowerPoint también en las miniaturas, al terminar de escribir en un cuadro
  y en el PNG de la diapositiva (que ahora lleva adentro las fuentes de la página).
- Fuentes nuevas con las medidas de Office: Sitka Text, Lucida Sans Unicode e Ink Free. Las potencias imposibles dan
  `#¡NUM!` o `#¡DIV/0!` como en Excel, y un número en notación científica se lee como número.

## 0.17.1 — noviembre de 2026

- **Más firme con documentos y envíos grandes:** el servidor atiende por turnos lo que llega muy grande y pone topes a
  lo que recibe (un envío de más de 32 MB, un cambio desmedido), sin que nadie note la diferencia al trabajar.
- **Si el servidor se reinicia o está ocupado, lo que escribiste se vuelve a mandar solo** (antes se podía perder lo
  último que alguien escribió en ese momento).
- **Copia en otro NOM:** corre aparte, así guardar los documentos nunca espera a la red, y tiene un tiempo máximo.
- **Entrar** es más rápido cuando entra un salón entero a la vez. Los nombres de cuenta nuevos usan las letras del
  alfabeto latino (con acentos y ñ).
- **App de Android:** la nube va siempre cifrada (HTTPS); el salón sigue igual. Dependencias al día.

## 0.17.0 — noviembre de 2026

- **Las imágenes y el archivo de Office original van aparte del documento** (un almacén por espacio): un
  PowerPoint de 10 MB que antes pasaba del tope de 25 MB (pesaba 30 MB dentro de NOM) ahora entra y el documento
  queda chico. La misma imagen usada en muchas diapositivas se guarda una sola vez. Al descargar sale igual que
  antes, y ya no se pierde el original de los Office de más de 14 MB (hasta 48 MB por archivo).
- Funciona con todo lo demás: **sin Internet** (lo importado se ve y se sube al volver la red), en los **espacios
  cifrados** (las imágenes salen de tu equipo ya cifradas: el servidor no las puede ver), en el **teléfono que
  comparte**, en la sincronía con la nube, en el **enlace para ver** y en el proyector, y dentro del **.nom** que
  descargas (lleva las imágenes, para abrirlo en cualquier lado).
- **Copiar o mover** un documento a otro espacio se lleva sus imágenes. Lo que ya ningún documento usa (ni sus
  versiones ni la papelera) se borra solo a los 30 días, y lo que ocupa cuenta en el espacio de tu cuenta.
- Si la red se corta justo al crear o subir un documento, al reintentar **ya no queda dos veces**.
- **Campos de Word que se actualizan, como en Word:** la fecha, el número de página, los títulos numerados
  («Figura 1», «Tabla II»), las referencias a un marcador y a su página, las condiciones (IF), las fórmulas (=) y las
  propiedades del documento. Se ve su último valor y **Actualizar campos (F9)** los calcula otra vez; Alt+F9 muestra
  sus códigos; **Insertar → Campo** pone uno nuevo. Comprobado contra Word 2024: los 55 campos de prueba dan lo mismo.
- **Imágenes EMF y WMF** (los dibujos de Windows) se ven, también las **ecuaciones viejas** de Word (Editor de
  ecuaciones 3.0 y MathType) y los **objetos incrustados** con su vista previa, como en Word. Al descargar sale el
  archivo original, idéntico.
- **Los renglones de un Word se cortan donde los corta Word.** Se midió en Word 2024 cómo decide (a la izquierda y
  justificado, en ocho fuentes): ahora un documento de Word se ve con las mismas palabras en cada renglón y en cada
  página (en las pruebas, de 85 % a más de 99 % de renglones idénticos). El alto de cada renglón también sale como
  en Word cuando el texto lleva otra letra que la del estilo.
- **Tabulaciones de Word:** las tablas de contenido y los índices importados se ven con sus puntos y el número de
  página alineado a la derecha, como en Word (también las tabulaciones de centro y decimales). Las columnas de una
  sección continua (por ejemplo, un índice a dos columnas) se reparten parejas, como en Word.
- **Guiones automáticos:** si el Word los tiene activados («Disposición → Guiones»), las palabras se parten al final
  del renglón con las reglas del español o del inglés; se activan y desactivan en **Más → Guiones automáticos**.
- **PowerPoint:** el texto de cada cuadro queda a la misma altura que en PowerPoint (se midió dónde pone cada renglón en
  12 fuentes y 8 interlineados), también al presentar y en el proyector.
- **Notas al final** (las de Word, numeradas aparte: i, ii, iii) y la **escala de caracteres** (texto estirado o
  angosto) se ven como en Word.
- **Hoja:** 224 funciones más de Excel (454 en total): estadística, ingeniería, finanzas, fechas, números complejos y
  bases de datos. Comprobadas contra Excel 2024: 499 de 500 fórmulas dan lo mismo.
- **Hoja: Quitar duplicados y Texto en columnas** (en el menú de la columna), como en Excel. Lo que escribes en una
  celda se entiende como en Excel: 1e3, (7), fechas con guiones y los formatos de moneda y porcentaje que pone Excel.
- **Arreglos:** el enlace para ver de un texto o una hoja muestra sus imágenes; un Word guardado como .nom y vuelto
  a abrir ya no reescribe al descargar los párrafos con ecuaciones (salen idénticos al original).

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
  original cuando otra persona también lo cambió: se abre la [rama](ayuda:espacios/ramas), como debe ser.
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
  `INDIRECTO`, `DESREF`). Ver la [lista de fórmulas](ayuda:editores/formulas).
- La **versión** de NOM ahora se ve en *Ajustes y preferencias*.

## 0.13.1 — noviembre de 2026

- **Hasta 50 personas a la vez** en cada espacio, también en el mismo documento (antes 5). Ver
  [Hasta 50 personas](ayuda:espacios/cuantas-personas).
- El tope de **cuentas nuevas por hora** ahora es **por dispositivo**: una escuela entera que sale a
  Internet con la misma dirección ya no se queda sin poder registrarse.
- Los **archivos del chat** cuentan en el almacenamiento del dueño del espacio, y no se pueden
  subir si ya está lleno.

## 0.13.0 — noviembre de 2026

- **Sin Internet no hay cuentas**: en el teléfono que comparte se entra con un
  [alias al azar](ayuda:sin-internet/alias), sin escribir nada; quien tiene Internet conserva su
  cuenta con un [pase](ayuda:sin-internet/con-mi-cuenta).
- **Ramas**: si editas sin Internet y otros cambian el documento, lo tuyo va a una
  [rama](ayuda:espacios/ramas) y el original no se toca; se compara y se fusiona.
- **Crear documentos dentro de un espacio sin Internet**, y botón
  [Instalar en esta computadora](ayuda:empezar/instalar-computadora).
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
