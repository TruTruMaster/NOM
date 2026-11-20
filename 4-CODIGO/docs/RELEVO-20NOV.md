# Relevo del 20 de noviembre de 2026

**Leer este primero.** Sigue a `RELEVO-19NOV.md`. Publicada: **0.18.1** (20-nov, noche; antes la 0.18.0 a las 16:55: la Pi,
GitHub, el APK y los buscadores), con todo lo del 19-nov (herramientas de Excel, renglones de Word en celdas) y lo de esta sesión. Antes, el
mismo día, salió la **0.17.1** (la 0.17.0 + arreglos de robustez del servidor, la página y el teléfono).

El prompt de cada sesión vive en `~/NOM/prompt.md` (fuera de git): desde hoy es general (sirve igual para cualquier
sesión hasta terminar el proyecto); lo que cambia de una sesión a otra va en el relevo.

Pedido vigente del dueño: todo lo que `PARIDAD-OFFICE.md` marca «No» o «A medias» en Word, Excel y PowerPoint queda en
«Sí», **comprobado contra Office 2024 real en la PC de Windows** (nunca LibreOffice), con su medida escrita en la fila.

## 1. Hecho en esta sesión

| Qué | Comprobado |
|---|---|
| **0.17.1**: topes y validación de lo que recibe el servidor (los envíos grandes, por turnos según la memoria), la copia en otro NOM en su propio hilo con tiempo y tamaño máximos, leer una página para citarla con tiempo total, y el inicio de sesión ya no hace una sola fila para todos (un salón entero entra a la vez) | JUnit 20/20 (3 pruebas nuevas), `pruebas/seguridad.py` 440/440 (11 casos nuevos), `prueba_funciones` y `prueba_red` (§2) |
| La página reintenta un lote de cambios cuando el servidor está ocupado o reiniciándose (429, 408, 5xx); antes lo descartaba y se perdía lo último escrito durante un reinicio de la Pi | — |
| El teléfono como servidor: tope de conexiones y de lo que recibe a la vez; la copia en otro NOM en su propio hilo | `assembleRelease`; APK 0.17.1 firmado con la llave de siempre |
| `prompt.md` general | — |
| **Regresión completa en Windows** sobre `main` (lo del 19-nov + la 0.17.1) | **34 de 34 suites bien** (1412 s; `~/.cache/nom-pruebas/registros/windows-2026-11-20_1624`) |
| **Bonos con el primer o el último período irregular** (`PRECIO.PER.IRREGULAR.1`/`.2`, `RENDTO.PER.IRREGULAR.1`/`.2` = ODDFPRICE, ODDLPRICE, ODDFYIELD, ODDLYIELD; `formulas.js`): 458 funciones. Reglas medidas en Excel: en el último período no hay regla de fin de mes y el largo normal es el real con las bases 1-3; en el primero, los cuasi cupones hasta el primer cupón se cuentan con pasos desde la liquidación (desde su fin de mes si el primer cupón lo es), y el vencimiento tiene que caer en el calendario del primer cupón (si no, `#¡NUM!`) | `funciones_excel.py`: **565 de 575** (66 de los 75 casos nuevos; los cuatro ejemplos de Microsoft exactos; los 9 que no, en `CONOCIDAS`) |
| **0.18.1: la primera carga de la página, ~10 veces más rápida.** El `<meta>` de la política de contenido de `index.html` hacía que Chrome no leyera por adelantado los 55 programas y los bajara de uno en uno (cada uno ~115 ms por Cloudflare y el túnel). Por HTTP la portada ya va sin él (`PaginasLimpias.sinPoliticaEnMeta`, en el servidor y en el del teléfono): la política sigue en la cabecera. La app (archivos locales) lo conserva. | `trabajo/20nov/carga/medir_carga.py` contra `www.nom.qd.je`: pantalla de entrar visible de 9.5 s a 0.9-1.0 s la primera vez; con caché de 0.8 a 0.4 s |
| **0.18.0 publicada** | `humo.py` contra `www.nom.qd.je` 34/34; los dos nombres dicen 0.18.0, interfaz 23 |

## 2. Cómo seguir

- Lo de la sesión del 19-nov sigue igual (servidor de desarrollo 8096, `comparar_real.py`, la PC de Windows): ver
  `RELEVO-19NOV.md` §2.
- **Límites del servidor** para cada cambio y cada envío: `Operacion.valida`, `Secuencia.rangosValidos`,
  `ServicioDocumentos.exigirTamanoEnvio` y `FiltroTamano` (32 MB por envío). Si una función nueva necesita más, que
  mande en partes. En la página, un mapa cuya llave venga del documento va en un `Map`, no en `{}`: los objetos base de
  JavaScript están cerrados (`nucleo.js`) y no aceptan propiedades nuevas.
- La siguiente versión es la 0.18.1 o 0.19.0 (APK 35).
- Bonos irregulares, lo que falta (9 casos, `trabajo/20nov/odd/banco.py` tiene el banco de variantes y
  `trabajo/20nov/funciones-odd*.txt` las medidas): (a) el último período con cupones trimestrales, o con base 3 y
  semestral, cuando el último interés es fin de mes (ni la regla de fin de mes ni sin ella dan lo de Excel: probar
  fechas encadenadas o contadas desde el vencimiento); (b) base real/real con primer período largo y primer cupón a fin
  de mes (E o DSC: probar contarlos con pasos desde la liquidación, como Nq).
- La 0.17.1 salió de la rama `publicar-0.17.1` (la 0.17.0 + los arreglos; etiqueta `v0.17.1`); en `main` están los mismos
  commits encima de lo del 19-nov.
- **`prueba_red` falla en «navegador del celular → página y → app en vivo»** (y a veces en «al volver la red se junta lo
  de los dos lados»), por tiempo: pasa igual con la 0.17.0 publicada (servidor y archivos de la 0.17.0), así que no lo
  causó esta versión. Falta ver si es la prueba (la emulación del celular) o la página.

## 3. Lo que falta

Igual que `RELEVO-19NOV.md` §3, menos las cuatro funciones de bonos (hechas, 66/75). En Excel siguen: el cuadro para
crear tablas de datos desde NOM, relleno rápido (Ctrl+E), escenarios, previsión y `FORECAST.ETS*`, `IMAGE`, disposición
de página e impresión, y el «####» con el hueco «_)». En Word: guiones con el diccionario de Word (w10, w13), texto a los
dos lados de un dibujo (w12), tablas flotantes, cuadros de texto, números de línea. En PowerPoint: SmartArt (p11) y
transiciones/Transformar contra el video. Después, «crear lo que hoy solo se ve» y lo demás de `PARIDAD-OFFICE.md`.

## 4. Cuánto falta (al cerrar el 20-nov)

Unas 58 filas o partes de fila de `PARIDAD-OFFICE.md` en «No» o «A medias» (Word ~25, Excel ~14, PowerPoint ~15). A unas
10-12 funciones medidas contra Office por sesión larga, son varias sesiones más. Esta sesión se fue sobre todo en la
0.17.1 y en publicar la 0.18.0.
