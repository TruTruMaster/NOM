# Relevo del 18 de noviembre de 2026 (tarde)

**Leer este primero.** Sigue a `RELEVO-18NOV.md` (madrugada; su sección 7 es la mañana de este mismo día).

Pedido vigente del dueño (18-nov): que todo lo que `PARIDAD-OFFICE.md` marca «No» o «A medias» quede en «Sí» en Word,
Excel y PowerPoint, **comprobado contra Office 2024 real** (nunca contra LibreOffice: lo pidió con enojo el 18-nov), con
los archivos «con todo» (t01-t03), y que todo funcione con `.nom`, cifrado, sincronía, sin Internet y el teléfono
como servidor. Orden: (1) almacén de binarios → 0.17.0; (2) lo que cambia cómo se VE un archivo (EMF/WMF, campos,
notas al final, guiones, escala de caracteres, modelo de acomodo de Word, interlineado por fuente en PowerPoint,
SmartArt contra el video); (3) las 143 funciones de Excel que faltan y las herramientas de Excel; (4) crear lo que hoy
solo se ve; (5) lo demás.

## 1. Hecho en esta sesión (todo en la rama `almacen-binarios`, que se vuelve la 0.17.0)

| Qué | Comprobado |
|---|---|
| Almacén de binarios terminado (relevo 18-nov §7): clave de creación sin duplicados, copiar/mover con binarios, limpieza cada 6 h, lo creado sin conexión sube por el almacén | `seguridad.py` 423/423; t01, t02, t03 por el almacén 18/18 y 9/9 |
| Versión 0.17.0 (interfaz 22, APK 32), ayuda, novedades, `PLAN-MEDIOS.md` (fase 2 hecha para imágenes y el original) | — |
| **EMF y WMF** (`js/emf.js`): conversor propio de los registros GDI a SVG (líneas, curvas, texto con posiciones de cada letra, Symbol/MT Extra/Marlett a Unicode, mapas de bits con máscaras, tramas, degradados, recortes, modos de mapeo). El documento guarda el original (sale idéntico); se convierte al dibujar (`binarios.js`). Objetos OLE de Word (ecuaciones viejas, hojas incrustadas) con su vista previa, también dentro de tablas | `metarchivos_real.py`: **Word** coloca cada uno de los 535 metarchivos del corpus en su página y exporta a PDF; parecido 96.2 %, 88.5 % en lo dibujado (las letras dentro del SVG usan las del equipo: un `<img>` no ve las fuentes de la página) |
| **Campos de Word** (`js/campos.js`): DATE/TIME/CREATEDATE/SAVEDATE con `\@`, SEQ (`\c \r \h \s`), REF, PAGEREF (`\p`), PAGE, NUMPAGES, IF, `=` (funciones y `\#`), `\*` (ROMAN, ALPHABETIC, CardText, OrdText, Ordinal, DollarText, Upper…), AUTHOR, TITLE, DOCPROPERTY, STYLEREF, FILENAME, NUMWORDS. Como Word: se ve el último valor, F9 actualiza todos, Alt+F9 códigos, Insertar → Campo; en la ruedita (Más) y la paleta. Marcadores que usan los campos importados como marcas invisibles; propiedades del documento e idioma de cada campo | `campos_real.py`: Word arma 55 campos y los actualiza; **55 de 55 iguales** en NOM tras F9; Word abre lo descargado sin avisos |
| **Escala de caracteres** (`w:w`, 50-200 %) y **notas con marca propia** («*»): como Word. El comparador cuenta la llamada de una nota como la verdad de Word (un carácter de control): w07 90.4 → 93.0 % | `comparar_real.py`; los 12 Word sin bajar |
| **Notas al final** (endnotes): importadas, numeradas aparte (i, ii, iii), después del último renglón con la raya corta de Word, también con secciones; la raya de las notas al pie también corta como Word | w07 contra Word: se ve donde Word (falta medir el número exacto: ver 3) |
| **Pruebas contra Office real, no LibreOffice**: `funciones_excel.py` (Excel 2024 escribe cada fórmula con `Formula2`: **183 de 183**), `office_abre.py` (Word/Excel/PowerPoint abren lo descargado, sin avisos ni reparar) en `ediciones.py`, `editar_real.py`, `documento_grande.py`, `prueba_pptx.py`, `prueba_diag1.py`; fuera `visor.py` y `fidelidad --visual` | — |
| **Funciones de Excel** (`js/formulas.js`): +224 en dos lotes (trigonometría, redondeos, días laborales, FRAC.AÑO, distribuciones e inversas, BD*, depreciaciones, TIR/VNA con fechas, ingeniería, números complejos, Bessel, pruebas T/F/Z/ji², ESTIMACION.LINEAL, CRECIMIENTO, bonos, cupones, letras del Tesoro, ESFORMULA): **454** | `funciones_excel.py` contra Excel 2024 real: **499 de 500**; la que no: `BESSELJ` con x ≥ 8 (Excel se aparta ~5e-7 del valor exacto; se probaron A&S, Hankel, Miller, series e integrales; anotada como conocida en la prueba) |
| **División automática de palabras** (`js/guiones.js`, patrones de TeX en `js/vendor/`, con su licencia): «Guiones automáticos» del archivo (zona, mayúsculas, «no dividir» por párrafo), interruptor en Más y en la paleta, se guarda al descargar; los guiones son de pantalla | w13-guiones 19.6 → 57.1 % de renglones idénticos (`windows/hacer_guiones.py`); la zona a la mitad es lo que más se parece a Word |
| **Renglones cortados como Word** (`acomodarComoWord` en `modo-texto.js`): con 144 párrafos hechos en Word real (`windows/hacer_renglones.py`) se sacó la regla (`analisis/acomodo_word.py`): a la izquierda Word usa los anchos exactos de la fuente sin tolerancia; justificado, mete la palabra que no cabe encogiendo los espacios solo si se pasa a lo más la mitad del hueco que quedaría y hasta 25 % por espacio (0 errores en 559 decisiones); corta después de guion o raya pegados. El navegador mide hasta 0.1 % más angosto (tamaño redondeado a 1/64 px): NOM calcula los cortes con lienzo a 1000 px y los fuerza con `<br class="wd-rompe wd-pintado">` (el lector del DOM, las posiciones y copiar los ignoran); justificado con espacios a −30 % y `text-align-last: justify`, el último renglón con un relleno. Además: alto de renglón con la letra de la marca de párrafo (`w:pPr/w:rPr` → `tm`, `fm`) y sangrías con 4 decimales (con 2 se movían hasta 0.14 pt) | w14-renglones 85.2 → **99.6 %** (23 páginas como Word; antes 26), w01 86.1 → 94.9, w02 84.9 → 94.1, w04 86.5 → 93.4, w07 93.0 → 96.5, w08 → 100, w09 78.5 → 93.3, w12 77.8 → 82.9, t01 83.8 → 89.6; `prueba_acomodo.py` 9/9; ediciones docx 98/98; funciones 203/203 |
| **Tabulaciones de Word** (importa `w:tabs` del estilo y del párrafo con «clear» y `defaultTabStop`; `anchoTab` en `modo-texto.js`: izquierda, derecha, centro y decimal, sangría francesa como parada; un `.wd-tabw` de pantalla con su ancho y su relleno; la tabulación misma con `tab-size: 0`). **Columnas de sección continua** equilibradas como Word (las pruebas del equilibrio ya no toman el margen empujado como espacio; el párrafo vacío del salto de sección no cuenta; 2 px de holgura al colocar). **Arreglo**: los atributos que no se pintan (`tm`, `fm`, `sinGuiones`) fuera de `PAR` (el lector los daba por borrados con cada tecla y el cursor saltaba: lo encontró `documento_grande`) | w09 78.5 → **98.2 %** (9 hojas como Word), w02 96.2 %, w04 94.3 %, t01 90.8 %; `prueba_acomodo.py` 11/11 (también tabulación derecha con puntos); `documento_grande` 17/17 |
| **Arreglos que encontró la regresión en Windows**: `w:sym` con la fuente Symbol se importa como su letra Unicode (`textoSym` en ooxml.js, con la tabla de `emf.js`: «κνβερναω», antes «knbernaw» en pantalla y al descargar); los cambios de pantalla de `revisarAcomodos` se descartan del observador (`tomarRegistros`), porque con la siguiente tecla el editor releía párrafos ajenos y reescribía al descargar uno que no se tocó; la prueba de la hoja vuelve a escribir las celdas que no quedaron (con carga, el clic por coordenadas caía en otra celda; no se pudo reproducir una pérdida de teclas de la app ni con la CPU 6× más lenta ni en celular) | `documento_grande` 17/17 y ediciones docx 98/98 en local |
| **PowerPoint: cada renglón donde lo pone PowerPoint** (`ajustarRenglones` en `diapositivas.js` con la tabla `PP_BASE`; el párrafo importado lleva la letra de su primer texto —sin eso el renglón crecía al mezclar fuentes—; interlineado exacto redondeado a puntos). Datos: `windows/hacer_interlineado.py [--mas]` → p12/p13; análisis `analisis/interlineado_ppt.py` (la tinta de la línea base: texto «HxxH», sin descendentes) | p12 ±0.5 pt (antes hasta 10); p01 mediana 2 → 1 px, p07 4 → 1.5 px; prueba_pptx, pptx2, pres1, diag1 bien; ediciones pptx 78/78 |
| **Regresión en paralelo y en la PC de Windows** (20 hilos, 32 GB, Office ahí): `pruebas/regresion.py` (carriles con su propio servidor), `pruebas/regresion_windows.py` (manda código, jar y archivos y corre allá), `office/windows/preparar_windows.ps1` (Java 17, Playwright, .NET 8: ya instalados), `office/remoto.py` (Office por SSH o local) | 1.ª corrida completa: ~25 min contra ~70 en la laptop; fallas que eran del entorno, ya corregidas: candado de archivos en Windows (seguridad), `prueba_limites.py` con ruta fija, botones que no esperaban con carga |

## 2. Cómo seguir

- Servidores de esta sesión: 8099 (regresión vieja, ya no hace falta), 8096 (desarrollo, interfaz en vivo desde
  `~/NOM`, datos en `~/NOM/trabajo/srv/datos-8096`), 8097 (vacío). Jar en `~/NOM/trabajo/srv/`.
- Regresión: `python3 pruebas/regresion_windows.py` (completa, ~25 min) o `--solo a,b`. Los registros quedan en
  `~/.cache/nom-pruebas/registros/windows-<fecha>/`. `junit` corre en la laptop al compilar.
- Medir Word: `comparar_real.py` con `NOM_BASE=http://localhost:8096`. Ojo: un campo DATE ahora se ve con la fecha de
  hoy (como Word al abrir), y la verdad de Word se sacó el 15-nov: ese renglón sale distinto en w07.
- Lo de Windows se manda por `remoto.py` (un solo paquete; copiar cientos de archivos sueltos por SSH se atora).

## 3. Lo que falta (en orden)

1. **Unir a `main` como 0.17.0 y publicar** cuando la corrida completa en Windows salga bien (la 4.ª: 30 bien y 3
   fallas que eran de las pruebas —ya corregidas—; la 5.ª va con los renglones de Word). Publicar: push a `main`,
   Pi (`ssh pi5 /usr/local/sbin/nomcolab-pi actualizar`, con la ruta completa) + `humo.py`, APK `assembleRelease` +
   `publicar-apk.sh`, `avisar.py`; el A55 cuando esté en la misma red. **La Pi es el DHCP y el DNS de la casa: no
   reiniciarla.**
2. Renglones de Word, lo que falta: párrafos con tabuladores, con guiones automáticos (combinar la zona con la regla
   del justificado), con escala de caracteres o letra capital, junto a objetos que flotan (hoy los acomoda el
   navegador) y el texto de tablas y cuadros de texto; w10 (texto español marcado en inglés: Word parte con su
   diccionario, p. ej. «extra-ordinariamente», «con-exión»: juntar con Word real cientos de palabras partidas y
   usarlas como excepciones); un renglón de Verdana 10 que Word deja pasar 0.13 pt.
3. Punto 2 del pedido, lo demás: SmartArt contra el video; en PowerPoint, acomodar también las miniaturas y la imagen
   PNG de cada diapositiva (`diapoPng` arma un SVG sin poder medir), y volver a acomodar el cuadro al terminar de
   escribir en él (hoy espera al siguiente repintado). (Hechos el 18-nov: el interlineado por fuente en PowerPoint y el
   índice a dos columnas.)
4. Punto 3: `FORECAST.ETS*`, `ODDF*`/`ODDL*`, cálculo manual y F9, nombres, configurar página, rastrear precedentes,
   evaluar fórmula, herramientas de datos.
5. Puntos 4 y 5 (ver `PARIDAD-OFFICE.md`).

## 3b. Lo que se había escrito antes en esta sección (ya hecho o reacomodado arriba)



1. Terminar de validar y **unir a `main` como 0.17.0**, publicar (push, Pi + `humo.py`, APK `assembleRelease` +
   `publicar-apk.sh`, `avisar.py`). La regresión en Windows de la 0.17.0 casi completa (ver 1).
2. Punto 2 del pedido, lo que falta: división automática de palabras (`w:autoHyphenation`; lo que más baja w10:
   Word parte con las reglas del **idioma del texto** —en w10 en-US sobre texto en español: «con-exión», «in-cluye»,
   «her-ramientas»—; plan: patrones de TeX (hyph-en-us es de copia libre; revisar la licencia de hyph-es) y guiones
   suaves SOLO en lo pintado: el lector del DOM (`tokens`, la reconciliación) tiene que ignorarlos sin confundirlos con
   los guiones suaves que sí son del documento —`w:softHyphen` ya se guarda como U+00AD—, p. ej. pintándolos dentro de
   un elemento marcado; respetar `hyphenationZone`, `consecutiveHyphenLimit` y `doNotHyphenateCaps`),
   el modelo de acomodo de Word con datos (relevo 18-nov §2.2), interlineado por fuente en
   PowerPoint, SmartArt contra el video; el índice alfabético a dos columnas se encima (w07 pág. 6).
3. Punto 3: las 143 funciones de Excel que faltan (agregarlas a `funciones_excel.py` y medir contra Excel), cálculo
   manual y F9, nombres, configurar página, rastrear precedentes, evaluar fórmula, herramientas de datos.
4. Puntos 4 y 5 (ver `PARIDAD-OFFICE.md`).
