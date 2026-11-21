# Relevo del 20 de noviembre de 2026 (noche)

**Leer este primero.** Sigue a `RELEVO-20NOV.md`. Publicada: **0.18.1** (20-nov, noche: la Pi, GitHub, el APK 35 y los
buscadores). Lo de esta sesión está en `main` sin publicar hasta que se diga otra cosa abajo.

Pedido vigente del dueño (20-nov, noche): seguir con todo lo de `~/NOM/prompt.md` hasta el jueves 26-nov: todo
`PARIDAD-OFFICE.md` en «Sí» contra Office 2024 real, todo funcionando, todo probado y sin errores.

## 1. Hecho en esta sesión

| Qué | Comprobado |
|---|---|
| **Cuentas**: las contraseñas se guardan con versión (`pbkdf2$600000$…`); las de antes siguen entrando, tardan lo mismo y se pasan al formato nuevo al entrar (`Cifrado`, `ServicioCuentas`). En la página, la lectura de secuencias revisa que el bloque sea propio (`crdt.js`) | JUnit 26/26 (`CifradoTest`, `CuentasHashTest`, nuevos), `seguridad.py` 446/446 (`cf57d79`) |
| **Guiones automáticos como Word (diccionario)**: se le preguntó a Word 2024 real dónde parte 83 158 palabras (las 50 000 más frecuentes del español y del inglés, las de los archivos de prueba y las de la ayuda), cada una en es-MX y en en-US (`analisis/guiones_word.py` + `windows/guiones_txt.py`: Courier New, ~12 000 palabras por documento en 9 s). patgen las aprende como patrones de Liang (`herramientas/guiones/armar.py`, `LEEME.md`): `js/guiones-es.js` (4393 patrones, 27 KB) y `js/guiones-en-us.js` (21 624, 152 KB); fuera los de TeX | patrones = Word en 166 321 de 166 322 (`analisis/guiones_nom.js`); en palabras que no vio: es 95.8 %, en 78.4 % (TeX: 96.7 % y 66 %; 40 % con texto en español) |
| **Guiones automáticos como Word (cuándo)**: 3284 renglones hechos en Word (`windows/hacer_guiones2.py`: 4 fuentes, 3 tamaños, 3 anchos, izquierda y justificado, es/en, límite de guiones) y la regla (`analisis/guiones_decision.py`): nunca deja menos de 3 letras abajo; a la izquierda parte si el hueco pasa de 18.3 pt, en el último lugar que cabe; justificado, si cada espacio se estiraría más de 0.4875 de su ancho, donde los espacios quedan más parejos (encoger cuenta doble, ≤ 25 %), y solo si la palabra no entró entera encogiendo; la zona de división no cambia nada desde Word 2013 (medido con 3 zonas guardadas en el archivo). Además: qué signos dejan partir (`/ . @ - …` no; `( " # ! …` sí; medido), lo que no cabe ni en un renglón entero (dirección web) se parte en un guion de cualquier pedazo o en la última letra, solo las primeras 64 letras de una palabra, cada párrafo con su idioma (`w:lang` → `pa.lg`, no se pinta). En `modo-texto.js` (`leerAcomodo`, `cortesWord`, `escribirAcomodo`): el guion es de pantalla (`.wd-guionv`, no editable) | decisiones: 99.4 % a la izquierda, 99.5 % justificado. **w13-guiones 58.9 → 100 %**, **w10 35.1 → 90.4 %**, w15-guiones-z18 97.1 %, lim2 99.1 %; los demás Word igual (w01-w04, w08, w09, w11 100 %; w05 98.5; w06 99.3; w14 99.6; t01 95.4 → 95.6) |
| **Arreglo (venía de antes)**: con «Guiones automáticos» el cursor podía quedar dentro de un guion suave del navegador (`.wd-guion`) y lo escrito ahí se perdía; ahora se saca el cursor al texto de al lado, el lector conserva lo escrito ahí y las posiciones lo cuentan. En los párrafos acomodados como Word los guiones suaves ya no se usan (`hyphens: none`, se quitan y se vuelven a poner si el párrafo deja de acomodarse). Las tabulaciones de pantalla tampoco son editables | `prueba_acomodo.py` 20/20 (9 casos nuevos de guiones: guion de pantalla, idioma por párrafo, copiar, escribir junto al guion, descargar, sin páginas), `prueba_funciones` 204/204 |
| **Medir mejor** (`comparar_real.py`): los renglones de NOM se separan por sus cortes (`br.wd-rompe`: con interlineado exacto menor que la letra dos renglones se encimaban y se contaban como uno) y el texto de Wingdings/Symbol se compara con las letras de siempre (Word lo da en U+F0xx) | w10 |
| **Números de línea de Word** (antes «No»): importa `w:lnNumType` de cada sección (`ooxml.js` `numLineasDe`: Word guarda el inicio − 1) y `w:suppressLineNumbers`; se dibujan en el margen (`modo-texto.js` `pintarNumerosLinea`, después de paginar): renglones del cuerpo y vacíos sí; celdas, cuadros, notas, encabezado/pie y el párrafo vacío del salto de sección no; múltiplos de «cada»; reinicio por página, sección o continuo; distancia automática 0.25" (una columna) o 0.125" (varias), medida; en la línea base del renglón. Más y paleta: ninguno, continuos, cada página, cada sección, suprimir en el párrafo, opciones (`_doc.numLineas`, vale para todo el documento como «Aplicar a todo»). Al descargar: `w:lnNumType` en cada `w:sectPr` (`conNumLineas`) y `w:suppressLineNumbers`; un párrafo regenerado ahora también escribe sus tabulaciones y `w:suppressAutoHyphens` (antes se perdían) | `pruebas/office/numlineas_real.py` (nuevo, en la regresión como `numlineas_word`): w16-numlineas (hecho en Word, `windows/hacer_numlineas.py`) **92/92** números en su hoja, renglón y lugar; vuelta (cambiados en NOM, descargados, abiertos en Word con `windows/numlineas_ver.py`): mismas opciones en las 4 secciones y **60/60**; distancia automática con columnas 92/92 |

## 2. Cómo seguir

- Empezar por aquí; el orden sigue el de `~/NOM/prompt.md` (punto 1: cómo se VE un archivo; luego 2, 3 y 4).
- **Comprobar primero** (quedó corriendo al cerrar): la regresión parcial de Windows de los números de línea
  (`~/.cache/nom-pruebas/registros/windows-2026-11-20_*` más nuevo; suites `office_ediciones`, `documento_grande`,
  `office_fidelidad`, `numlineas_word`, `campos_word`) y `prueba_funciones` (la última corrida no dejó ver su resumen:
  volver a correrla, `NOM_BASE=http://localhost:8096 ~/.cache/nom-pruebas/venv/bin/python pruebas/interfaz/prueba_funciones.py DIR`;
  la anterior, con los guiones, dio 204/204). Si algo falla, arreglarlo antes de seguir.
- Servidor de desarrollo: 8096 (interfaz en vivo desde `~/NOM`, datos en `~/NOM/trabajo/srv/datos-8096`; el jar de
  `trabajo/srv/nom-dev.jar` es el de la 0.18.1; si se cambia Java, `mvn -o -q package` y copiarlo ahí).
- Mediciones de esta sesión: `~/NOM/trabajo/20nov-noche/` (c0…c5: `comparar_real.py --salida`; `guiones/`: cosechas
  de Word, listas y análisis; `depurar/`: guiones de prueba sueltos).
- Guiones: para agregar palabras, `pruebas/office/analisis/guiones_word.py LISTA --idioma es-MX --idioma en-US --salida X.json`
  (~10 s por 12 000 palabras) y `herramientas/guiones/armar.py X.json` (patgen; ver su `LEEME.md`). Antes de cada commit,
  revisar que las listas y los patrones no tengan nombres que prohíben las reglas del proyecto (se quitaron tres palabras de las
  listas por eso).
- La regla de cuándo parte Word está en `cortesWord` (`UMBRAL_GUION_IZQ`, `ESTIRA_GUION_JUST`); los lugares, en
  `guiones.js` (`lugares`, `NO_ANTES`, `MIN_SIGUE`, `LARGO_WORD`).
- **Publicar**: nada de esta sesión está publicado (commits en `main` y en GitHub). La versión que sigue es la 0.18.2 o
  0.19.0 (APK 36), después de la regresión completa de Windows (`python3 pruebas/regresion_windows.py --carriles 10`).
  Trae cambios del servidor (contraseñas con versión): en la Pi primero, como siempre.

## 3. Lo que falta (de `PARIDAD-OFFICE.md`, al 20-nov noche)

**Cómo se ve (punto 1)**
- Word: w12 86.1 % (texto a los dos lados de un dibujo, contornos estrecho/a través, la franja que sube a párrafos de
  arriba del ancla); t01 95.6 % (126 hojas en Word, 121 en NOM); w05 98.5 % (tabla flotante `tblpPr`, filas entre
  hojas); w10 90.4 % (una palabra de 96 letras en un párrafo en-US, la dirección web que sigue y un párrafo a dos
  columnas; ver c3/c4); párrafos justificados o centrados con tabuladores; cuadros de texto (sus renglones no se miden);
  w06 18/17 hojas; los 10 casos de la regla de guiones que no se explican (`analisis/guiones_decision.py`: «interesting»,
  «trabajos», «commitment»… con hueco grande y Word no parte).
- PowerPoint: SmartArt (p11 89.5 %, p06), transiciones y Transformar (Morph) contra el video.
- Excel: el «####» con el hueco «_)» (x10; ver `RELEVO-19NOV.md` §3).

**Herramientas de Excel (punto 2)**: el cuadro para crear tablas de datos desde NOM, relleno rápido (Ctrl+E), escenarios,
previsión y `FORECAST.ETS*`, `IMAGE`, disposición de página e impresión; los 9 casos de bonos irregulares.

**Crear lo que hoy solo se ve (punto 3) y lo demás (punto 4)**: igual que `RELEVO-19NOV.md` §3, menos números de línea
(hecho). Seguridad (plan privado §D): queda el seguimiento de M-3 (WebViewAssetLoader en la app) y B-4; ver el plan.

## 4. Cuánto falta

Unas 56 filas o partes de fila de `PARIDAD-OFFICE.md` en «No» o «A medias» (Word ~23, Excel ~14, PowerPoint ~15, comunes
~11). Esta sesión: guiones de Word (w13 100 %, w10 90 %) y números de línea (de «No» a «Sí», medidos y de vuelta en Word).
