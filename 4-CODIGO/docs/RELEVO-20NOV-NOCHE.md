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

## 2. Cómo seguir

- Servidor de desarrollo: 8096 (interfaz en vivo desde `~/NOM`, datos en `~/NOM/trabajo/srv/datos-8096`; el jar de
  `trabajo/srv/nom-dev.jar` es el de la 0.18.1).
- Mediciones de esta sesión: `~/NOM/trabajo/20nov-noche/`.

## 3. Lo que falta

(se llena al auditar `PARIDAD-OFFICE.md`)

## 4. Cuánto falta

(se llena al cerrar)
