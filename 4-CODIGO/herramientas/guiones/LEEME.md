# Guiones automáticos como Word

NOM parte las palabras al final del renglón en los mismos lugares que Word 2024 («Disposición → Guiones → Automático»).
Lo que hay aquí sale de Word real, no de un diccionario de otro programa.

## Qué hay

- `palabras-es-MX.txt.gz` y `palabras-en-US.txt.gz`: 83 158 palabras, cada una con un guion donde Word la puede partir
  con ese idioma. Son las 50 000 más frecuentes del español y del inglés (listas de frecuencia de FrequencyWords, solo
  como lista de entrada), las de los archivos de prueba y las de la documentación de NOM; cada palabra se preguntó en
  los dos idiomas (un documento con el idioma en inglés y el texto en español se parte con las reglas del inglés, como
  hace Word).
- `armar.py`: con `patgen` (el programa de TeX que aprende patrones de Liang) saca patrones que dan los mismos lugares
  que Word y escribe `app/src/main/resources/static/js/guiones-es.js` y `guiones-en-us.js`. Las palabras de más de 30
  letras y las que los patrones no dan igual van como excepciones. Con argumentos, primero junta cosechas nuevas.

## Cómo se sacó de Word

`pruebas/office/analisis/guiones_word.py` arma un .docx con Courier New (todas las letras del mismo ancho): cada
párrafo es un relleno de «x» y la palabra, de modo que en el primer renglón caben exactamente k letras de ella más el
guion. Word parte en el lugar más a la derecha que sea ≤ k (o la baja entera); preguntando con k = n − 2 y luego con el
lugar encontrado − 1, en pocas vueltas salen todos los lugares de todas las palabras. Word lo guarda como texto con un
salto en cada renglón (`windows/guiones_txt.py`). Unas 12 000 palabras por documento, ~9 s cada uno.

## Medidas

- patgen reproduce lo de Word en las 83 158 palabras: español 100 % (4393 patrones, 27 KB), inglés 99.99 % (21 624
  patrones y 10 excepciones, 152 KB). `pruebas/office/analisis/guiones_nom.js` comprueba con el mismo código de la
  página: 166 321 de 166 322 iguales.
- En palabras que no se le enseñaron (se entrenó con el 90 % y se probó con el otro 10 %): español 95.8 %, inglés
  78.4 % (los patrones de TeX que había antes daban 96.7 % y 66 % en el mismo idioma, y 40 % con texto en español).
- Cuándo parte Word (no dónde): `pruebas/office/windows/hacer_guiones2.py` hizo 3284 renglones en 4 fuentes, 3 tamaños,
  3 anchos, a la izquierda y justificados, en español y en inglés; `analisis/guiones_decision.py` y el acomodo de
  `modo-texto.js` (`cortesWord`) siguen la regla que salió: 99.4 % de las decisiones a la izquierda y 99.5 % en
  justificado. La zona de división del archivo no cambia nada en Word 2013 y después (medido con 0.1", 0.25" y 0.5").
- Archivos de prueba contra Word 2024: w13-guiones 58.9 → 100 % de renglones idénticos.

## Lo que se sabe de Word (y está en `js/guiones.js` y `modo-texto.js`)

- Nunca deja menos de 3 letras al principio del renglón siguiente en el texto corrido, ni menos de 2 al final del de
  arriba. (Su diccionario tiene lugares que dejan 2: salen en la cosecha, donde la palabra va sola al final del párrafo,
  pero en los 3284 renglones medidos no usó ninguno, tampoco con la última palabra de un párrafo.)
- No parte la palabra que va pegada después de `/ . \ _ @ - : ; , + & = * ~ ^ |`, un número o una raya; sí después de
  `( [ { < ' " # $ % ! ?`. Lo de después de la palabra no estorba. Lo que no tiene espacios y no cabe ni en un renglón
  entero (una dirección web) sí lo parte por cualquiera de sus pedazos, y si no hay lugar, en la última letra que cabe.
- Solo parte dentro de las primeras 64 letras de una palabra.
- A la izquierda parte si sin la palabra quedarían más de 18.3 pt hasta el margen; justificado, si cada espacio se
  estiraría más de 0.4875 de su ancho, en el lugar que deja los espacios más parejos (encoger cuenta el doble). Antes
  intenta meter la palabra entera encogiendo los espacios.
- «Limitar guiones consecutivos a» se respeta; sin límite no hay límite.
