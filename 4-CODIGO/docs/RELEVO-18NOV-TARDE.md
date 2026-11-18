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
| **Notas al final** (endnotes): importadas, numeradas aparte (i, ii, iii), después del último renglón con la raya corta de Word, también con secciones; la raya de las notas al pie también corta como Word | w07 contra Word: se ve donde Word (falta medir el número exacto: ver 3) |
| **Pruebas contra Office real, no LibreOffice**: `funciones_excel.py` (Excel 2024 escribe cada fórmula con `Formula2`: **183 de 183**), `office_abre.py` (Word/Excel/PowerPoint abren lo descargado, sin avisos ni reparar) en `ediciones.py`, `editar_real.py`, `documento_grande.py`, `prueba_pptx.py`, `prueba_diag1.py`; fuera `visor.py` y `fidelidad --visual` | — |
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

1. Terminar de validar y **unir a `main` como 0.17.0**, publicar (push, Pi + `humo.py`, APK `assembleRelease` +
   `publicar-apk.sh`, `avisar.py`). La regresión en Windows de la 0.17.0 casi completa (ver 1).
2. Punto 2 del pedido, lo que falta: división automática de palabras (`w:autoHyphenation`; lo que más baja w10),
   escala de caracteres (`w:w`), el modelo de acomodo de Word con datos (relevo 18-nov §2.2), interlineado por fuente en
   PowerPoint, SmartArt contra el video; notas al pie con marca propia («*», `customMarkFollows`: NOM pone número y
   el asterisco); el índice alfabético a dos columnas se encima (w07 pág. 6).
3. Punto 3: las 143 funciones de Excel que faltan (agregarlas a `funciones_excel.py` y medir contra Excel), cálculo
   manual y F9, nombres, configurar página, rastrear precedentes, evaluar fórmula, herramientas de datos.
4. Puntos 4 y 5 (ver `PARIDAD-OFFICE.md`).
