# Paridad con Word, Excel y PowerPoint: investigación

Investigación del 14 de noviembre de 2026 (relevo: `docs/RELEVO-PARIDAD-Y-PRUEBAS-REALES.md`, sección 5).
Fuentes: documentación de Microsoft (lista oficial de comandos de la cinta, support.microsoft.com,
learn.microsoft.com, `[MS-OI29500]`, `[MS-DOCX]`, `[MS-XLSX]`, `[MS-PPTX]`, `[MS-ODRAWXML]`) y los esquemas
de ECMA-376 5.ª edición. El estado de NOM-COLAB se tomó del código (`archivo:línea`).

**Copia parcial** (lo que había el 14-nov a las 01:20). La versión completa se termina en el
escritorio (`flast-desktop:~/investigacion-office/salida/`); al terminar, copiarla aquí encima.

**Estado revisado el 15-nov-2026 (noche): la investigación no terminó.** Se cortó el 14-nov a las 03:10
porque se acabó el **límite semanal** de uso de la cuenta con que corría (se renueva el 18-nov, 3:00 hora de
Nueva York), y el guion que esperaba ya no corre (el escritorio se reinició). Lo que alcanzó a quedar en el
escritorio, más nuevo que esta copia:

| Archivo en el escritorio | Cuánto quedó |
|---|---|
| `WORD.md` (342 KB, 617 filas) | Casi completo: todas las pestañas (Archivo, Inicio, Insertar, Dibujar, Diseño, Disposición, Referencias, Correspondencia, Revisar, Vista, Ayuda, Desarrollador) y el principio de las contextuales. Faltan el resto de las contextuales, los totales, los 40 faltantes más importantes y las fuentes. |
| `POWERPOINT.md` (48 KB) | Solo Inicio e Insertar (igual que esta copia). |
| `EXCEL.md` | No se escribió. |
| `EXCEL-FUNCIONES.csv` (528 funciones) | Completo: nombre en inglés y español, categoría, versión, prefijo en el archivo, si la app la tiene y prioridad. |
| `ooxml-elementos.csv`, `ooxml-atributos.csv` | Iguales a esta copia (preliminares); falta `OOXML-COBERTURA.md`. |

Para traerlo cuando el escritorio vuelva a arrancar en Linux:
`scp 100.84.188.120:investigacion-office/salida/{WORD.md,EXCEL-FUNCIONES.csv} 4-CODIGO/docs/paridad/`.
Para terminarlo: `~/investigacion-office/continuar.sh` en el escritorio, después de que se renueve el límite.

| Archivo | Qué es | Estado de esta copia |
|---|---|---|
| `WORD.md` | Catálogo de Word: cinta completa y diálogos; por fila: qué hace, cómo se guarda (parte, elemento, sección de ECMA-376), si NOM-COLAB lo tiene (con evidencia), qué pasa al importar y exportar, prioridad, cómo implementarlo y en qué archivo del corpus aparece | Archivo, Inicio e Insertar |
| `POWERPOINT.md` | Lo mismo para PowerPoint (más transiciones, animaciones, geometrías y herencia) | Inicio e Insertar |
| `EXCEL.md`, `EXCEL-FUNCIONES.csv` | Lo mismo para Excel y todas sus funciones | todavía no |
| `ooxml-elementos.csv`, `ooxml-atributos.csv` | Cada elemento/atributo de los esquemas y de las extensiones de Microsoft: en cuántos archivos del corpus aparece, si el importador lo lee y si el exportador lo escribe | **Preliminar**: la columna `lo_lee` tiene falsos «no» (p. ej. `sheetData`, `sld`, `document` salen como no leídos); usar la versión final con su `OOXML-COBERTURA.md`, que explica el método |
| `herramientas/` | Los programas que generan las tablas (reproducibles) | parcial |
| `encargo/` | Las instrucciones exactas de cada tarea (para repetirla o extenderla) | completo |
