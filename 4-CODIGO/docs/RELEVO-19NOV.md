# Relevo del 19 de noviembre de 2026

**Leer este primero.** Sigue a `RELEVO-18NOV-TARDE.md` (0.17.0 publicada el 18-nov a las 22:20).

Pedido vigente del dueño (18-nov, repetido el 19-nov): todo lo que `PARIDAD-OFFICE.md` marca «No» o «A medias» en
Word, Excel y PowerPoint queda en «Sí», **comprobado contra Office 2024 real en la PC de Windows** (nunca LibreOffice),
con su medida escrita en la fila. Cualquier archivo abre, se ve, se edita y se descarga al 100 %, y funciona con `.nom`,
cifrado, sincronía, sin Internet y el teléfono como servidor (`ApiTelefono` repite cada ruta nueva). Lo que no se toca
de un archivo de Office sale idéntico. Lo que de verdad no se pueda igualar no se salta: se anota aquí qué se intentó,
qué se midió y qué falta. Las filas «decisión del dueño» (p. ej. el Diseñador «sin IA») se quedan como están.

**Terminado** = `PARIDAD-OFFICE.md` sin «No» ni «A medias» (salvo decisiones del dueño o imposibles documentados), cada
«Sí» con su medida contra Office real, la regresión completa de Windows en verde y la versión publicada (0.18.0).

Orden: (0) auditar PARIDAD; (1) lo que queda de cómo se VE un archivo (renglones de Word en tablas, cuadros, junto a
lo que flota, justificado/centrado con tabuladores, guiones con justificado, escala, letra capital; diccionario de
guiones de Word; párrafo vacío con otra letra en la ¶; números de línea; PowerPoint: renglones en miniaturas,
`diapoPng` y al terminar de escribir; SmartArt contra el video; transiciones y Transformar); (2) herramientas de Excel
contra Excel real; (3) crear lo que hoy solo se ve; (4) lo demás.

## 1. Hecho en esta sesión

| Qué | Comprobado |
|---|---|
| Auditoría de `PARIDAD-OFFICE.md` (filas viejas corregidas con su medida) | ver 1a |

### 1a. Auditoría de PARIDAD (19-nov, al empezar)

Medida de partida contra Word 2024 (renglones idénticos, `comparar_real.py`, servidor 8096 con la 0.17.0):
(se llena al terminar la corrida: `~/NOM/trabajo/19nov/base.txt`).

## 2. Cómo seguir

- Servidor de desarrollo: 8096 (interfaz en vivo desde `~/NOM`, datos en `~/NOM/trabajo/srv/datos-8096`).
- Mediciones de esta sesión: `~/NOM/trabajo/19nov/` (no se borra al apagar; `/tmp` sí).
- Regresión: `python3 pruebas/regresion_windows.py --solo a,b` durante el trabajo; la completa antes de publicar.

## 3. Lo que falta (lista real, de la auditoría)

(se llena con la auditoría)
