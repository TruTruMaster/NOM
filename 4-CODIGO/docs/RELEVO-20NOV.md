# Relevo del 20 de noviembre de 2026

**Leer este primero.** Sigue a `RELEVO-19NOV.md`. Publicada: **0.17.1** (la Pi, GitHub y el APK) = la 0.17.0 + arreglos
de robustez del servidor, la página y el teléfono (`CAMBIOS.md`). Lo del 19-nov que está en `main` (herramientas de
Excel, renglones de Word en celdas) **todavía no está publicado**: sale en la 0.18.0, cuando pase la regresión completa
de Windows.

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

## 2. Cómo seguir

- Lo de la sesión del 19-nov sigue igual (servidor de desarrollo 8096, `comparar_real.py`, la PC de Windows): ver
  `RELEVO-19NOV.md` §2.
- **Límites del servidor** para cada cambio y cada envío: `Operacion.valida`, `Secuencia.rangosValidos`,
  `ServicioDocumentos.exigirTamanoEnvio` y `FiltroTamano` (32 MB por envío). Si una función nueva necesita más, que
  mande en partes. En la página, un mapa cuya llave venga del documento va en un `Map`, no en `{}`: los objetos base de
  JavaScript están cerrados (`nucleo.js`) y no aceptan propiedades nuevas.
- Publicar la 0.18.0: como siempre (`prompt.md`, «Publicar»), con la versión siguiente a la 0.17.1 (APK 34).
- La 0.17.1 salió de la rama `publicar-0.17.1` (la 0.17.0 + los arreglos; etiqueta `v0.17.1`); en `main` están los mismos
  commits encima de lo del 19-nov.
- **`prueba_red` falla en «navegador del celular → página y → app en vivo»** (y a veces en «al volver la red se junta lo
  de los dos lados»), por tiempo: pasa igual con la 0.17.0 publicada (servidor y archivos de la 0.17.0), así que no lo
  causó esta versión. Falta ver si es la prueba (la emulación del celular) o la página.

## 3. Lo que falta

Igual que `RELEVO-19NOV.md` §3 y §4 (unas 60 filas o partes de fila de `PARIDAD-OFFICE.md` en «No» o «A medias»), más
lo que se haga y se mida a partir de aquí (ver §4 de este relevo cuando la sesión cierre).
