# Relevo del 18 de noviembre de 2026 (madrugada)

**Leer este primero.** Sigue a `RELEVO-17NOV-NOCHE.md` (que queda como detalle de lo de antes).

## 0. La Raspberry volvió (18-nov, 05:54) y por qué el A55 se quedó sin dirección

- El dueño la desconectó y la volvió a conectar: arrancó con el kernel de siempre (`a/`, 6.18.59-min) y NOM, el túnel
  y el vigía volvieron solos; `https://www.nom.qd.je` contesta 200 y los datos y respaldos están intactos. (El reloj
  de la Pi está en EST: marca una hora más que la laptop.)
- **El kernel de prueba (`b/`) no dejó ni una línea en los registros**: se detuvo antes de que arrancara el sistema,
  lo que coincide con la sospecha de un bloqueo temprano (PCIe/NVMe). Para saber dónde se detiene hace falta una
  pantalla HDMI en la Pi durante el arranque de prueba. Con el kernel de siempre el disco sí aparece en el bus PCIe
  (`[1c5c:1339]`); solo falta el controlador (`/dev/nvme*` no existe).
- **La Pi es el DHCP y el DNS de toda la casa** (Pi-hole; el módem no reparte direcciones). Mientras la Pi estuvo
  colgada (de las 23:11 a las 05:54), quien se conectaba a la red no recibía dirección IPv4: eso le pasó al A55 al
  salir del modo avión en el que lo había dejado por error `app_emulador.py` (ya corregido: por Wi-Fi no usa el
  modo avión). Comprobado en las concesiones del Pi-hole: la laptop renovó a las 05:54:27 y el A55 recibió dirección
  a las 05:55:24, un minuto después de que volvió la Pi; desde entonces la depuración inalámbrica conecta sola. Lo
  que ya tenía dirección la conservó, pero sin DNS (el único que reparte el Pi-hole es la propia Pi), salvo los
  equipos que usen uno propio.
- **Antes de otra prueba de kernel**: avisar al dueño y hacerla cuando nadie necesite la red, o encender por un rato
  el DHCP del módem (apagando el del Pi-hole). Pendiente: que un kernel de prueba que se cuelga vuelva solo al de
  siempre (el vigilante de hardware no alcanza a encenderse si se detiene tan pronto).

## 1. Hecho en esta sesión (comprobado)

| Qué | Resultado |
|---|---|
| Rama vieja del 28-oct (su nombre rompía la regla de menciones) | Borrada en GitHub. Respaldo local: `~/NOM/trabajo/respaldos/rama-vieja-28oct.bundle` (traía los scripts del kernel de la Pi; ver 3) |
| `~/NOM/tmp-17-11-26` (3 GB) | Borrado. Quedan 4 carpetas vacías de systemd que son de root: `sudo rm -rf ~/NOM/tmp-17-11-26` |
| Certificado de `nom.qd.je` (GitHub Pages) | Se quitó y se volvió a poner el dominio dos veces (como dice la documentación de GitHub). El DNS está bien desde los dos servidores de DigitalPlat (A y AAAA de GitHub, CAA sin restricción), pero GitHub responde `bad_authz` («la autorización ACME quedó mal»). No depende de nosotros; ver 4 |
| Regresión completa (`regresion.sh`, 68 min) | Todo bien (seguridad 401/401, validador de Microsoft 206/206, funciones contra LibreOffice 178/0, ediciones de Office, Firefox, celular, documento grande) salvo tres suites que resultaron ser **de la prueba**: `rastreo` y `rastreo_movil` (submenú de Pegar especial que ya no aparece, Cortar apagado por Bloquear, la capa de dibujar trayectoria) y `editar_real` (en Firefox cada letra queda como elemento aparte y la búsqueda no encontraba la palabra seguida). Las tres pruebas se corrigieron; `editar_real` con ms-complex2010.docx: 10/10 |
| La vuelta de los paneles en **PowerPoint real** (`vuelta_paneles.py`) | Desde cero (19 efectos con todas las opciones del panel, 5 transiciones, las 9 acciones), sobre p08 (video y audio) y sobre dos archivos con sonidos: ida y vuelta sin diferencias, también por `.nom`; validador 0 errores nuevos; PowerPoint los abre sin avisos ni «reparar» (el único que no abre es un archivo de prueba de LibreOffice cuyo original tampoco abre) |
| ¿Se **ve** igual en PowerPoint? (`video_paneles.py`) | 16 efectos armados con el panel, grabados por PowerPoint (CreateVideo) contra el motor de NOM: **98.3 %** de parecido medio |
| Defectos encontrados y arreglados con eso | El sonido de una transición se perdía al cambiarla; nombres de partes no ASCII (`Cortázar.wav`) dejaban un paquete que ni el validador abría; «atenuar» y «ocultar con el siguiente clic» se aplicaban hasta un clic y PowerPoint los aplica al empezar el siguiente bloque (sondeo `v03-atenuar`: 90.0 → 98.2 %). Catálogos: a01 97.0 → **98.4 %**, a02 98.2 % |
| Emulador de Android (la app con lo último) | `app_emulador.py` **15/15** (texto y hoja en vivo, sin red y al volver), `rama_emulador.py` **6/6** (147 s sin red → rama), y el id del servidor del teléfono **no cambia** al cerrar la app a la fuerza (lo nuevo de la 0.16.0) |
| Archivos «con todo» (idea del dueño) | `windows/hacer_todo.py`: `t01-todo-word.docx` (w01-w12, 126 págs.), `t02-todo-excel.xlsx` (33 hojas), `t03-todo-powerpoint.pptx` (188 diapositivas). Medidos: Excel **100 %** (7 326 celdas, 889 rellenos, anchos y altos; se arregló una referencia 3D con espacios que daba #REF!); Word 83.5 % de renglones idénticos y 121 de 126 páginas; **PowerPoint no entra** (ver 2) |

Commits en `main`: `8a93bbe` (paneles en PowerPoint real), `8c35ac4` (archivos «con todo», referencia 3D, pruebas),
`eb4902e` (este relevo), `c8f5e3f` (versión 0.16.0: interfaz 21, APK 31), `e6951bd` (pruebas) y `fb4be15` (lo que se
manda al servidor ya no pierde las claves «_» de dentro de los datos). Nada publicado todavía.

## 2. Lo más importante que se descubrió

1. **El tope de 25 MB por documento no deja entrar un PowerPoint de 10 MB.** Medido en `t03`: dentro de NOM pesa
   29.8 MB porque lleva el original en base64 (13.9 MB) y **las mismas imágenes otra vez** como datos dentro de los
   bloques (10.4 MB de `src`, 1.7 de fondos, 1.9 de capas). Además, un Office de más de 14 MB no guarda su original
   (al descargar se pierden SmartArt, video y audio). Arreglo de fondo: **la fase 2 de `PLAN-MEDIOS.md` ampliada a
   todo lo binario** (imágenes, el Office original, video y audio) en un almacén del servidor por espacio, con
   permisos, cifrado de extremo a extremo en los espacios cifrados, caché en el equipo para sin Internet, también en
   el teléfono como servidor y dentro del `.nom` al descargarlo; el documento solo guarda la referencia. Es el
   siguiente bloque grande.
2. **Cómo acomoda Word el texto** (para pasar de ~85 % a 100 % de renglones idénticos): en su PDF cada letra avanza
   un número **entero de píxeles a 96 ppp** y el tamaño de la letra se fija en píxeles por em enteros (28 pt → 37 px;
   el PDF lo escribe como 27.775). Con ese modelo se aciertan el 94-97 % de los pasos entre letras del cuerpo
   (`analisis/avances_word.py`); con los anchos exactos, el 10 %. **Pero Word decide dónde cortar con algo más
   cercano a los anchos ideales**: una «NOM A» con avances redondeados empeoró w02 (84.9 → 80.8 %). Y en justificado
   Word a veces encoge espacios (pág. 9 de w02: espacios de 2 px, el renglón llega 2.2 pt más allá del margen) y a
   veces no (pág. 1: no mete «el», que necesitaba menos). El «encoger hasta 20 %» que NOM tiene desde una sesión
   anterior no explica los dos casos (quitarlo arregla 2 renglones y descompone 1). Siguiente paso: generar en Word
   real cientos de renglones controlados (texto al azar, izquierda y justificado, varias fuentes y tamaños), sacar
   los cortes de su PDF y ajustar el modelo con datos, no a ojo.
3. La edición entre varios en un Word real funciona (lo de la falla de la regresión era de la prueba).

## 3. El kernel de la Pi con NVMe (lo que se hizo)

- El kernel `-min` se compiló en la propia Pi con **la misma configuración que corría** (`/proc/config.gz`) más
  `CONFIG_BLK_DEV_NVME=y` y `CONFIG_NVME_HWMON=y`, sobre `raspberrypi/linux` `rpi-6.18.y` (`b0b7d200`, 6.18.64).
  Script: `~/compilar-nvme.sh` en la Pi; copia en la laptop en `~/NOM/trabajo/pi-kernel/` (fuera de git) junto con
  la configuración que corría (`config-min`) y `doas-pi.exp` (corre un comando con `doas` en la Pi pasando la
  contraseña por la terminal). Los scripts de siempre (`kernel-min`, `build-kernel.sh`, fragmentos) están en la
  propia Pi en `/usr/local/share/servidor-alpine/kernel/` y en el respaldo de la rama vieja. Ninguna opción se perdió.
- El árbol 6.18.64 tiene un defecto: `nvme/host/pci.c` usa `sgl_free()` pero `BLK_DEV_NVME` no pide `SGL_ALLOC` (en
  otras configuraciones lo pide el cifrado o SCSI). El script le agrega el `select`.
- Quedó en la ranura `b/` con `kernel-min install`; `kernel-min trial` reinició una vez en `b/` y **no volvió** (ver 0).
  Pánico habría reiniciado solo a los 10 s; el vigilante de hardware a los 15 s si ya corría: por eso se sospecha
  un bloqueo temprano (PCIe/NVMe) antes del vigilante.
- La salida de rescate más simple si se quiere el disco ya: el kernel de Alpine (`s/`, 6.18.58, trae NVMe como
  módulo y `initramfs`): `doas kernel-min rollback s` y reiniciar. Pierde lo «mínimo» del kernel propio.

## 4. Lo que falta (en orden)

1. **La Pi ya volvió** (0). Falta ver con una pantalla HDMI por qué se detuvo `b/`, decidir (arreglar `b/` o usar el
   de Alpine) y pasar los datos al NVMe (`SERVIDOR-PROPIO.md`, «El disco nuevo»). Cuidado: si la Pi se cae, la casa
   se queda sin DHCP ni DNS (0).
2. ~~Publicar la 0.16.0~~ **Hecho el 18-nov**: push a `main` (`98d6ffe`; el trabajo de Azure terminó bien, aunque la
   aplicación de Azure sigue apagada), jar en la Pi (`https://www.nom.qd.je` da 0.16.0, interfaz 21; `humo.py` 34/34),
   APK 0.16.0 (código 31) como release «latest» `v0.16.0` (la descarga fija responde), IndexNow (76 direcciones) y
   las páginas de `nom-qd-je` dicen «NOM». Falta solo instalarla en el A55 (estando en la misma red que la laptop, o
   con el botón «Descarga la app»).
3. **Certificado de `nom.qd.je`**: con `bad_authz` dos veces y el DNS correcto, lo probable es el límite de Let's
   Encrypt para `qd.je` (al no estar en la lista de sufijos públicos, todos los subdominios de todos los usuarios de
   DigitalPlat cuentan como uno) o un bloqueo del dominio. Opciones: esperar y volver a intentar en unos días (quitar
   y poner el dominio), preguntar en el soporte de GitHub, o pedir a DigitalPlat que agregue `qd.je` a la lista de
   sufijos públicos. `https://www.nom.qd.je` (el principal) no depende de esto.
4. **Almacén de binarios** (2.1): casi terminado en la rama `almacen-binarios` (ver 6).
5. **Word: modelo de acomodo con datos** (2.2).
6. Lo demás de `PARIDAD-OFFICE.md` y `RELEVO-ESTADO-Y-PENDIENTES.md` §5.

## 5. Herramientas nuevas

- `pruebas/office/vuelta_paneles.py` — paneles → .pptx → reimportar, `.nom`, validador y PowerPoint real.
- `pruebas/office/video_paneles.py` — 16 efectos del panel: video de PowerPoint contra el motor de NOM.
- `pruebas/office/windows/sondear_atenuar.py` — qué hace PowerPoint con «atenuar» y «ocultar al clic».
- `pruebas/office/windows/hacer_todo.py word|excel|powerpoint` — los archivos «con todo».
- `pruebas/office/analisis/avances_word.py` — posiciones de Word (PDF) contra modelos de ancho (HarfBuzz).
- `pruebas/office/analisis/parrafo_word.py` — estilos y renglones de un párrafo en NOM.
- `animaciones_real.py` ahora da los clics como el video de PowerPoint (al terminar el paso anterior).
- Tres servidores de prueba a la vez: 8099 (regresión), 8097 (medir) y uno de experimentos con otra carpeta
  `static` (8096): así se prueba un cambio sin tocar el repositorio.

## 6. Almacén de binarios (rama `almacen-binarios`, todavía sin unir a `main`)

Las imágenes y el Office original (después también video y audio) van a un almacén por espacio; el documento solo
guarda `nom:bin/<sha256>.<ext>`. Commits: `17bcdf3` (servidor y página), `c03e612` (unión con `main`), `4eb6ac4`
(también en el teléfono que comparte), `f176cb5` (prueba de extremo a extremo) y `a2ffb9d` (la sincronía
teléfono↔nube lleva y trae los binarios).

| Qué | Resultado |
|---|---|
| `t03-todo-powerpoint.pptx` (10 MB) por la interfaz (`pruebas/office/almacen_binarios.py`) | 9/9: se crea (antes «superaría 25 MB»), las imágenes se ven, la descarga sale idéntica parte por parte a la de sin almacén, `.nom` de ida y vuelta, otra persona del espacio y el enlace para ver (sin cuenta) ven las imágenes |
| Seguridad (`seguridad.py`, secciones 18 y 13.5) | 417/417: quién sube y quién baja, huella que no coincide, rutas, el enlace para ver solo entrega lo que su documento usa, la cuota cuenta los binarios; la sincronía sube al otro servidor lo que le falta y baja lo que falta aquí |
| Espacio cifrado (`interfaz/prueba_e2e.py`) | 33/33: el servidor solo guarda la imagen cifrada (su huella es la de lo cifrado) y no ve ni la referencia; los miembros y el enlace con la llave del documento la ven |
| Teléfono que comparte (emulador, APK de depuración de la rama) | La app sube una imagen (170 KB) al almacén del teléfono, el documento queda con la referencia y se baja idéntica |

Falta antes de unirla (en orden):

1. **Investigar primero**: `editar_real.py` en la rama dio 40/41 porque `ms-complex2010.docx`, al importarlo por la
   interfaz, cayó en «Lo creaste sin conexión…» (se guardó para subirlo después). El servidor de pruebas corría de una
   copia del jar (recompilar no lo interrumpía), así que lo más probable es un defecto de
   `NOM.binarios.crearEnEspacio` con ese archivo. Repetir con solo ese archivo y ver qué petición falla.
2. Copiar y mover documentos entre espacios copiando también sus binarios; programar la limpieza de los que nadie usa
   (`ServicioBinarios.limpiar` ya existe).
3. `t01-todo-word.docx` y `t02-todo-excel.xlsx` por el almacén; ayuda, novedades y `PLAN-MEDIOS.md`.
4. Unir a `main` después de publicar la 0.16.0, como 0.17.0 (subir `INTERFAZ`: la página nueva necesita el servidor
   nuevo).

La copia de trabajo de la rama estaba en una carpeta temporal que se borra al apagar; los commits sí están en el
repositorio. Para seguir: `git worktree prune` y `git worktree add <carpeta fuera de /tmp> almacen-binarios`.

## 7. Sesión del 18-nov (mañana): almacén de binarios, casi listo para unir (rama `almacen-binarios`)

Hecho y comprobado (servidor de pruebas 8099 con el jar de la rama):
- **ms-complex2010.docx «Lo creaste sin conexión»**: no se reproduce (solo: 10/10; con los 5 archivos de la regresión:
  50/50). Se endureció lo que lo podía causar: `api.js` arma el JSON antes de pedir (un error al armarlo ya no se
  confunde con «sin conexión») y el tiempo límite cuenta desde que sale la petición; crear un documento lleva una
  **clave de creación** (el servidor da el mismo documento si se repite: `ServicioDocumentos.crear(..., clave)`, también
  en `ApiTelefono`); `crearEnEspacio` no toca el almacén si el sistema dice que no hay red y, si la creación falla por la
  red pero el servidor contesta, la pide otra vez con la misma clave; `guardar(..., esperar)` avisa si no pudo subir.
  Lo creado sin conexión (`local.subir`) sube por el almacén, con su clave guardada (sin duplicados al reintentar).
- **Copiar y mover entre espacios con sus binarios** (`ServicioDocumentos.copiar/mover` + `RepositorioBinarios.copiar`;
  al mover también los de sus versiones). Volver a subir un binario que ya estaba lo «toca» (30 días más de gracia).
- **Limpieza programada**: `ServicioBinarios.limpiarTodo` cada 6 h (Spring `Tareas` y el teléfono).
- `seguridad.py` **423/423** (nuevos 18.6 clave, 18.7 copiar, 18.8 mover).
- **t01 y t02 por el almacén** (`almacen_binarios.py`): **18/18**. Arreglos que salieron de ahí: el enlace para ver de
  Word/Excel no traía las imágenes del almacén (`vista-ver.js` precarga); el HTML de una hoja ahora lleva sus imágenes;
  la prueba busca la hoja con imágenes y mira dentro del marco; y un defecto **anterior al almacén**: tras un `.nom`, un
  párrafo con ecuación salía regenerado (un campo `null` del objeto se perdía y cambiaba la huella): `ooxml.js`
  `mismoGrupo` (las huellas ya no cuentan los campos vacíos y las viejas siguen valiendo).
- t03 por el almacén: **9/9** (29.4 MB sin almacén → documento y versiones 0.5 MB + binarios 17.8 MB; descarga y
  `.nom` idénticos; otra persona y el enlace para ver ven las 84 imágenes).

Falta para unirla a `main` como 0.17.0: ayuda (`4-CODIGO/ayuda/contenido`, `generar.py`), novedades (`CAMBIOS.md`),
`PLAN-MEDIOS.md` (fase 2 hecha), subir versión a 0.17.0 (APK 32) e **INTERFAZ a 22** (`Version.java`, `nucleo.js`,
`sw.js`, `build.gradle.kts`), regresión completa, unir, y publicar (push, Pi + `humo.py`, APK + `publicar-apk.sh`).
Pendiente aparte: el enlace para ver de Word y Excel es el HTML exportado (simplificado), no la vista del editor.
Después sigue el punto 2 del pedido (EMF/WMF, campos, notas al final, guiones, modelo de acomodo de Word…).
