# Relevo del 18 de noviembre de 2026 (madrugada)

**Leer este primero.** Sigue a `RELEVO-17NOV-NOCHE.md` (que queda como detalle de lo de antes).

## 0. Urgente: la Raspberry está apagada y la página pública no contesta

- A las 23:11 del 17-nov se probó el kernel nuevo de la Pi con `kernel-min trial` (arranca **una sola vez** la ranura
  `b/`). **La Pi no volvió** (sin red desde entonces; `https://www.nom.qd.je` da 530 de Cloudflare).
- **Qué hacer:** desconectar la Pi y volver a conectarla. Arranca con el kernel de siempre (`a/`, 6.18.59-min), porque
  `tryboot` es de una sola vez; NOM vuelve solo, con sus datos.
- Después, para saber por qué se detuvo el nuevo: `doas cat /boot/boot-report.txt` (lo escribe a los 30 s y 2 min de
  cada arranque; si el kernel nuevo llegó a arrancar, el del arranque bueno lo sobrescribe: copiarlo antes de 30 s
  no es práctico) y conectar una pantalla HDMI a la Pi para ver dónde se detiene. Sospechas: el controlador NVMe
  de ese disco (SK hynix 1c5c:1339) al iniciar el PCIe, o el árbol 6.18.64 de Raspberry Pi.
- No se levantó un sustituto en la laptop: serviría otros datos y, al volver la Pi, la gente quedaría repartida entre
  dos servidores con el mismo túnel (el defecto que se corrigió el 17-nov).

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

Commits: `8a93bbe` (paneles en PowerPoint real) y `8c35ac4` (archivos «con todo», referencia 3D, pruebas). Nada
publicado todavía (la Pi está apagada).

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

1. **Reiniciar la Pi a mano** (0). Luego: ver por qué se detuvo `b/`, decidir (arreglar `b/`, o usar el de Alpine) y
   pasar los datos al NVMe (`SERVIDOR-PROPIO.md`, «El disco nuevo»).
2. **Publicar la 0.16.0**: ya pasó la regresión completa, la vuelta en PowerPoint real y el emulador. Falta el A55
   (activar *Depuración inalámbrica* y vincular) y la Pi encendida. Pasos en `RELEVO-17NOV-NOCHE.md` §4.1 (4 y 5).
3. **Certificado de `nom.qd.je`**: con `bad_authz` dos veces y el DNS correcto, lo probable es el límite de Let's
   Encrypt para `qd.je` (al no estar en la lista de sufijos públicos, todos los subdominios de todos los usuarios de
   DigitalPlat cuentan como uno) o un bloqueo del dominio. Opciones: esperar y volver a intentar en unos días (quitar
   y poner el dominio), preguntar en el soporte de GitHub, o pedir a DigitalPlat que agregue `qd.je` a la lista de
   sufijos públicos. `https://www.nom.qd.je` (el principal) no depende de esto.
4. **Almacén de binarios** (2.1): el siguiente bloque grande de construcción.
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
