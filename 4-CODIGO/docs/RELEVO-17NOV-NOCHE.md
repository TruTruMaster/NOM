# Relevo del 17 de noviembre de 2026 (noche): lo que ya quedó corregido y lo que falta

**Leer este primero.** Es el relevo que manda desde la noche del 17-nov. El detalle (tablas de medición, la lista de
emparejamiento de cada función, el orden por fases) sigue en `RELEVO-ESTADO-Y-PENDIENTES.md`; el inventario de Office
contra NOM, en `PARIDAD-OFFICE.md`.

## 1. Dónde estamos (comprobado a las 23 h)

| Qué | Estado |
|---|---|
| Código (`~/NOM`, rama `main`) | `a22b6bb`: **28 commits sin publicar** sobre `origin/main` (`b65438d`, la 0.15.4). Dice 0.15.4 en el código; la ayuda y `CAMBIOS.md` ya describen la 0.16.0. |
| Publicado | Raspberry: 0.15.4, interfaz 20. GitHub: release «latest» v0.15.4 (APK código 30). Azure: apagada (403). |
| `https://www.nom.qd.je` y `https://nomcolab.dpdns.org` | Contestan (la Pi, por el túnel de Cloudflare). |
| `nom.qd.je` (sin www) | Solo por `http://` (GitHub Pages no ha emitido su certificado). Su página dice «NOM-COLAB». |
| Raspberry | Se reinició a las 22:53 (hora de la Pi) para poner el disco; NOM volvió solo con sus datos (3.4 MB). Sistema en la microSD de 64 GB. **El NVMe de 500 GB está conectado y el PCIe lo ve, pero el kernel `6.18.59-min` no trae NVMe, USB ni módulos: no se puede usar todavía** (`SERVIDOR-PROPIO.md`, «El disco nuevo»). |
| A55 | Fuera de la red de Tailscale desde la tarde del 17-nov. |
| Windows con Office real (`user@100.65.29.5`) | En línea. |

## 2. Ya corregido y comprobado

Todo esto está en los commits `e3d6d13` y `a22b6bb` y lo cubren pruebas automáticas que pasan:
`seguridad.py` **401 de 401**, `prueba_funciones.py` **203 de 203**, `prueba_ayuda.py` **24 de 24**, regresión rápida
(todo bien salvo las 6 marcas de `rastreo`, que son de la propia prueba: ver 4.3) y la app de Android compila.

| Defecto | Arreglo | Prueba |
|---|---|---|
| Un editor podía escribir filas de la bitácora a nombre de otra persona (falsificar su recibo y sus aportes) | Solo a nombre propio, o del propietario que sincroniza su teléfono; horas no futuras | `seguridad.py` 17.1 |
| El proyector (sin sesión) recibía comentarios, notas del orador y el Office original; su código servía hasta 8 h aunque a quien presentaba le quitaran el acceso | Sin eso; el código muere con el acceso; al mover o quitar el acceso se cierran sus pantallas | 17.2 |
| El enlace para ver llevaba el Office original (con sus comentarios y notas) | Sin el original | 17.3 |
| Con un código vigente se creaban alias sin fin | Tope por red (300/h) y total (3000/h) | 17.4 |
| Cualquiera en la misma red podía falsear su IP (cabeceras de reenvío creídas desde `fe80::`) | Solo en Azure | revisión de código |
| Errores de Java con nombres de clases, o 500 sin mensaje | 400 con mensaje propio | 17.5 |
| Los documentos abiertos nunca salían de la memoria (la Pi se llenaba con unos cuantos Office) | Salen tras 20 min sin uso, ya guardados | 17.6 |
| El id del servidor cambiaba al reiniciar: se perdían los alias guardados («volver a ser el mismo») | Se guarda (`<datos>-llaves/nodo.id`; en el teléfono, su carpeta privada) | 17.7 |
| Quien tenía abierto un documento movido a otro espacio seguía recibiéndolo | Se le cierra | revisión de código |
| La bitácora de sincronía y la de ramas cortaban en la primera racha vieja | Sin cortar | revisión de código |
| Una ecuación terminada en `\` tiraba el documento y la exportación | Se escribe tal cual | probado con node |
| El proyector y las miniaturas no veían las transiciones del panel nuevo | El proyector las juega como PowerPoint | `prueba_funciones` |
| «Dejar entrar sin cuenta: cada quien escribe un apodo» (falso desde la fase 13) | «Alias al azar» | a ojo |
| El texto oculto de Word se veía | Oculto, y se conserva; el elevado/bajado se dibuja | `comparar_real` w01 |
| Un Office de más de 14 MB no avisaba que no guarda su original | Aviso al importar | a ojo |
| `nomcolab local up` podía encender en la laptop el mismo túnel de la Pi (gente repartida entre dos servidores) | No lo enciende si la nube contesta desde otro equipo | probado en la laptop |
| Rojo de error sin contraste en tema oscuro | `--error` | a ojo |
| La prueba de la transición buscaba el menú viejo (la sección del lienzo no corría) | Usa el panel nuevo | `prueba_funciones` 203/203 |

**Documentación corregida** en todo el repositorio: LEEME, README, ayuda (qué se ve igual que Office, preguntas
frecuentes, 230 funciones), novedades de la 0.16.0, plan de construcción (casillas de lo ya hecho, cifras viejas,
erratas), servidor propio (`doas nomcolab-pi` solo con la ruta completa; aviso del túnel compartido; el disco nuevo),
dominio (`nom.qd.je` sin https), diseño, marca, medios (500 GB), Vercel (no se usa), relevos viejos marcados como
históricos. Las mediciones contra Office real se repitieron con todo lo sin publicar: **nada bajó** (tablas en
`RELEVO-ESTADO-Y-PENDIENTES.md` §2).

## 3. Hecho antes, pero todavía sin confirmar

Del 17-nov por la tarde y noche (commits `f3d3ea8` a `71243e9`): el panel de animaciones como PowerPoint, el panel
de transiciones, los vínculos y botones de acción, los videos, audios y sonidos al presentar, las gráficas y SmartArt
animados por partes, y la conversión a AV1/Opus (que todavía no se usa). Pasaron las pruebas automáticas de
esta noche, pero **falta abrir en PowerPoint real lo que arman** y probarlos en el teléfono (sección 4.1).

## 4. Lo que falta

### 4.1 Para publicar la 0.16.0 (en este orden)
1. **Regresión completa**: `pruebas/regresion.sh` (sin «rapido»: Office, ancho de celular, Firefox y el validador de
   Microsoft). Unos 60-80 min.
2. **La vuelta en PowerPoint real** (Windows): un archivo armado con el panel de animaciones (repetir, rebote,
   atenuar, ocultar, por párrafos, trayectoria propia), transiciones, vínculos, medios y sonidos → abrirlo en
   PowerPoint sin «reparar» + `pruebas/office/validar_openxml.py`. Herramientas: `RELEVO-17NOV-ANIMACIONES.md` §3.
3. **Teléfono**: el A55 (volver a vincular la depuración inalámbrica) o el emulador (`pruebas/reales/app_emulador.py`,
   `telefono_emulador.py`, `rama_emulador.py`): paneles con el dedo, medios en el WebView, el teléfono como servidor
   con el id nuevo (los alias se conservan al apagarlo y prenderlo).
4. **Subir la versión**: `Version.java` (APP 0.16.0, INTERFAZ 21), `nucleo.js` (`NOM.VERSION`, `NOM.INTERFAZ`), `sw.js`
   (`VERSION`), `android/app/build.gradle.kts` (0.16.0 / 31).
5. **Publicar todo** (decisión del dueño del 13-nov): commit y push a `main` (despliega a Azure), jar a la Pi
   (`nomcolab pi actualizar` y `NOM_BASE=https://www.nom.qd.je python3 pruebas/humo.py`), APK (`assembleRelease` +
   `android/publicar-apk.sh`), el A55, `herramientas/buscadores/avisar.py`, y la página de `nom-qd-je` con «NOM».

### 4.2 Lo que decide o hace el dueño
- **Kernel de la Pi con NVMe** (`CONFIG_BLK_DEV_NVME=y`; para discos USB, lo de USB) y pasar los datos al disco
  (`SERVIDOR-PROPIO.md`). Sin eso, los medios (fase 2 del plan) no tienen dónde vivir.
- **Certificado de `nom.qd.je`**: según la documentación de GitHub Pages, quitar y volver a poner el dominio en el
  repositorio `nom-qd-je` (cambio en GitHub: no se hizo sin permiso).
- Borrar la rama vieja del 28-oct que sigue en GitHub (rompe la regla de menciones).
- Borrar `~/NOM/tmp-17-11-26/` (3 GB, ya salvado lo útil) y, si quiere, `~/NOM-COLAB`.
- Revocar el token de Vercel que se pegó en una conversación.

### 4.3 Paridad con Office: lo que falta de verdad
La lista completa, por pestaña, está en `PARIDAD-OFFICE.md`. Lo más importante, con su causa medida:

- **Word** (renglones idénticos 77–100 %): división automática de palabras (w10: 34 %; el documento está en inglés y
  Word partió el español con reglas del inglés), escala de caracteres `w:w`, versalitas más anchas que las de
  Word, fuentes sin medidas de Office (Sitka, Lucida Sans Unicode), w06 con una hoja de menos, tabla flotante,
  tabuladores con puntos, notas al final, números de línea, EMF/WMF, campos que se actualizan, citas y
  combinación de correspondencia **de Word**, control de cambios con sus marcas, comparar, temas del documento,
  pegar con formato completo, `.doc`.
- **Excel** (100 % salvo una fila): la fila con texto ajustado que Excel recalcula al abrir (x01, «Letra», fila 24);
  crear tablas dinámicas y tablas, minigráficos, segmentaciones, herramientas de datos (quitar duplicados, texto en
  columnas, buscar objetivo, Solver, escenarios, tablas de datos), cálculo manual, definir nombres, rastrear
  precedentes, configurar página e impresión, comentarios con hilo, 143 de las funciones que lista el inventario
  (sobre todo financieras, estadísticas y de base de datos), Power Query, Power Pivot, VBA, `.xls`.
- **PowerPoint** (89.5–99.5 % de parecido): interlineado del texto por fuente (no es un corrimiento fijo: ±5 px por
  renglón), SmartArt sin dibujo (p11 89.5 %), calibrar las transiciones y Transformar (Morph) contra el video,
  insertar video y audio (plan de medios, fases 2-5), crear SmartArt, editar el patrón de diapositivas, secciones,
  ocultar diapositivas, presentaciones personalizadas, ensayar y grabar, efectos 3D, EMF/WMF, `.ppt`, video en la
  pantalla del proyector.

### 4.4 Lo que las pruebas no cubren (para no confundirse)
- `rastreo.py` marca 6 «MAL» que son de la prueba: toca «Bloquear» y después «Cortar» (que queda apagado), «Pegar
  especial» depende de lo copiado, y «Acción al hacer clic» abre una hoja que tapa la ruedita. Revisado a mano: la
  ruedita funciona.
- `comparar_real.py` compara texto: desde esta noche lee las MAYÚSCULAS de Word como se ven.
- No hay prueba automática de la salida de memoria con un documento abierto ni del cierre de canales al mover un
  documento (se revisaron en el código).

## 5. Cómo comprobar rápido que todo sigue bien

```sh
cd ~/NOM/4-CODIGO/app && mvn package                       # 13 pruebas de Java
# servidor de pruebas (ver RELEVO-ESTADO-Y-PENDIENTES.md §1) y luego:
cd ~/NOM/4-CODIGO && python3 pruebas/seguridad.py           # 401 de 401
pruebas/regresion.sh rapido                                 # o completa, sin «rapido»
~/.cache/nom-pruebas/venv/bin/python pruebas/office/comparar_real.py ~/.cache/nom-pruebas/office-real/archivos/w0*.docx
ssh pi5 /usr/local/sbin/nomcolab-pi estado                  # la Pi
```
