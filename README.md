# NOM

Escribe, calcula, presenta y dibuja en equipo, al mismo tiempo, con o sin Internet. Abre y guarda
Word, Excel y PowerPoint. Gratis, sin anuncios y de código abierto (**GPL-3.0**), en el navegador y
en Android. Hasta **50 personas conectadas a la vez** por espacio.

- **En Internet:** **https://www.nom.qd.je**.
- **App de Android:** la última versión siempre en
  [Releases](https://github.com/TruTruMaster/NOM/releases/latest) (`NOM.apk`). Cada
  versión trae su SHA-256 y la huella del certificado de firma, para comprobar que es la original.
- **Sin Internet:** una computadora (o un teléfono) del salón comparte en el mismo Wi-Fi.

WACK

> **Si lo bajaste de GitHub:** las carpetas `2-PAGINA-Servidor/` y `3-APP-Android/` no vienen (son
> programas ya compilados). El APK está en Releases; el servidor se compila con
> `cd 4-CODIGO/app && mvn package` y se arranca con `java -jar target/nom.jar` (ver
> `4-CODIGO/README.md`).

### In English

NOM is a free (GPL-3.0) real-time collaborative office suite: documents, spreadsheets,
presentations and diagrams that several people edit at once, online or on a classroom Wi-Fi without
Internet (a laptop or an Android phone acts as the server). It opens and saves Word, Excel and
PowerPoint files (including password-protected ones) without third-party libraries. Server: Spring
Boot (Java 17); client: plain JavaScript; Android: a WebView app. Live at https://www.nom.qd.je.
Documentation and the code comments are in Spanish.

## Qué hay en esta carpeta

| Carpeta | Qué es | Qué hacer |
|---|---|---|
| **`ENTREGA/`** | **Los PDF para entregar:** el plan de pruebas `5IM7_WACK_PP.pdf`, la presentación `5IM7_WACK_Presentacion.pdf` y la especificación de casos de uso `5IM7_WACK_Casos_de_Uso` (.docx y .pdf) | Entregarlos |
| `1-ENTREGA-Diagramas-de-Actividades/` | **PDF para entregar** (y el mismo en .docx), las 7 imágenes sueltas y los editables de draw.io | Entregar `Diagrama de Actividades NOM.pdf` |
| `2-PAGINA-Servidor/` | La página completa (servidor + interfaz), **con Java incluido** | Windows: doble clic en `INICIAR-Windows.bat` · Linux: `./INICIAR-Linux.sh` |
| `3-APP-Android/` | `NOM.apk` (Android 7 en adelante) | Pasarla al celular e instalarla |
| `4-CODIGO/` | Todo el código fuente, diagramas UML y documentación | Para revisar o seguir desarrollando |
| `6-PLAN-DE-PRUEBAS/` | Fuentes del plan de pruebas: `fuentes/plan.html`, `fuentes/imprimir.py` (lo genera en `ENTREGA/5IM7_WACK_PP.pdf`), `fuentes/evidencias.py` (toma las capturas de evidencia), `fuentes/casos_integracion.py`, `fuentes/casos_permisos.py`, `fuentes/prueba_limites.py` y `registros/` con la salida de cada suite | Solo si se cambia el plan |
| `7-PRESENTACION/` | Fuentes de la presentación del proyecto y de la especificación de casos de uso: `fuentes/casos.py` (los textos), `fuentes/generar_casos.py` (arma el Word sobre el formato vacío), `fuentes/presentacion.py` (arma el PDF de 20 diapositivas) y `fuentes/capturas.py` (toma las capturas de la app que usa la presentación); los dos dejan su resultado en `ENTREGA/` | Solo si se cambia la presentación o un caso de uso |
| `8-ENSAYO/` | Fuentes del ensayo individual «NOM como sistema distribuido»: `fuentes/generar.py` arma el Word y el PDF de las dos versiones en `ENTREGA/` (`Ensayo_Rodriguez_Esparza`: 4 páginas con portada; `..._portada_aparte`: portada + 4) y `fuentes/componentes.puml` es el diagrama de componentes (PlantUML) | Solo si se cambia el ensayo |
| `LOG.md` | Diario de todo lo que se hizo, día por día (es historia: lo viejo puede ya no ser cierto; lo vigente está en `CAMBIOS.md` y en la ayuda) | — |
| `CAMBIOS.md` | Las versiones y lo que cambió en cada una, en corto | — |

## En Internet: la Raspberry Pi

**https://www.nom.qd.je** — la página completa, corriendo en una Raspberry Pi 5 que está siempre
prendida y publicada gratis con un túnel de Cloudflare (sin abrir puertos). La app Android la usa
como "En Internet". Cada cuenta tiene un límite de almacenamiento a la vista (200 MB) y solo guarda
documentos de NOM, comprimidos y cifrados. Respaldo diario en la Pi (14 días) y una copia diaria en
otra computadora del equipo (30 días). Cómo se administra, los respaldos y la instalación:
`4-CODIGO/docs/SERVIDOR-PROPIO.md`.

La Raspberry y un servidor en el salón son independientes (cada uno tiene sus cuentas y documentos).

## Usarlo en 3 pasos

1. **En la computadora que comparte** (la del profesor o un compañero): abrir
   `2-PAGINA-Servidor` y ejecutar `INICIAR-Windows.bat` (o `INICIAR-Linux.sh`).
   No hay que instalar nada: el Java ya viene dentro. La página se abre sola en
   el navegador. **No cierres la ventana negra** mientras trabajen.
   - Si Windows pregunta por el **Firewall**, elige *Permitir acceso* en redes
     privadas; si no, los demás equipos no podrán entrar.
   - La ventana muestra la dirección para los demás, por ejemplo
     `Desde la red: http://192.168.1.20:8080`.
2. **En las demás computadoras** de la misma red: abrir esa dirección en el
   navegador. También funciona en el navegador del celular.
3. **En Android:** instalar `NOM.apk` (permite "instalar apps de origen
   desconocido" si lo pide). Con el celular en el **mismo Wi-Fi**, la app
   encuentra el servidor con **Buscar en este Wi-Fi** (o con el QR); si hay varios, eliges uno. Si
   la red bloquea el descubrimiento, escribe la dirección (`192.168.1.20:8080`).

### Sin computadora: el teléfono es el servidor
En Android 8 o más reciente, en la pantalla **«¿Dónde trabaja tu equipo?»** → **Compartir desde este
teléfono** (o en *Este dispositivo* → **Trabajar sin Internet con este teléfono**). Ese teléfono
guarda los espacios y los documentos (no tiene cuentas con contraseña) y los demás entran:
- con la app, en el mismo Wi-Fi: **Buscar en este Wi-Fi** (aparece el nombre del teléfono) o
  escaneando el QR que muestra;
- con el navegador (computadora u otro celular): la dirección que muestra la app, por ejemplo
  `http://192.168.0.232:8080` (botón **Compartir** para mandarla por WhatsApp o correo).

Quien llega **sin Internet** entra con un **alias al azar** (no escribe nada); quien **sí tiene
Internet** entra con su cuenta de la nube y la app le da un pase firmado de 10 minutos para el
teléfono. Si dos personas cambian el mismo documento mientras una estaba sin conexión, lo de esa
persona queda en una **rama** y el original no se toca: después se compara y se fusiona.

Sin router: activa el **punto de acceso** de ese teléfono y conecta a los demás a esa red.
Mientras está encendido aparece una notificación con la dirección y el botón **Detener**; gasta
algo más de batería. Se apaga ahí o en *Este dispositivo → Trabajar sin Internet con este teléfono*.

### Sin cuenta
En la pantalla de inicio está el botón **"Usar sin cuenta"**. En la app Android también aparece
cuando no hay servidor, así que la app funciona sola en el teléfono. Sin cuenta se pueden crear,
editar, importar y exportar documentos en los 4 modos, pero:
- se guardan **solo en ese navegador o teléfono**;
- **si se borran la caché o los datos del navegador (o de la app), se pierden**, salvo que se hayan
  descargado (menú **… → Descargar copia (.nom)**);
- no se comparten ni tienen historial de versiones.

Si después se crea una cuenta, esos documentos aparecen en Inicio como **"En este equipo (sin subir)"**
y con **… → Subir a un espacio** quedan guardados en el servidor.

### Con cuenta
**Crear cuenta → Nuevo espacio → Nuevo documento.** Al crear la cuenta sale un **código de
recuperación**: guárdalo (sirve si se olvida la contraseña).
Para que otras personas entren a tu espacio usa **Invitar**: **Crear código** da un código de 8
letras y un QR (se puede **dejar entrar sin cuenta**: cada quien entra con un alias al azar, sin
escribir nada; dura 1 hora, 1 día, 1 semana o 30 días, y se anula cuando quieras); o escribe el
nombre de usuario de alguien que ya tiene cuenta.
Quien tenga el código lo escribe en **Tengo un código de invitación** o escanea el QR con la cámara.

Cada espacio tiene su **chat** (con imágenes y archivos de hasta 2 MB). Al crear un espacio se puede
marcar **Cifrado de extremo a extremo**: el nombre, los documentos y el chat se cifran en el equipo
de cada miembro y ni el servidor los puede leer (no se puede quitar después y no hay búsqueda dentro
del contenido).

Los datos se guardan **cifrados** en `2-PAGINA-Servidor/datos/`: usuarios y sesiones en la base de
datos H2 (`nom-db.mv.db`); espacios, documentos y versiones en JSON. La llave está aparte, en
`2-PAGINA-Servidor/datos-llaves/reposo.key`: **sin ella no se abre nada**, así que respáldala en
otro lugar. Para empezar desde cero, borra las dos carpetas. **No hay administrador:** cada cuenta controla solo sus espacios y documentos. El
propietario de un espacio decide quién entra, cambia a cada miembro entre *Puede editar* y
*Solo lectura*, o le quita el acceso.

## Licencia

El código es software libre bajo la **GPL-3.0** (archivo `LICENSE`). Las bibliotecas y fuentes de
terceros conservan sus licencias (ver `4-CODIGO/README.md`, sección Licencia).
