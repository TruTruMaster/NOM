# Publicación y cumplimiento: qué pide cada tienda, la ley y la accesibilidad

Revisión del 22 de noviembre de 2026 sobre la versión **0.22.0** (APK 39), para la entrega del 27-nov. Dice qué cumple
NOM hoy (con su prueba), qué falta, quién lo hace y qué decide el dueño. Las reglas de Google Play cambian seguido:
antes de enviar la app, volver a leer las páginas de la sección 10.

## 0. En corto

| # | Qué falta | Por qué | Quién | Tarea |
|---|---|---|---|---|
| 1 | **Correo del proyecto** (ponerlo en `CONTACTO` de `ayuda/generar.py` y volver a generar) | Google Play exige un correo de contacto en la ficha; la LFPDPPP pide un medio para ejercer los derechos ARCO (hoy es el aviso privado de GitHub, que obliga a tener cuenta de GitHub). Cómo tener uno propio: §7 | El equipo (panel de Cloudflare) | Fase 14 |
| 2 | ~~Domicilio del responsable~~ **hecho el 22-nov**: Mar Mediterráneo 227, Popotla, Miguel Hidalgo, 11400, CDMX | Artículo 15 de la LFPDPPP de 2025 | — | Fase 24 (M5) |
| 3 | **Revisión de un abogado** del aviso, los términos y el uso aceptable | Los mismos textos dicen que nadie los ha revisado; el equipo no tiene abogado (ver §7) | El equipo | Fase 24 (M5) |
| 4 | **Llave de firma propia y paquete `.aab`** | Google Play rechaza lo firmado con la llave de depuración y pide App Bundle para apps nuevas | Decisión del dueño (2.1) | Fase 24 (M5) |
| 5 | **Bloquear y reportar** a una persona o un contenido desde la app | Política de contenido generado por usuarios de Google Play (hay chat y @menciones) | Código (M2) | Fase 24 (M2) |
| 6 | **Cuenta de desarrollador y prueba cerrada** | Una cuenta personal nueva necesita 12 personas en prueba cerrada durante 14 días antes de pedir producción | El equipo | Fase 24 (M5) |
| 7 | **Registrarse como desarrollador de Android** | Desde 2027 (México incluido) un APK de un desarrollador no registrado deja de instalarse en teléfonos certificados, también el de GitHub (sección 3) | El equipo | — |
| 8 | **Accesibilidad a mano**: lector de pantalla (TalkBack, NVDA) y tamaño de letra del sistema en la app | Lo automático ya da 0 fallas (sección 5); lo demás solo se ve probando | El equipo y código | — |
| 9 | **Nombre y logotipo**: búsqueda en el IMPI antes de registrar «NOM-COLAB»; un logotipo nuevo sin parecido a Pac-Man | Evitar una oposición o una demanda de marca (sección 7) | El equipo | — |

Nada de esto impide la entrega del 27-nov (la app se distribuye por GitHub y la página está en línea); 1, 2 y 3 sí
conviene tenerlos antes de abrir la app a más personas, y 4 a 6 solo hacen falta para Google Play.

## 1. Lo que ya cumple (comprobado el 22-nov)

| Requisito | Cómo está | Prueba |
|---|---|---|
| Aviso de privacidad integral y simplificado (LFPDPPP 2025) | `/legal/privacidad` y `/legal/privacidad-corto`, con datos, finalidades, encargados (Cloudflare, Azure, GitHub), conservación, derechos ARCO, menores, cookies y autoridad (Secretaría Anticorrupción y Buen Gobierno) | En línea: https://www.nom.qd.je/legal/privacidad |
| Términos, uso aceptable, reportar contenido, accesibilidad, permisos de Android, licencias, almacenamiento | `/legal/` (fase 14), se aceptan al crear la cuenta y se vuelven a pedir si cambian | `prueba_ayuda.py` |
| Borrar la cuenta desde la app **y** desde la web | Menú de la cuenta → «Borrar mi cuenta…» (también en www.nom.qd.je); qué se borra y en cuánto tiempo: `ayuda/cuentas/borrar-la-cuenta` | `seguridad.py` |
| Descargar mis datos (derecho de acceso) | Ajustes → «Descargar los datos de mi cuenta» | `seguridad.py` |
| Sin publicidad, sin estadísticas de terceros, sin rastreadores | No hay SDK de terceros en la app (solo `desugar_jdk_libs`); la página no manda nada a otros servicios | `build.gradle.kts`, revisión del código |
| Nivel de API (Google Play pide **36** desde el 31-ago-2026) | `targetSdk = 37`, `compileSdk = 37`, `minSdk = 24` | `aapt2 dump badging` del APK publicado |
| Páginas de memoria de 16 KB (apps con código nativo) | El APK no trae bibliotecas nativas (`.so`): no aplica | `unzip -l NOM-COLAB.apk` (0 archivos `.so`) |
| Android 16 y 17: atrás predictivo, de borde a borde, pantallas grandes | `OnBackInvokedCallback`, márgenes con `WindowInsets` (barras y teclado), sin orientación fija | `MainActivity.java` |
| Permisos mínimos y explicados | Cámara y notificaciones se piden al usarlas; nada de ubicación, contactos, micrófono, SMS ni «todos los archivos» | `/legal/permisos-android` |
| La app no es de depuración | El APK publicado no lleva `debuggable` | `aapt2 dump badging` |
| Accesibilidad automática (WCAG 2.2 AA) | **0 fallas en 28 pantallas** (axe-core 4.10) después de los arreglos del 22-nov | `pruebas/interfaz/prueba_axe.py` 28/28 |
| Licencia del código y avisos de terceros | GPL-3.0 (`LICENSE`), `/legal/licencias` con cada biblioteca y fuente (el 22-nov se agregaron Mediabunny y las fuentes de las medidas de Office) | `/legal/licencias` |
| Página principal del repositorio | `README.md` en la raíz (antes se llamaba `LEEME.md` y GitHub no lo mostraba) | github.com/TruTruMaster/NOM-COLAB |

## 2. Google Play (si se publica ahí)

### 2.1 El paquete

- **Firma.** El APK se firma con la llave de depuración de la computadora (`CN=Android Debug`). Google Play **no
  acepta** paquetes firmados así. Para Play hace falta una llave de carga propia (`keytool -genkeypair`, guardada con
  respaldo y su contraseña fuera del repositorio) y la Firma de apps de Play. **Decisión del dueño**, porque un
  teléfono no instala una app encima de otra firmada con otra llave:
  - (a) Play con otro nombre de paquete (por ejemplo `mx.wack.nom.play`) y GitHub igual que hoy: conviven, pero son dos apps; o
  - (b) cambiar también la llave del APK de GitHub: quien ya la tiene la desinstala una vez (lo guardado en el
    teléfono se pierde si no se subió antes).
- **Formato.** Las apps nuevas se suben como **Android App Bundle**: `./gradlew bundleRelease` →
  `app/build/outputs/bundle/release/app-release.aab` (el APK de GitHub sigue igual).
- **Versión.** `versionCode` sube en cada envío (ya se hace en cada publicación).

### 2.2 Declaraciones de Play Console

- **Servicio en primer plano.** El teléfono como servidor usa el tipo `specialUse`. Play pide una declaración con
  la descripción («el teléfono comparte los documentos con otros equipos del mismo Wi-Fi mientras se ve la
  notificación Compartiendo») y un **video** con los pasos para encenderlo. Play revisa `specialUse` con más cuidado:
  evaluar si `connectedDevice` describe lo mismo (la app ya tiene los permisos de Wi-Fi que ese tipo pide).
- **Cámara**: solo para escanear el QR; se pide al tocar «Escanear QR». No hace falta declaración aparte.
- **Notificaciones**: se piden al empezar a compartir (Android 13+).

### 2.3 Seguridad de los datos (respuestas sugeridas)

Nombres como en Play Console en español (pueden cambiar un poco). Lo que viaja cifrado de extremo a extremo (espacios
cifrados) no cuenta como recopilado; lo que pasa por la red del salón tampoco (no llega al equipo de NOM).

| Tipo de dato | ¿Se recopila? | ¿Obligatorio? | Para qué |
|---|---|---|---|
| Información personal → ID de usuario (nombre de usuario) | Sí | Sí | Funcionalidad de la app; administración de la cuenta |
| Información personal → Otra información (escuela o grupo, preferencias) | Sí | No | Funcionalidad de la app |
| Mensajes → Otros mensajes en la app (chat del espacio, comentarios) | Sí | No | Funcionalidad de la app |
| Fotos y videos (los que se insertan en documentos o en el chat) | Sí | No | Funcionalidad de la app |
| Archivos y documentos | Sí | No | Funcionalidad de la app |
| Actividad en la app → Otro contenido generado por el usuario; Otras acciones (quién cambió qué y cuándo) | Sí | Sí | Funcionalidad de la app |
| IDs de dispositivo (resumen SHA-256 del identificador de Android, solo en memoria una hora) | Sí | Sí | Prevención de fraudes, seguridad y cumplimiento |
| Ubicación, contactos, salud, finanzas, navegación, registros de fallas, ID de publicidad | No | — | — |

- **¿Se comparte con terceros?** No (Cloudflare y Azure procesan datos por cuenta de NOM: son proveedores, no «compartir»).
- **¿Cifrado en tránsito?** Sí (la nube solo acepta HTTPS; `seguridad_red.xml`).
- **¿Se puede pedir que se borren?** Sí: en la app y en la web; enlace para la ficha:
  `https://www.nom.qd.je/ayuda/cuentas/borrar-la-cuenta` (explica qué se borra y que los respaldos se van en 14 días,
  30 en la copia del equipo). Quien ya no puede entrar a su cuenta lo pide por el contacto de la sección 8 del aviso.
- **Política de privacidad (URL):** `https://www.nom.qd.je/legal/privacidad`.

### 2.4 Contenido de la app

- **Acceso a la app:** dar a Google una cuenta de prueba (usuario y contraseña) con un espacio y documentos.
- **Anuncios:** no tiene. **Compras:** no tiene.
- **Clasificación de contenido (IARC):** herramienta de productividad; responder que las personas **interactúan y
  comparten contenido** (chat y documentos en espacios cerrados).
- **Público objetivo:** 13 años o más (no está dirigida a menores de 13, así no aplica la política de Familias). El
  aviso ya dice que un menor de edad usa NOM bajo la responsabilidad de sus padres o tutores.
- **Contenido generado por usuarios:** se aceptan los términos antes de publicar (sí, al crear la cuenta); reportar:
  sí en el chat (descarga lo reportado para enviarlo) y con el propietario del espacio; **bloquear a una persona: falta**
  (tarea 5). Mientras falte, Play puede rechazar la app.
- **Cifrado:** usa cifrado estándar (TLS y AES-256-GCM). Si se distribuye fuera de México, revisar las reglas de
  exportación del país (el cifrado estándar de uso masivo normalmente no pide trámite).

### 2.5 La ficha de la tienda

**Lista para copiar en `docs/FICHA-GOOGLE-PLAY.md`** (textos medidos, capturas de 1080 × 1920 y gráfico de 1024 × 500
en `marca/play/`).

Nombre (hasta 30 caracteres), descripción corta (80) y larga (4000), ícono de 512 × 512 PNG, gráfico de funciones de
1024 × 500, de 2 a 8 capturas de teléfono (las de `7-PRESENTACION/imagenes/app/*-cel.png` sirven de base), categoría
«Productividad», **correo de contacto** (obligatorio, tarea 1) y la URL de la política de privacidad. Si se cambia el
logotipo (sección 7), el ícono de la ficha y el de la app tienen que ser el mismo.

### 2.6 La cuenta de desarrollador

Cuota única de 25 USD y verificación de identidad. Una **cuenta personal creada después del 13-nov-2023** debe correr
una **prueba cerrada con al menos 12 personas durante 14 días seguidos** antes de pedir producción (una cuenta de
organización con D-U-N-S no lo pide, pero el D-U-N-S tarda hasta 30 días). Para el 27-nov solo es posible la prueba
interna o cerrada, no producción.

## 3. Fuera de Play: el APK de GitHub y la verificación de desarrolladores

Google anunció en 2025 que los teléfonos Android certificados solo instalarán apps de **desarrolladores registrados**,
vengan de donde vengan (también un APK de GitHub). Empezó el **30-sep-2026** en Brasil, Indonesia, Singapur y
Tailandia; **en 2027 llega a todos los países**. Para que el APK de GitHub se siga instalando cuando llegue a México:

- Registrarse en la consola de desarrolladores de Android y registrar el paquete `mx.wack.nom` con la huella de su
  llave (por eso importa la decisión de 2.1: se registra la llave con la que se va a seguir firmando).
- Hay una cuenta para **estudiantes y aficionados** sin cuota ni identificación oficial, limitada a 20 dispositivos
  autorizados; la cuenta completa no tiene ese límite.
- La página web no se ve afectada: en el navegador NOM funciona igual.

## 4. Apple: el «privacy manifest» no aplica

El *privacy manifest* (`PrivacyInfo.xcprivacy`) es un requisito de Apple para apps de iPhone y iPad en el App Store.
NOM **no tiene app de iPhone** (en iPhone se usa el navegador), así que no aplica. Lo equivalente en Android es lo de
2.3 (Seguridad de los datos) más la política de privacidad. Si algún día hay app de iPhone: manifiesto de privacidad,
«etiquetas nutricionales» de privacidad en App Store Connect, borrar la cuenta dentro de la app (ya existe) y Swift o
un WebView con las reglas de Apple.

## 5. Accesibilidad

- **Meta:** WCAG 2.2 nivel AA (`/legal/accesibilidad`).
- **Revisión automática (22-nov):** axe-core 4.10 con las reglas WCAG 2.0, 2.1 y 2.2 A y AA en 28 pantallas
  (entrada, avisos, ayuda, inicio, espacio, los cuatro editores, Ajustes, paleta de comandos y tema oscuro, en
  computadora y celular). Encontró **6 botones sin nombre** para lectores de pantalla (buscar e historial en los
  editores, nueva hoja, nueva diapositiva y las dos vistas del texto) y **iniciales con poco contraste**: 5 de los 8
  colores de persona daban menos de 4.5 a 1 con letra blanca (el naranja, 2.5 a 1). Arreglados: ahora cada botón
  tiene nombre y las iniciales y la etiqueta del cursor escogen letra blanca, la tinta oscura o el mismo tono un poco
  más oscuro, lo que dé al menos 4.5 a 1. Después: **28 de 28 sin fallas** (`pruebas/interfaz/prueba_axe.py`).
- **Lo que lo automático no ve** (pendiente, a mano): recorrer la app con TalkBack (A55) y NVDA o Narrador
  (Windows); el Accessibility Scanner de Google en el A55; los lienzos (presentación y diagrama) todavía no
  describen cada figura a un lector de pantalla (ya lo dice la declaración).
- **Tamaño de letra en la app de Android:** la app no sigue el tamaño de letra del sistema (`setTextZoom(100)`) y su
  ajuste propio llega a 140 %; WCAG 1.4.4 pide poder llegar a 200 %. Propuesta: empezar con la escala del sistema
  y dejar llegar a 200 %, probándolo en el A55 para que la interfaz no se rompa.

## 6. Compatibilidad

| Dónde | Mínimo | Probado |
|---|---|---|
| Android (app) | Android 7.0 (API 24) con **Android System WebView 111 o más reciente** (se actualiza en Play Store). Android 8 para que el teléfono comparta | Emulador (WebView 124) y Samsung A55 (WebView 157) |
| Chrome, Edge, Samsung Internet | 111 o más reciente | Chromium de pruebas (regresión completa) |
| Firefox | 121 o más reciente | Firefox de pruebas (`funciones_firefox` en la regresión de Windows) |
| Safari (Mac, iPhone, iPad) | 16.2 o más reciente | Sin probar en un equipo de Apple |
| Pantalla | Desde 320 px de ancho | `prueba_funciones` en ancho de celular |
| Archivos | `.docx`, `.xlsx`, `.pptx` (también con contraseña de Office), `.nom`, `.md`, `.csv`, PDF al descargar | Contra Office 2024 real (`PARIDAD-OFFICE.md`) |
| No abre | `.doc`, `.xls`, `.ppt`, OpenDocument, `.rtf`, ni archivos con macros | Fase 24 (M3) |

Los mínimos salen de lo que usa la interfaz: `color-mix()` en CSS (Chrome 111, Firefox 113, Safari 16.2), `:has()`
(Firefox 121, Safari 15.4) y unidades `dvh`. Convertir video en el equipo usa WebCodecs: mejor en Chrome o Edge.

## 7. Ley mexicana, marcas y el logotipo

- **Datos personales (LFPDPPP, DOF 20-mar-2025):** el aviso cubre lo que pide la ley, también el domicilio del
  responsable (desde el 22-nov); falta un **medio de contacto** propio (tarea 1).
- **Correo propio:** una dirección siempre lleva el dominio completo («@nom» solo no existe). Hoy se puede tener
  **`contacto@nomcolab.dpdns.org`** gratis con *Email Routing* de Cloudflare, porque esa zona ya está en la cuenta de
  Cloudflare del equipo: Cloudflare pone los registros y reenvía lo que llegue a un Gmail del equipo (hay que
  confirmar ese Gmail con el correo que manda Cloudflare). Solo recibe: para contestar *como* esa dirección se usa
  «Enviar como» de Gmail. **`contacto@nom.qd.je`** se vería mejor, pero `nom.qd.je` no puede ser zona de Cloudflare
  (sección del dominio en `DOMINIO-Y-BUSCADORES.md`) y no está comprobado que el panel de DigitalPlat acepte registros
  MX; si los acepta, haría falta otro servicio de reenvío. Con la dirección lista: `CONTACTO["correo"]` en
  `ayuda/generar.py`, volver a generar y publicar.
- **Sin abogado:** los textos ya lo dicen con franqueza («todavía no lo ha revisado un abogado»), y eso no incumple
  la ley; lo que la ley pide está en el aviso. Si el proyecto sigue después de la escuela, conviene que un abogado o
  una clínica jurídica universitaria los revise antes de abrirlo a más personas. Transferencias: no hay (si un día las hay, el
  artículo 35 pide una cláusula para aceptarlas o no). La autoridad ya es la Secretaría Anticorrupción y Buen
  Gobierno (el INAI desapareció en 2025).
- **Marcas de otros:** los términos y las licencias dicen que Word, Excel, PowerPoint y Office son de Microsoft,
  Canva de Canva Pty Ltd, Android de Google y «QR Code» de DENSO WAVE, y que NOM no está afiliado. Abrir y guardar
  Open XML está permitido (es norma ISO/IEC 29500 y Microsoft la cubre con su Open Specification Promise). Las fuentes
  que imitan las medidas de Office no copian el dibujo de ninguna fuente de Microsoft (`fonts/LICENCIAS-FUENTES.txt`).
- **El nombre:** antes de registrar «NOM-COLAB» como marca, buscarlo en el IMPI (MARCia) en las clases 9 (software)
  y 42 (servicios en línea). Ojo: «NOM» también es la sigla de *Norma Oficial Mexicana*, que es un signo oficial; la
  marca compuesta «NOM-COLAB» se distingue mejor que «NOM» sola.
- **El logotipo nuevo (propuesta de caritas que «comen» el documento):** la idea (nom nom = comer, tres colores =
  colaborar) es buena, pero tres círculos con la boca en cuña se parecen mucho al personaje de **Pac-Man**, que Bandai
  Namco tiene registrado como marca en decenas de países, también para software, y que defiende. Para usarla sin
  riesgo, las caritas tienen que ser originales (otra silueta, sin la cuña ni los puntos que se comen, sin amarillo)
  y el logotipo debe seguir las reglas de `DISENO-UI.md` (plano, sin brillos ni volumen) y leerse a 16 px.

## 8. Licencias de terceros

Todo es compatible con la GPL-3.0: Spring Boot (Apache-2.0), H2 (MPL-2.0/EPL-1.0), controlador de SQL Server (MIT),
qrcode-generator (MIT), jsQR (Apache-2.0), Lucide (ISC), Mediabunny (MPL-2.0: su código fuente se indica en
`/legal/licencias`), `desugar_jdk_libs` (GPL-2.0 con excepción de enlace) y las fuentes (SIL OFL 1.1, con nombres
propios en las modificadas). Cada archivo de terceros guarda su licencia junto a él (`js/vendor/LICENCIAS.md`,
`fonts/LICENCIAS-FUENTES.txt`, `img/LICENSE-lucide.txt`).

## 9. Herramientas para revisar antes de publicar

| Herramienta | Qué dice | Dónde |
|---|---|---|
| Informe previo al lanzamiento de Play Console | Corre la app en teléfonos reales: fallas, accesibilidad, capturas | Play Console, al subir a una prueba |
| Play Policy Insights (Android Studio) y Android Lint | Avisos de políticas de Play y errores comunes en el código de Android | Android Studio |
| Accessibility Scanner | Botones chicos, contraste y etiquetas en el teléfono | App de Google en el A55 |
| axe-core | Reglas WCAG en cada pantalla | `pruebas/interfaz/prueba_axe.py` |
| Lighthouse | Rendimiento, accesibilidad, buenas prácticas y buscadores de la página | Chrome (herramientas de desarrollo) o `npx lighthouse` |
| Estándares de la comunidad de GitHub | README, licencia, guía para colaborar, código de conducta, plantillas de reportes | Insights → Community del repositorio |
| REUSE (FSFE) | Que cada archivo diga su licencia | `reuse lint` |
| PWABuilder | Si la página instalable está lista para tiendas (manifiesto, service worker) | pwabuilder.com |

## 10. Fuentes

- Nivel de API de Google Play: https://developer.android.com/google/play/requirements/target-sdk
- Prueba cerrada de cuentas personales nuevas: https://support.google.com/googleplay/android-developer/answer/14151465
- Borrar la cuenta (requisito de Play): https://support.google.com/googleplay/android-developer/answer/13327111
- Servicios en primer plano: https://support.google.com/googleplay/android-developer/answer/13392821
- Contenido generado por usuarios: https://support.google.com/googleplay/android-developer/answer/9876937
- Verificación de desarrolladores de Android: https://developer.android.com/developer-verification/guides y
  https://android-developers.googleblog.com/2026/03/android-developer-verification-rolling-out-to-all-developers.html
- Firmar la app (Play no acepta la llave de depuración): https://developer.android.com/studio/publish/app-signing
- LFPDPPP vigente: https://www.diputados.gob.mx/LeyesBiblio/pdf/LFPDPPP.pdf
- Marca PAC-MAN (ejemplo de registro): https://ised-isde.canada.ca/cipo/trademark-search/1014978
