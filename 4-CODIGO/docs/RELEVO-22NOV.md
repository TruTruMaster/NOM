# Relevo del 22 de noviembre de 2026

**Leer este primero.** Sigue a `RELEVO-21NOV.md`. Al empezar estaba publicada la **0.22.0** (interfaz 27, APK 39).
Esta sesión no fue de paridad con Office: el dueño pidió revisar todo para la entrega del **viernes 27-nov**
(presentación, especificación de casos de uso y plan de pruebas), lo que pide Google Play y la ley para publicar sin
rechazos ni problemas, la accesibilidad, la compatibilidad, y su idea de logotipo nuevo.

## 1. Hecho

| Qué | Dónde | Comprobado |
|---|---|---|
| **Avance por módulo a partir del plan** | `PLAN-CONSTRUCCION.md`: fase 24 (lo más importante que falta de Office y lo que pide Google Play, cada tarea con su módulo) y la sección «Avance por módulo» (qué fases cuenta cada módulo). Dos casillas viejas marcadas con su prueba (puntos guía y guiones de Word) | M1 100 %, M2 93 %, M3 79 %, M4 82 %, M5 77 %, **M6 Digitalización 0 % (nuevo, planeado)**; total 166 de 211 = 79 % |
| **Presentación** (`ENTREGA/5IM7_WACK_Presentacion.pdf`, 20 diapositivas, las mismas) | `7-PRESENTACION/fuentes/presentacion.py` ahora **lee el avance del plan** (ya no está escrito a mano) y la versión de `build.gradle.kts` y el número de funciones de la ayuda | Diapositiva 11: mapa de navegación completo (recuperar acceso, cuenta y ajustes, este equipo, teléfono servidor, proyector) con color por módulo y lo planeado punteado (texto fuente, digitalizar). 12: tabla de los seis módulos con lo que falta (lo más importante de Office y lo de después del 27-nov). 13: casos por módulo con el M6. 8: 424 funciones. 10 y 19: www.nom.qd.je. Capturas nuevas de la 0.22.0 sin el consejo de la ruedita |
| **Casos de uso** (`ENTREGA/5IM7_WACK_Casos_de_Uso.docx` y `.pdf`) | `casos.py`: módulo M6 «Digitalización (planeado)» con **CU-31 Digitalizar una hoja impresa**; `generar_casos.py`: el último flujo va pegado a la postcondición (ya no queda la fila POSTCONDICION sola en otra página) | 13 casos, 19 páginas; **el PDF lo hizo Word 2024** (`pruebas/office/windows/word_a_pdf.py`, PDF con etiquetas de accesibilidad) |
| **Plan de pruebas, revisión 3** (`ENTREGA/5IM7_WACK_PP.pdf`, 28 páginas) | `6-PLAN-DE-PRUEBAS/fuentes/plan.html` §6.8 «Seguimiento del 22 de noviembre»; DEF-50 cerrado; DEF-55 y DEF-56 nuevos y cerrados | Repetidas sobre la 0.22.0: unitarias 26/26, integración 14/14, límites 18/18, permisos 30/30 (`registros/22-nov/`); VL-18 leía mal la respuesta (el servidor contesta 413 y corta el envío): la prueba ahora la lee como el navegador |
| **Accesibilidad** (`5f7ddb2`) | 6 botones de los editores sin nombre (DEF-55); iniciales y etiqueta del cursor con letra que da al menos 4.5 a 1 sobre cualquier color de persona (`NOM.sobreColor`, DEF-56) | **`pruebas/interfaz/prueba_axe.py` nueva** (axe-core 4.10, WCAG 2.2 A y AA, 28 pantallas): 28/28; `prueba_funciones` 204/204; `prueba_e2e` 33/33 |
| **Licencias y README** (`c010998`) | `/legal/licencias` con Mediabunny (MPL-2.0) y las fuentes de las medidas de Office; `LEEME.md` → `README.md` (GitHub no mostraba la portada del repositorio) | `generar.py` |
| **Publicación y cumplimiento** | `docs/PUBLICACION-Y-CUMPLIMIENTO.md`: qué cumple, qué falta y de quién (Google Play, verificación de desarrolladores de Android, «privacy manifest» de Apple, LFPDPPP, marcas, accesibilidad, compatibilidad, licencias, herramientas) | APK publicado revisado con `aapt2` y `apksigner` |
| Logotipo | Propuesta en `~/NOM/trabajo/22nov/logo/` (fuera de git, sin aplicar): `simbolo-v3.svg`, `logo-horizontal-v3*.svg`, comparación `v3.png` | Ver §3 |
| Lo de seguridad que vería Google Play | Solo en el plan privado (`Auditoria-Seguridad-PLAN.md` §K) | — |

## 1b. Publicación de la 0.22.1 (22-nov, 13:41): hecha

- Lleva lo de accesibilidad (`5f7ddb2`), el aviso de licencias completo (`c010998`) y el **domicilio del responsable**
  en el aviso de privacidad (`c1942aa`: Mar Mediterráneo 227, Popotla, Miguel Hidalgo, 11400; la versión del aviso
  sigue en 2026-11-13: nadie tiene que aceptarlo de nuevo). Commit de versión `8434b1b` (interfaz 27, APK 40).
- **Regresión completa de Windows** (`--carriles 10`, 1396 s): 39 de 41. Las dos que fallaron, solas en Windows: bien
  (`1e15717`): `limites` era de la prueba (VL-18: en Windows el corte del envío da WinError 10053 y se perdía el 413;
  ahora se lee en otro hilo) y `funciones_movil` 206/207 bajo carga (Beto tardó en ver la sugerencia aceptada; sola,
  207/207, igual que en la 0.22.0). De paso: `regresion_windows.py` bajaba la carpeta «numlineas» en vez de la de la
  corrida (ahora solo las que tienen fecha).
- Pi: `nomcolab pi actualizar` (respaldo `datos-2026-11-22_134146`), `/api/red/info` 0.22.1 / 27, `humo.py` contra
  www.nom.qd.je **34/34**. GitHub hasta `1e15717`. APK 40: release `v0.22.1` (SHA-256 `dc8a4aed…`, misma llave
  `2ef146f6…`). IndexNow 200 (76 direcciones). **Falta el A55** (no estaba conectado).
- Presentación regenerada: dice 0.22.1 y M5 79 % (167 de 212 tareas en total).

## 1c. Tarde: DNS de nom.qd.je por la API de DigitalPlat, correo y 0.22.2 (22-nov, 16:33)

- **API de DigitalPlat** (`https://domain-api.digitalplat.org/api/v1`, llave en `~/.config/nomcolab/digitalplat`, fuera
  del repositorio; Cloudflare bloquea el cliente de Python por su firma: mandar `User-Agent: nom-colab-dns/1.0`). La
  documentación (70 rutas) la guardó el dueño; lo útil: `domains/{d}/dns/records` (GET/POST/PATCH/DELETE, con
  `Idempotency-Key`), `dns/export`, `dns/import/preview`. **No hay ruta para la redirección**: está solo en el panel.
- Zona de `nom.qd.je` (respaldo de antes en `~/NOM/trabajo/22nov/dns/zona-antes.txt`): se quitaron los 4 A y 4 AAAA de
  GitHub Pages y se puso `@ A 23.238.4.245` (la entrada de la redirección de DigitalPlat); **`www` no se tocó** (el panel
  sugiere `www CNAME pages-origin.digitalplat.org`: eso tumbaría la página principal). Correo: `@ MX 10/20
  mx{1,2}.improvmx.com.` y `@ TXT "v=spf1 include:spf.improvmx.com ~all"` (falta que el dueño cree la cuenta de ImprovMX).
- La redirección **sigue esperando** («Domain redirect is waiting for HTTPS setup», certificado de origen de Cloudflare en
  la entrada): lo más probable es que «Include www» esté marcado. Al activarse: comprobar 301 con ruta y `?c=`, apagar
  GitHub Pages del repositorio `nom-qd-je`, `SITIO = "nom.qd.je"` en `presentacion.py`, y actualizar `DOMINIO-Y-BUSCADORES.md`.
- **0.22.2 publicada** (16:33): atajos `/app` (baja el APK) y `/codigo` (`Atajos.java` en el núcleo; el servidor y el
  teléfono igual). Regresión de Windows **completa sin fallas** (1404 s), Pi `humo.py` 36/36, GitHub, APK 41
  (`932bcddc…`, misma llave), IndexNow. Falta el A55.
- Google Play: `docs/FICHA-GOOGLE-PLAY.md` y `marca/play/` (capturas de 1080 × 1920 y gráfico de 1024 × 500,
  `herramientas/play/capturas_play.py`). Caché de Cloudflare: ya activa (no hace falta regla). Turnstile: después del 27-nov.

## 2. Lo que decide o hace el dueño (de `PUBLICACION-Y-CUMPLIMIENTO.md` §0)

1. Crear el **correo del proyecto** (`contacto@nomcolab.dpdns.org` con Email Routing de Cloudflare, ver
`PUBLICACION-Y-CUMPLIMIENTO.md` §7) y ponerlo en `CONTACTO` (`ayuda/generar.py`); 2. ~~domicilio~~: hecho el 22-nov
(Mar Mediterráneo 227, Popotla, Miguel Hidalgo, 11400; lo eligió el dueño, sin nombrar a la escuela); 3. **revisión
de un abogado** (no tienen: los textos lo dicen); 4. **llave de firma propia** (opción a o b de §2.1) si se va a Google Play;
5. cuenta de desarrollador y prueba cerrada de 12 personas por 14 días; 6. registrarse en la verificación de
desarrolladores de Android antes de 2027; 7. el logotipo (§3); 8. probar con TalkBack/NVDA a mano.

## 3. Logotipo (la idea del dueño: caritas que «comen» el documento)

- Sus dos borradores (`~/Downloads/nuevo logo.png`, `NOM COLAB PRUEBA 1.png`): la idea es buena (nom nom = comer; tres
  colores = colaborar; el documento = lo que hace NOM), pero el azul y el naranja son círculos con la boca en cuña,
  la silueta de **Pac-Man** (marca de Bandai Namco, que la defiende, también en software); el verde es otra especie
  (un grumo con cejas enojadas que parece virus); el estilo 3D con brillo va contra `DISENO-UI.md` §1 y §14; las
  letras N-O-M en cuadritos no se leen a 48 px; el PNG no es vectorial.
- Propuesta v3: hoja plana con la esquina doblada y tres «comelones» originales (cuadros redondeados, dos ojos, boca
  abierta con dientes mordiendo la orilla), en índigo `#6767E8`, naranja `#F06A3D` y verde azulado `#14A38B` sobre el
  mosaico de la marca `#2B2952`; se distingue a 48 px. **No se aplicó**: cambiarlo es cambiar el ícono de la app
  (APK), la página, `marca/`, el manual y la presentación. Si el dueño lo aprueba, hacerlo completo y de una vez.

## 4. Cómo seguir

- Servidor de desarrollo: 8096 con `trabajo/srv/nom-dev.jar` = el jar de la 0.22.0 (el anterior quedó como
  `nom-dev-21nov-1654.jar`).
- Regenerar los entregables: `~/.cache/nom-pruebas/venv/bin/python 7-PRESENTACION/fuentes/presentacion.py` (lee el plan:
  al marcar casillas basta con volver a correrlo; revisar que el texto de `FALTA` siga siendo cierto),
  `python3 7-PRESENTACION/fuentes/generar_casos.py` y el PDF con Word (`word_a_pdf.py`), y
  `~/.cache/nom-pruebas/venv/bin/python 6-PLAN-DE-PRUEBAS/fuentes/imprimir.py`.
- Antes de cada publicación: `prueba_axe.py` (accesibilidad) además de lo de siempre.
- Después del 27-nov: el orden de la fase 20 (`PLAN-CONSTRUCCION.md`, «Orden recomendado») y la fase 24.
