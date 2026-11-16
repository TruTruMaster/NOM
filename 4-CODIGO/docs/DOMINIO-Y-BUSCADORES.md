# El nombre corto nom.qd.je y los buscadores

Escrito el 16 de noviembre de 2026. El sitio tiene dos nombres que llevan al mismo servidor (la
Raspberry, con sus cuentas y su base de datos):

| Nombre | Para qué |
|---|---|
| `https://nom.qd.je` | El nombre corto: el que se comparte y el que se ve en Google, Bing y los demás. |
| `https://nomcolab.dpdns.org` | El de siempre. **No se apaga:** la app de Android le habla a este nombre, los pases van firmados con él y quien ya guardó cosas en el navegador con este nombre las sigue teniendo ahí. |

No se redirige un nombre al otro, porque el navegador guarda la sesión y los documentos de cada
nombre por separado. Una redirección dejaría a la gente sin sesión y sin lo que tenía guardado sin
Internet. Para que los buscadores no los cuenten como dos sitios, cada página dice cuál es su dirección
principal (`<link rel="canonical">`). Esa dirección sale de `NOM_CANONICO` en `/etc/nomcolab.conf`.

## 1. Por qué Cloudflare no acepta nom.qd.je

`dpdns.org` está en la [lista pública de sufijos](https://publicsuffix.org/list/public_suffix_list.dat)
(como `.com`), así que para Cloudflare `nomcolab.dpdns.org` es un dominio completo y se puede agregar
como zona gratis. `qd.je` no está en esa lista: para Cloudflare, `nom.qd.je` es un subdominio de
`qd.je`, y agregar subdominios como zona solo se puede con el plan Enterprise.

La salida es **Cloudflare for SaaS**, que sí viene en el plan gratis (hasta 100 nombres). La zona
`nomcolab.dpdns.org` acepta tráfico de otro nombre (`nom.qd.je`) y lo manda al mismo túnel de la
Raspberry. Así sigue el método de siempre: sin abrir puertos y con certificado HTTPS de Cloudflare.

```
navegador ──https://nom.qd.je──▶ Cloudflare (nombre personalizado de la zona nomcolab.dpdns.org)
                                   └─▶ origen de respaldo origen.nomcolab.dpdns.org
                                         └─▶ túnel «nomcolab» ──▶ Raspberry :8080
```

## 2. Lo que ya está hecho

- **DNS en Cloudflare:** `origen.nomcolab.dpdns.org` → el túnel `nomcolab` (CNAME con proxy). Será el
  «origen de respaldo».
- **Servidor** (`Buscadores.java`, `PaginasApi.java`): la portada y cada página de ayuda y legal
  llevan su dirección principal, lo que muestran las redes al compartir el enlace (Open Graph y
  tarjeta de X, con `img/compartir.png` de 1200×630) y datos de schema.org. En la portada son WebSite
  con los nombres con que la buscan (NOM, NOMCOLAB, NOM COLAB), Organization con el logo y
  WebApplication. En la ayuda son WebPage y la ruta de migas. `robots.txt` y `sitemap.xml` se arman
  con la dirección principal. `/unirse` y `/proyectar` salen con `noindex`.
- **Página:** un título que describe la app; la portada de `index.html` trae texto que leen los
  buscadores que no corren JavaScript (la app la quita al arrancar); el manifiesto lleva descripción y
  categorías.
- **La app reconoce nom.qd.je** como la nube (`acceso.js`, `vista-cuentas.js`, `vista-espacios.js`) y
  `pases/nubes.json` la incluye con la misma llave.
- **Raspberry:** `servidor-propio/pi/dominio-corto.sh` y `nomcolab-pi dominio` (ver el paso 5).
- **IndexNow:** la clave `static/3a92e080d98e04396ecf06f47cdca7a4.txt` y
  `herramientas/buscadores/avisar.py`.
- `pruebas/humo.py` revisa robots, sitemap, la portada y el `noindex` de `/unirse`.

## 3. Paso A: quién responde por nom.qd.je (DNS)

DigitalPlat solo **delega**: no guarda registros, solo dice qué servidores de nombres responden por
`nom.qd.je`. Hace falta un DNS gratis que acepte un subdominio como zona. Se usa **deSEC**
(<https://desec.io>), que acepta cualquier correo. Hurricane Electric (dns.he.net) no aceptó ni Gmail
ni Yahoo al registrarse (16-nov).

1. <https://desec.io> → «Sign up»: correo y captcha. Confirmar con el enlace que llega al correo.
2. «+» (nuevo dominio) → `nom.qd.je`. deSEC muestra sus servidores de nombres: `ns1.desec.io` y
   `ns2.desec.org`. Los registros DS (DNSSEC) se ignoran: DigitalPlat no los pide.
3. En el panel de DigitalPlat (<https://dash.domain.digitalplat.org>), en `nom.qd.je`, poner como
   servidores de nombres `ns1.desec.io` y `ns2.desec.org`.
4. En deSEC, dentro de `nom.qd.je`, agregar con el «Subname» **vacío** (es el propio `nom.qd.je`):
   - tipo **A**: `104.21.87.200` y `172.67.145.239`
   - tipo **AAAA**: `2606:4700:3037::ac43:91ef` y `2606:4700:3034::6815:57c8`

   Son las direcciones de Cloudflare de la zona (`dig origen.nomcolab.dpdns.org`; las mismas que las de
   `nomcolab.dpdns.org`). deSEC no tiene ALIAS, así que se ponen a mano. Cloudflare casi nunca las cambia,
   pero si lo hiciera el nombre corto dejaría de abrir. Con un token de la API de deSEC se puede programar
   una revisión que las actualice sola.
5. Los registros TXT del paso B van en el mismo dominio, con «Subname» `_cf-custom-hostname` y
   `_acme-challenge`. Si deSEC rechaza el valor, se escribe entre comillas dobles: `"valor"`.

## 4. Paso B: Cloudflare for SaaS (panel de Cloudflare)

En la zona `nomcolab.dpdns.org` → **SSL/TLS → Custom Hostnames**:

1. «Enable Cloudflare for SaaS». Es gratis hasta 100 nombres; puede pedir una forma de pago
   para el perfil de cobro, pero no cobra mientras no se pasen de 100.
2. **Fallback Origin:** `origen.nomcolab.dpdns.org` → «Add Fallback Origin». Esperar a que diga
   *Active*.
3. **Add Custom Hostname:** `nom.qd.je`; *Certificate validation method*: **TXT**. Lo demás, como
   viene.
4. Cloudflare muestra dos registros TXT: uno de propiedad (`_cf-custom-hostname.nom.qd.je`) y uno del
   certificado (`_acme-challenge.nom.qd.je`). Agregar los dos, tal cual, en deSEC (paso A.5).
5. Esperar a que *Hostname status* y *Certificate status* digan **Active**. Tarda de minutos a un par
   de horas.

(El token del túnel, `~/.cloudflared/cert.pem`, solo puede manejar túneles y DNS. Por eso esta parte
se hace en el panel y no con un programa.)

## 5. Paso C: la Raspberry

La carpeta `servidor-propio/pi` ya está copiada en la Pi como `~/dominio-corto/`. Una sola vez, con la
contraseña de la Pi:

```sh
ssh pi5
doas sh ~/dominio-corto/dominio-corto.sh
```

Esto hace que el túnel también atienda `nom.qd.je` y fija `NOM_NUBE=https://nomcolab.dpdns.org`,
para que los pases los sigan aceptando los APK ya instalados aunque la cuenta se abra con el nombre
corto. También instala la orden `nomcolab-pi` nueva. Si `nom.qd.je` ya contesta, la vuelve dirección
principal (`NOM_CANONICO=https://nom.qd.je`). Si todavía no contesta (el DNS tarda), después basta
con correr esto, que no pide contraseña:

```sh
nomcolab-pi dominio
```

Para regresar la dirección principal a la de cada nombre: borrar la línea `export NOM_CANONICO=…` de
`/etc/nomcolab.conf` y `nomcolab-pi reiniciar`.

## 6. Paso D: los buscadores

Ningún buscador incluye un sitio porque sí: hay que avisarles y comprobar que el sitio es tuyo. Todo
es gratis.

| Buscador | Qué hacer | Cubre |
|---|---|---|
| **Google** | [Search Console](https://search.google.com/search-console) → Agregar propiedad → **Dominio** → `nom.qd.je` → copiar el TXT `google-site-verification=…` en deSEC (registro TXT con «Subname» vacío) → Verificar. Después: Sitemaps → `https://nom.qd.je/sitemap.xml`, e Inspección de URL → `https://nom.qd.je/` → «Solicitar indexación». Hacer lo mismo con `nomcolab.dpdns.org`, poniendo su TXT en el DNS de Cloudflare. | Google, Android, Chrome |
| **Bing** | [Bing Webmaster Tools](https://www.bing.com/webmasters) → «Importar desde Google Search Console» (toma los sitios y los sitemaps ya comprobados). | Bing, DuckDuckGo, Yahoo, Ecosia, Edge, Copilot |
| **Yandex** | [Yandex Webmaster](https://webmaster.yandex.com) → agregar `https://nom.qd.je` → comprobar con DNS (TXT `yandex-verification: …`) → Indexing → Sitemap files. | Yandex |
| **IndexNow** | Después de cada publicación: `python3 4-CODIGO/herramientas/buscadores/avisar.py` (lee el sitemap del servidor y avisa de todas sus páginas). | Bing, Yandex, Seznam, Naver, Yep |

Otras cosas que ayudan y no cuestan:
- En GitHub, el «Website» del repositorio (`TruTruMaster/NOM-COLAB` → About → ⚙) con `https://nom.qd.je`.
- Enlazar `https://nom.qd.je` desde donde se hable del proyecto (perfil, presentación, el README).
- Al compartir el enlace en WhatsApp, Telegram, Discord o Facebook ya sale la tarjeta con la imagen y
  la descripción. Para refrescar la de Facebook:
  <https://developers.facebook.com/tools/debug/>.

**Qué esperar:** buscando `nomcolab`, `nom colab` o `NOM-COLAB`, el sitio debería salir arriba unos
días o semanas después de que Google lo indexe: es un nombre propio y casi no hay competencia. Con
solo `nom` es muy difícil, porque es una palabra muy común (en México la ocupan las Normas Oficiales
Mexicanas). Eso no se arregla con ajustes, solo con que mucha gente enlace y busque el sitio. Los datos
de schema.org dicen que «NOM» es un nombre del sitio, para que Google pueda mostrar «NOM-COLAB» como
nombre del sitio en los resultados.

## 7. Comprobar

```sh
dig +short nom.qd.je                         # direcciones de Cloudflare
curl -sI https://nom.qd.je | head -3         # HTTP/2 200
curl -s https://nom.qd.je/robots.txt         # Sitemap: https://nom.qd.je/sitemap.xml (con NOM_CANONICO)
NOM_BASE=https://nom.qd.je python3 4-CODIGO/pruebas/humo.py
```

Los datos estructurados se revisan en <https://search.google.com/test/rich-results> y
<https://validator.schema.org>. La tarjeta para compartir, en <https://www.opengraph.xyz>.
