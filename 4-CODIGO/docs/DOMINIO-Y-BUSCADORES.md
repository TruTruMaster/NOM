# El nombre corto nom.qd.je y los buscadores

Escrito el 16 de noviembre de 2026. El sitio tiene dos nombres que llevan al mismo servidor (la
Raspberry, con sus cuentas y su base de datos):

| Nombre | Para qué |
|---|---|
| `https://www.nom.qd.je` | La dirección principal: la que se ve en Google, Bing y los demás. |
| `nom.qd.je` | El nombre para compartir: redirige a `www.nom.qd.je`. **Al 17-nov solo por `http://`**: GitHub Pages todavía no emite su certificado (por `https://` el navegador ve el de `*.github.io` y avisa). |
| `https://nomcolab.dpdns.org` | El de siempre. **No se apaga:** las apps 0.15.x le hablan a este nombre (desde la 0.16.0 la app usa `www.nom.qd.je` y trata los dos como la misma nube), los pases van firmados con él y quien ya guardó cosas en el navegador con este nombre las sigue teniendo ahí. |

No se redirige un nombre al otro, porque el navegador guarda la sesión y los documentos de cada
nombre por separado. Una redirección dejaría a la gente sin sesión y sin lo que tenía guardado sin
Internet. Para que los buscadores no los cuenten como dos sitios, cada página dice cuál es su dirección
principal (`<link rel="canonical">`). Esa dirección sale de `NOM_CANONICO` en `/etc/nomcolab.conf`.

## 1. Cómo quedó (16 de noviembre de 2026)

```
https://www.nom.qd.je ──CNAME──▶ origen.nomcolab.dpdns.org ──▶ Cloudflare for SaaS ──▶ túnel ──▶ Raspberry
https://nom.qd.je ──A/AAAA──▶ GitHub Pages (repo TruTruMaster/nom-qd-je) ──redirige──▶ https://www.nom.qd.je
```

- La dirección principal es **`https://www.nom.qd.je`** (`NOM_CANONICO`), por el mismo túnel de
  siempre: la casa no se expone y no hay puertos abiertos.
- `nom.qd.je` solo redirige, conservando la ruta, la consulta y el fragmento (los enlaces para unirse
  siguen sirviendo).

**Por qué `nom.qd.je` no puede ser la principal:** `qd.je` no está en la
[lista pública de sufijos](https://publicsuffix.org/list/public_suffix_list.dat) (`dpdns.org` sí). Por
eso los servicios tratan a DigitalPlat como dueño de todo `qd.je`:
- Cloudflare no la acepta como zona. Cloudflare for SaaS solo activa un nombre que sea CNAME hacia la
  zona, y en la raíz de una zona no se puede poner CNAME (DigitalPlat: «CNAME records cannot be
  created at the zone apex»). La raíz solo se puede con el plan Enterprise.
- Vercel funcionó como reenvío (probado de punta a punta; ver `servidor-propio/vercel/`), pero pide
  comprobar el dominio con un TXT en `_vercel.qd.je`, que es de DigitalPlat, porque otra cuenta de
  Vercel ya usa un `*.qd.je`.
- Netlify corta las peticiones reenviadas a los 26 s (el canal en vivo) y tiene el mismo bloqueo de
  dominio. Deno Deploy se cierra en abril de 2027.
- La salida limpia es que DigitalPlat agregue `qd.je` a la lista pública de sufijos, como hizo con
  `dpdns.org`. Entonces `nom.qd.je` entra a Cloudflare gratis como zona propia, igual que
  `nomcolab.dpdns.org`, y pasa a ser la principal (`NOM_CANONICO=https://nom.qd.je`).

## 2. Lo que hay en el código

- **DNS en Cloudflare:** `origen.nomcolab.dpdns.org` → el túnel `nomcolab` (CNAME con proxy). Será el
  «origen de respaldo».
- **Servidor** (`Buscadores.java`, `PaginasApi.java`): la portada y cada página de ayuda y legal
  llevan su dirección principal, lo que muestran las redes al compartir el enlace (Open Graph y
  tarjeta de X, con `img/compartir.png` de 1200×630) y datos de schema.org. En la portada son WebSite
  con los nombres con que la buscan (NOM y los de antes: NOM-COLAB, NOMCOLAB, nomcolab), Organization con el logo y
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

## 3. Lo que se configuró afuera

**DigitalPlat** (DNS de `nom.qd.je`, «Use DigitalPlat DNS»):

| Nombre | Tipo | Valor |
|---|---|---|
| `www` | CNAME | `origen.nomcolab.dpdns.org.` (con el punto final; sin él, el panel le agrega `.nom.qd.je`) |
| (raíz) | A | `185.199.108.153`, `185.199.109.153`, `185.199.110.153`, `185.199.111.153` (GitHub Pages) |
| (raíz) | AAAA | `2606:50c0:8000::153`, `2606:50c0:8001::153`, `2606:50c0:8002::153`, `2606:50c0:8003::153` |

**Cloudflare**, zona `nomcolab.dpdns.org`: `origen.nomcolab.dpdns.org` (CNAME con proxy al túnel) es el
origen de respaldo de Cloudflare for SaaS; el nombre personalizado `www.nom.qd.je` está activo.

**GitHub:** el repositorio público `TruTruMaster/nom-qd-je` (Pages, dominio `nom.qd.je`) trae `index.html` y
`404.html` que mandan a `https://www.nom.qd.je` con la misma ruta. **Pendiente:** el certificado de `nom.qd.je`
(la API de Pages dice `https_enforced: false` y sin certificado); cuando GitHub lo emita, activar «Enforce HTTPS».
Las dos páginas todavía dicen «NOM-COLAB»: cambiarlas a «NOM» al publicar la 0.16.0.

**Raspberry:** `doas sh ~/dominio-corto/dominio-corto.sh` (una vez) puso el túnel con
`www.nom.qd.je`, `NOM_NUBE=https://nomcolab.dpdns.org` (los pases valen en los APK ya instalados) y la
orden `nomcolab-pi dominio`, que dejó `NOM_CANONICO=https://www.nom.qd.je`. Para cambiar la principal:
editar esa línea en `/etc/nomcolab.conf` y `nomcolab-pi reiniciar`.

## 4. Los buscadores

Ningún buscador incluye un sitio porque sí: hay que avisarles y comprobar que el sitio es tuyo. Todo
es gratis.

| Buscador | Qué hacer | Cubre |
|---|---|---|
| **Google** | [Search Console](https://search.google.com/search-console) → Agregar propiedad → **Dominio** → `nom.qd.je` (cubre también `www`) → copiar el TXT `google-site-verification=…` en DigitalPlat (registro TXT en la raíz) → Verificar. Después: Sitemaps → `https://www.nom.qd.je/sitemap.xml`, e Inspección de URL → `https://www.nom.qd.je/` → «Solicitar indexación». Hacer lo mismo con `nomcolab.dpdns.org`, poniendo su TXT en el DNS de Cloudflare. | Google, Android, Chrome |
| **Bing** | [Bing Webmaster Tools](https://www.bing.com/webmasters) → «Importar desde Google Search Console» (toma los sitios y los sitemaps ya comprobados). | Bing, DuckDuckGo, Yahoo, Ecosia, Edge, Copilot |
| **Yandex** | [Yandex Webmaster](https://webmaster.yandex.com) → agregar `https://www.nom.qd.je` → comprobar con DNS (TXT `yandex-verification: …`) → Indexing → Sitemap files. | Yandex |
| **IndexNow** | Después de cada publicación: `python3 4-CODIGO/herramientas/buscadores/avisar.py` (lee el sitemap del servidor y avisa de todas sus páginas). | Bing, Yandex, Seznam, Naver, Yep |

Otras cosas que ayudan y no cuestan:
- En GitHub, el «Website» del repositorio (`TruTruMaster/NOM-COLAB` → About → ⚙) con `https://www.nom.qd.je`.
- Enlazar `https://nom.qd.je` desde donde se hable del proyecto (perfil, presentación, el README).
- Al compartir el enlace en WhatsApp, Telegram, Discord o Facebook ya sale la tarjeta con la imagen y
  la descripción. Para refrescar la de Facebook:
  <https://developers.facebook.com/tools/debug/>.

**Qué esperar:** buscando `nomcolab`, `nom colab` o `NOM`, el sitio debería salir arriba unos
días o semanas después de que Google lo indexe: es un nombre propio y casi no hay competencia. Con
solo `nom` es muy difícil, porque es una palabra muy común (en México la ocupan las Normas Oficiales
Mexicanas). Eso no se arregla con ajustes, solo con que mucha gente enlace y busque el sitio. Los datos
de schema.org dicen que «NOM» es un nombre del sitio, para que Google pueda mostrar «NOM» como
nombre del sitio en los resultados.

## 5. Comprobar

```sh
dig +short www.nom.qd.je                     # origen.nomcolab.dpdns.org y direcciones de Cloudflare
curl -s https://www.nom.qd.je/robots.txt     # Sitemap: https://www.nom.qd.je/sitemap.xml
curl -s http://nom.qd.je/ayuda | grep -o 'url=[^"]*'   # url=https://www.nom.qd.je/ (redirige con meta y JavaScript, no con 301: curl -L no la sigue)
NOM_BASE=https://www.nom.qd.je python3 4-CODIGO/pruebas/humo.py
```

Los datos estructurados se revisan en <https://search.google.com/test/rich-results> y
<https://validator.schema.org>. La tarjeta para compartir, en <https://www.opengraph.xyz>.
