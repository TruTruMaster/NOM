package mx.wack.nom.servidor.api;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jakarta.servlet.http.HttpServletRequest;
import mx.wack.nom.servidor.seguridad.Origen;

/**
 * Lo que leen los buscadores (Google, Bing, DuckDuckGo, Yandex, Ecosia...) y las redes al compartir un
 * enlace (WhatsApp, Facebook, X, Telegram, Discord): la direccion principal de cada pagina, su
 * titulo y resumen, la imagen para compartir y los datos estructurados (schema.org).
 *
 * La direccion principal ("canonica") sale de NOM_CANONICO (p. ej. https://nom.qd.je). Asi el sitio
 * se puede abrir por varios nombres (nom.qd.je y nomcolab.dpdns.org) y los buscadores juntan todo en
 * uno solo, sin que nadie pierda lo que tiene guardado en el navegador con el otro nombre (cada
 * nombre tiene su propio almacenamiento: por eso no se redirige). Sin NOM_CANONICO, cada nombre es
 * su propia direccion principal.
 *
 * Los archivos (static/) no llevan nada de esto: la app de Android los abre desde sus archivos y ahi
 * no sirve. Se agrega al servirlos por la red.
 */
public final class Buscadores {

    /** La direccion con la que se escribieron sitemap.xml, robots.txt y los enlaces de la ayuda. */
    static final String ESCRITA = "https://www.nom.qd.je";
    private static final Pattern BASE_VALIDA = Pattern.compile("https?://[a-z0-9.-]+(:\\d{1,5})?");
    private static final Pattern TITULO = Pattern.compile("<title>([^<]*)</title>");
    private static final Pattern RESUMEN = Pattern.compile("<meta name=\"description\" content=\"([^\"]*)\"");
    private static final Pattern MIGAS = Pattern.compile("<nav class=\"migas\"[^>]*>(.*?)</nav>", Pattern.DOTALL);
    private static final Pattern MIGA = Pattern.compile("<a href=\"([^\"]*)\">([^<]*)</a>");

    static final String NOMBRE = "NOM";
    static final String RESUMEN_INICIO = "Escribe, calcula, presenta y dibuja en equipo, al mismo tiempo, con o sin Internet. "
            + "Abre y guarda Word, Excel y PowerPoint. Gratis y de código abierto, en el navegador y en Android.";
    static final String APK = "https://github.com/TruTruMaster/NOM-COLAB/releases/latest/download/NOM-COLAB.apk";
    static final String CODIGO = "https://github.com/TruTruMaster/NOM-COLAB";

    private Buscadores() { }

    /** La direccion principal del sitio, sin diagonal final (siempre un esquema y un nombre validos). */
    static String base(HttpServletRequest req) {
        String c = System.getenv("NOM_CANONICO");
        String b = (c != null && !c.isBlank() ? c : Origen.publico(req)).trim().toLowerCase(Locale.ROOT).replaceAll("/+$", "");
        // El nombre viene de la cabecera Host: si no es un nombre limpio, no se escribe en la pagina.
        return BASE_VALIDA.matcher(b).matches() ? b : ESCRITA;
    }

    /** robots.txt: todo se puede leer menos la API y las direcciones de un solo uso (unirse, proyectar). */
    static String robots(String base) {
        return "# NOM\n"
                + "User-agent: *\n"
                + "Allow: /\n"
                + "Disallow: /api/\n"
                + "Disallow: /unirse\n"
                + "Disallow: /proyectar\n"
                + "Disallow: /ayuda/buscar\n"
                + "\n"
                + "Sitemap: " + base + "/sitemap.xml\n";
    }

    /** sitemap.xml tal como lo escribio ayuda/generar.py, con la direccion principal. */
    static String sitemap(String escrito, String base) {
        return escrito.replace(ESCRITA + "/", base + "/").replace(ESCRITA + "<", base + "/<");
    }

    /**
     * Agrega a la pagina (antes de </head>) la direccion principal, lo que muestran las redes al
     * compartirla y sus datos estructurados.
     *
     * @param ruta la direccion de la pagina sin la diagonal inicial ("" para la portada)
     */
    static String cabecera(String html, String base, String ruta) {
        int fin = html.indexOf("</head>");
        if (fin < 0) return html;
        boolean portada = ruta.isEmpty();
        String url = base + "/" + ruta;
        String titulo = grupo(TITULO, html, NOMBRE);
        String resumen = portada ? RESUMEN_INICIO : grupo(RESUMEN, html, RESUMEN_INICIO);
        String imagen = base + "/img/compartir.png";
        StringBuilder s = new StringBuilder(2048);
        s.append("<!-- Para buscadores y redes (lo agrega el servidor) -->\n");
        s.append("<link rel=\"canonical\" href=\"").append(url).append("\">\n");
        s.append("<link rel=\"alternate\" hreflang=\"es\" href=\"").append(url).append("\">\n");
        s.append("<link rel=\"alternate\" hreflang=\"x-default\" href=\"").append(url).append("\">\n");
        s.append("<meta name=\"robots\" content=\"index, follow, max-image-preview:large, max-snippet:-1\">\n");
        s.append("<meta property=\"og:type\" content=\"").append(portada ? "website" : "article").append("\">\n");
        s.append("<meta property=\"og:site_name\" content=\"").append(NOMBRE).append("\">\n");
        s.append("<meta property=\"og:locale\" content=\"es_MX\">\n");
        s.append("<meta property=\"og:url\" content=\"").append(url).append("\">\n");
        s.append("<meta property=\"og:title\" content=\"").append(titulo).append("\">\n");
        s.append("<meta property=\"og:description\" content=\"").append(resumen).append("\">\n");
        s.append("<meta property=\"og:image\" content=\"").append(imagen).append("\">\n");
        s.append("<meta property=\"og:image:type\" content=\"image/png\">\n");
        s.append("<meta property=\"og:image:width\" content=\"1200\">\n");
        s.append("<meta property=\"og:image:height\" content=\"630\">\n");
        s.append("<meta property=\"og:image:alt\" content=\"NOM: documentos, hojas de cálculo y presentaciones en equipo, con o sin Internet\">\n");
        s.append("<meta name=\"twitter:card\" content=\"summary_large_image\">\n");
        s.append("<meta name=\"twitter:title\" content=\"").append(titulo).append("\">\n");
        s.append("<meta name=\"twitter:description\" content=\"").append(resumen).append("\">\n");
        s.append("<meta name=\"twitter:image\" content=\"").append(imagen).append("\">\n");
        s.append("<script type=\"application/ld+json\">")
                .append(portada ? datosPortada(base) : datosPagina(html, base, url, titulo, resumen))
                .append("</script>\n");
        return html.substring(0, fin) + s + html.substring(fin);
    }

    /** schema.org de la portada: el sitio (con los nombres con que lo buscan), quien lo hace y la app. */
    static String datosPortada(String base) {
        String org = base + "/#organizacion";
        return "{\"@context\":\"https://schema.org\",\"@graph\":["
                + "{\"@type\":\"WebSite\",\"@id\":\"" + base + "/#sitio\",\"url\":\"" + base + "/\",\"name\":\"" + NOMBRE + "\","
                // (el nombre de antes, para que también se encuentre así)
                + "\"alternateName\":[\"NOM-COLAB\",\"NOMCOLAB\",\"NOM COLAB\",\"nomcolab\"],"
                + "\"description\":" + json(RESUMEN_INICIO) + ",\"inLanguage\":\"es\",\"publisher\":{\"@id\":\"" + org + "\"}},"
                + "{\"@type\":\"Organization\",\"@id\":\"" + org + "\",\"name\":\"" + NOMBRE + "\",\"url\":\"" + base + "/\","
                + "\"logo\":{\"@type\":\"ImageObject\",\"url\":\"" + base + "/img/icono-512.png\",\"width\":512,\"height\":512},"
                + "\"sameAs\":[\"" + CODIGO + "\"]},"
                + "{\"@type\":\"WebApplication\",\"@id\":\"" + base + "/#app\",\"name\":\"" + NOMBRE + "\",\"url\":\"" + base + "/\","
                + "\"description\":" + json(RESUMEN_INICIO) + ","
                + "\"applicationCategory\":\"BusinessApplication\",\"applicationSubCategory\":\"Suite de oficina colaborativa\","
                + "\"operatingSystem\":\"Navegador web, Android, Windows, macOS, Linux\",\"browserRequirements\":\"Requiere JavaScript\","
                + "\"inLanguage\":\"es\",\"isAccessibleForFree\":true,\"license\":\"https://www.gnu.org/licenses/gpl-3.0.html\","
                + "\"downloadUrl\":\"" + APK + "\",\"installUrl\":\"" + APK + "\",\"image\":\"" + base + "/img/compartir.png\","
                + "\"screenshot\":\"" + base + "/img/compartir.png\","
                + "\"offers\":{\"@type\":\"Offer\",\"price\":\"0\",\"priceCurrency\":\"MXN\"},"
                + "\"featureList\":[\"Documentos de texto como Word\",\"Hojas de cálculo como Excel\",\"Presentaciones como PowerPoint\","
                + "\"Diagramas y dibujos\",\"Edición en equipo en tiempo real\",\"Funciona sin Internet en la red local\","
                + "\"Abre y guarda .docx, .xlsx y .pptx\",\"Cifrado de extremo a extremo\"],"
                + "\"publisher\":{\"@id\":\"" + org + "\"}}"
                + "]}";
    }

    /** schema.org de una pagina de ayuda o legal: el articulo y el camino de migas. */
    static String datosPagina(String html, String base, String url, String titulo, String resumen) {
        List<String[]> migas = new ArrayList<>();
        migas.add(new String[] { base + "/", NOMBRE });
        Matcher n = MIGAS.matcher(html);
        if (n.find()) {
            Matcher a = MIGA.matcher(n.group(1));
            while (a.find()) migas.add(new String[] { a.group(1).startsWith("/") ? base + a.group(1) : a.group(1), a.group(2) });
        }
        String nombre = titulo.replaceAll("\\s*·\\s*(Ayuda de )?NOM\\s*$", "");
        migas.add(new String[] { url, nombre });
        StringBuilder l = new StringBuilder();
        for (int i = 0; i < migas.size(); i++) {
            if (i > 0) l.append(',');
            l.append("{\"@type\":\"ListItem\",\"position\":").append(i + 1)
                    .append(",\"name\":").append(json(desescapar(migas.get(i)[1])))
                    .append(",\"item\":").append(json(migas.get(i)[0])).append('}');
        }
        return "{\"@context\":\"https://schema.org\",\"@graph\":["
                + "{\"@type\":\"WebPage\",\"@id\":\"" + url + "\",\"url\":\"" + url + "\",\"name\":" + json(desescapar(nombre)) + ","
                + "\"description\":" + json(desescapar(resumen)) + ",\"inLanguage\":\"es\","
                + "\"isPartOf\":{\"@type\":\"WebSite\",\"@id\":\"" + base + "/#sitio\",\"name\":\"" + NOMBRE + "\",\"url\":\"" + base + "/\"}},"
                + "{\"@type\":\"BreadcrumbList\",\"itemListElement\":[" + l + "]}"
                + "]}";
    }

    private static String grupo(Pattern p, String html, String si) {
        Matcher m = p.matcher(html);
        return m.find() ? m.group(1) : si;
    }

    /** El texto de la pagina viene escapado para HTML; dentro del JSON va como texto. */
    private static String desescapar(String s) {
        return s.replace("&quot;", "\"").replace("&#39;", "'").replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&");
    }

    /** Texto como cadena JSON, segura dentro de un <script> (sin "</" ni "<!--"). */
    static String json(String s) {
        StringBuilder b = new StringBuilder(s.length() + 2).append('"');
        for (char c : s.toCharArray()) {
            switch (c) {
                case '"' -> b.append("\\\"");
                case '\\' -> b.append("\\\\");
                case '\n' -> b.append("\\n");
                case '\r' -> b.append("\\r");
                case '\t' -> b.append("\\t");
                case '<' -> b.append("\\u003c");
                case '>' -> b.append("\\u003e");
                case '&' -> b.append("\\u0026");
                default -> { if (c < 0x20) b.append(String.format("\\u%04x", (int) c)); else b.append(c); }
            }
        }
        return b.append('"').toString();
    }
}
