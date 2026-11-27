package mx.wack.nom.core.servicio;

import java.util.ArrayDeque;
import java.util.Deque;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Direcciones limpias del sitio de ayuda y legal: /legal/terminos en vez de
 * /legal/terminos/index.html.
 *
 * Los archivos del sitio usan enlaces relativos que terminan en index.html porque la app Android
 * los abre desde sus archivos (file://), donde una carpeta no se puede abrir. Al servirlos por la
 * red (la computadora, la Raspberry o el telefono que comparte), cada enlace relativo se
 * cambia por su direccion completa y sin index.html. Asi la pagina se puede mostrar en una
 * direccion sin diagonal final y sin nombre de archivo, y sus hojas de estilo, scripts e imagenes
 * siguen cargando.
 */
public final class PaginasLimpias {

    /** ayuda, legal, ayuda/seccion, legal/pagina o ayuda/seccion/articulo (sin diagonal al inicio ni al final). */
    private static final Pattern LIMPIA = Pattern.compile("(ayuda|legal)(/[a-z0-9-]+){0,2}");
    private static final Pattern VIEJA = Pattern.compile("((ayuda|legal)(/[a-z0-9-]+){0,2})/(index\\.html)?");
    private static final Pattern ATRIBUTO = Pattern.compile("\\b(href|src|action)=\"([^\"]*)\"");

    private PaginasLimpias() { }

    /** true si la ruta (sin la diagonal inicial) es la direccion limpia de una pagina de ayuda o legal. */
    public static boolean esLimpia(String ruta) {
        return ruta != null && LIMPIA.matcher(ruta).matches();
    }

    /**
     * Si la ruta (sin la diagonal inicial) es la forma vieja de una pagina (con diagonal final o con
     * index.html), regresa su direccion limpia; si no, null.
     */
    public static String limpiaDe(String ruta) {
        if (ruta == null) return null;
        Matcher m = VIEJA.matcher(ruta);
        return m.matches() ? m.group(1) : null;
    }

    /**
     * Cambia los enlaces relativos de una pagina por direcciones completas y limpias.
     *
     * @param html   el index.html tal como esta en los archivos
     * @param pagina la direccion limpia de esa pagina, sin diagonales en los extremos (legal/terminos)
     */
    public static String limpiar(String html, String pagina) {
        Matcher m = ATRIBUTO.matcher(html);
        StringBuilder sb = new StringBuilder(html.length() + 512);
        while (m.find()) {
            String nueva = absoluta(m.group(2), pagina);
            m.appendReplacement(sb, Matcher.quoteReplacement(m.group(1) + "=\"" + nueva + "\""));
        }
        m.appendTail(sb);
        return sb.toString();
    }

    /** La direccion completa de un enlace relativo escrito en la pagina; lo demas se deja igual. */
    static String absoluta(String enlace, String pagina) {
        if (enlace.isEmpty() || enlace.startsWith("#") || enlace.startsWith("/") || enlace.startsWith("?")
                || enlace.matches("^[a-zA-Z][a-zA-Z0-9+.-]*:.*")) return enlace;
        String resto = "";
        int corte = indiceDe(enlace, '#', '?');
        String camino = enlace;
        if (corte >= 0) { camino = enlace.substring(0, corte); resto = enlace.substring(corte); }
        Deque<String> partes = new ArrayDeque<>();
        for (String p : pagina.split("/")) if (!p.isEmpty()) partes.addLast(p);
        for (String p : camino.split("/")) {
            if (p.isEmpty() || p.equals(".")) continue;
            if (p.equals("..")) { if (!partes.isEmpty()) partes.removeLast(); }
            else partes.addLast(p);
        }
        if (!partes.isEmpty() && partes.peekLast().equals("index.html")) partes.removeLast();
        return "/" + String.join("/", partes) + resto;
    }

    private static int indiceDe(String s, char a, char b) {
        int i = s.indexOf(a), j = s.indexOf(b);
        if (i < 0) return j;
        if (j < 0) return i;
        return Math.min(i, j);
    }

    /**
     * La portada servida por HTTP, sin la política de contenido en un {@code <meta>}: el servidor ya la manda en la
     * cabecera (FiltroCabeceras y el servidor del teléfono), y con el {@code <meta>} Chrome no lee por adelantado los
     * programas de la página y los baja de uno en uno (medido el 20-nov a través de Cloudflare: 8.5 s la primera vez,
     * 1.1 s sin él). La app (que abre sus propios archivos, sin cabeceras) la conserva.
     */
    public static String sinPoliticaEnMeta(String html) {
        if (html == null) return null;
        int i = html.indexOf("<meta http-equiv=\"Content-Security-Policy\"");
        if (i < 0) return html;
        int f = html.indexOf('>', i);
        if (f < 0) return html;
        int fin = f + 1;
        while (fin < html.length() && (html.charAt(fin) == '\n' || html.charAt(fin) == '\r' || html.charAt(fin) == ' ')) fin++;
        return html.substring(0, i) + html.substring(fin);
    }

    /**
     * La portada pide sus programas y estilos con {@code ?v=VERSION}: Cloudflare (y el navegador) guardan los .js y .css
     * varias horas, y con el candado de versión un navegador que trajera un .js de antes quedaría con «Actualiza NOM»
     * aunque recargue. Con la versión en la dirección, cada versión nueva es otra dirección. Solo los archivos propios
     * ({@code js/} y {@code css/}); los iconos y lo de fuera, tal cual. La app (que abre sus archivos sin servidor) no
     * pasa por aquí.
     */
    public static String conVersion(String html, String version) {
        if (html == null || version == null || !version.matches("[0-9A-Za-z.-]{1,20}")) return html;
        return html.replaceAll("((?:src|href)=\"(?:js|css)/[^\"?#]+\\.(?:js|css))\"", "$1?v=" + version + "\"");
    }
}
