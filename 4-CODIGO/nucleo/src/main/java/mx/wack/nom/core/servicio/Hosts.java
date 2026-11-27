package mx.wack.nom.core.servicio;

/**
 * Qué nombres de servidor (cabecera Host) atiende NOM: una IP, localhost, los nombres de la red local (sin punto, .local,
 * .lan…) y los de la nube que se configuran (con y sin «www.»). Lo usan el servidor (FiltroRutas) y el del teléfono.
 */
public final class Hosts {

    private Hosts() { }

    public static java.util.Set<String> conocidos(String... fuentes) {
        java.util.Set<String> r = new java.util.HashSet<>();
        for (String f : fuentes) {
            if (f == null) continue;
            for (String x : f.split("[,\\s]+")) {
                String h = x.trim().toLowerCase(java.util.Locale.ROOT).replaceFirst("^https?://", "").replaceAll("[/:].*$", "");
                if (h.isEmpty()) continue;
                r.add(h);
                r.add(h.startsWith("www.") ? h.substring(4) : "www." + h);
            }
        }
        return r;
    }

    /** ¿La dirección (tal como llegó, sin la consulta) tiene solo caracteres y partes simples? */
    public static boolean rutaValida(String uri) {
        if (uri == null || uri.isEmpty() || uri.charAt(0) != '/') return false;
        if (uri.indexOf(';') >= 0 || uri.indexOf('\\') >= 0 || uri.contains("//")) return false;
        String bajo = uri.toLowerCase(java.util.Locale.ROOT);
        if (bajo.contains("%2e") || bajo.contains("%2f") || bajo.contains("%5c") || bajo.contains("%25") || bajo.contains("%00") || bajo.contains("%3b")) return false;
        for (String s : uri.split("/", -1)) if (s.equals(".") || s.equals("..")) return false;
        for (int i = 0; i < uri.length(); i++) {
            char c = uri.charAt(i);
            if (c < 0x21 || c > 0x7e) return false;
        }
        return true;
    }

    private static final java.util.regex.Pattern IPV4 = java.util.regex.Pattern.compile("\\d{1,3}(\\.\\d{1,3}){3}");
    private static final String[] LOCALES = { ".local", ".lan", ".home", ".home.arpa", ".internal", ".localdomain", ".localhost", ".intranet", ".fritz.box" };

    public static boolean permitido(String host, java.util.Set<String> conocidos) {
        if (host == null || host.isBlank()) return true;   // (los navegadores siempre lo mandan)
        String h = host.trim().toLowerCase(java.util.Locale.ROOT);
        if (h.startsWith("[")) return h.matches("\\[[0-9a-f:.]*:[0-9a-f:.]*(%[a-z0-9_.-]{1,32})?\\](:\\d{1,5})?");   // IPv6 escrita
        h = h.replaceFirst(":\\d{1,5}$", "");
        if (h.endsWith(".")) h = h.substring(0, h.length() - 1);
        if (h.isEmpty() || !h.matches("[a-z0-9.-]+")) return false;
        if (IPV4.matcher(h).matches() || h.equals("localhost") || !h.contains(".")) return true;
        for (String s : LOCALES) if (h.endsWith(s)) return true;
        return conocidos.contains(h);
    }

    /**
     * Para los topes de cuentas (como Origen.agrupar en la computadora): una IPv6 de Internet cuenta por su /64; las
     * IPv4 (también ::ffff:a.b.c.d), las IPv6 de la red local (fe80::/10, fc00::/7) y lo que no es una IP quedan como
     * llegan, para que los equipos del mismo Wi-Fi no compartan un solo tope.
     */
    public static String agruparRedPublica(String ip) {
        if (ip == null || ip.indexOf(':') < 0 || !ip.matches("[0-9a-fA-F:.%a-zA-Z0-9_]{2,64}")) return ip;
        try {
            java.net.InetAddress a = java.net.InetAddress.getByName(ip.replaceFirst("%.*$", ""));
            byte[] b = a.getAddress();
            if (b.length == 4) return a.getHostAddress();
            if (a.isLoopbackAddress() || a.isLinkLocalAddress() || a.isAnyLocalAddress() || (b[0] & 0xfe) == 0xfc) return ip;
        } catch (Exception e) {
            return ip;
        }
        return agruparRed(ip);
    }

    /** Agrupa una dirección IPv6 por su prefijo /64 (los demás valores salen igual). */
    public static String agruparRed(String ip) {
        if (ip == null || !ip.contains(":")) return ip;
        try {
            java.net.InetAddress a = java.net.InetAddress.getByName(ip.replaceFirst("%.*$", ""));
            byte[] b = a.getAddress();
            if (b.length == 4) return a.getHostAddress();
            if (b.length == 16) {
                return String.format(java.util.Locale.ROOT, "%x:%x:%x:%x::/64",
                        ((b[0] & 255) << 8) | (b[1] & 255), ((b[2] & 255) << 8) | (b[3] & 255),
                        ((b[4] & 255) << 8) | (b[5] & 255), ((b[6] & 255) << 8) | (b[7] & 255));
            }
        } catch (Exception ignorado) { }
        return ip;
    }

}
