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

    private static final java.util.regex.Pattern IPV4 = java.util.regex.Pattern.compile("\\d{1,3}(\\.\\d{1,3}){3}");
    private static final String[] LOCALES = { ".local", ".lan", ".home", ".home.arpa", ".internal", ".localdomain", ".localhost", ".intranet", ".fritz.box" };

    public static boolean permitido(String host, java.util.Set<String> conocidos) {
        if (host == null || host.isBlank()) return true;   // (los navegadores siempre lo mandan)
        String h = host.trim().toLowerCase(java.util.Locale.ROOT);
        if (h.startsWith("[")) return h.matches("\\[[0-9a-f:.%a-z]+\\](:\\d{1,5})?");   // IPv6 escrita
        h = h.replaceFirst(":\\d{1,5}$", "");
        if (h.endsWith(".")) h = h.substring(0, h.length() - 1);
        if (h.isEmpty() || !h.matches("[a-z0-9.-]+")) return false;
        if (IPV4.matcher(h).matches() || h.equals("localhost") || !h.contains(".")) return true;
        for (String s : LOCALES) if (h.endsWith(s)) return true;
        return conocidos.contains(h);
    }

}
