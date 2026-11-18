package mx.wack.nom.servidor.seguridad;

import java.net.InetAddress;

import jakarta.servlet.http.HttpServletRequest;

/**
 * De donde viene una peticion (su direccion de red), para los limites contra el abuso: intentos de
 * contrasena y cuentas nuevas por hora.
 *
 * Detras de un tunel o de un balanceador todas las peticiones llegan desde el mismo lugar y la
 * direccion real viene en una cabecera. Esa cabecera solo se cree cuando quien la entrega es de
 * confianza (si no, cualquiera podria escribirla):
 *  - este mismo equipo (el tunel de Cloudflare o Funnel corren aqui), o
 *  - en Azure, una direccion "de enlace" (169.254.x.x, fe80::), que es desde donde entrega las
 *    peticiones esa plataforma. Fuera de Azure no: cualquier equipo de la misma red (la casa de la
 *    Raspberry, el salon) puede llegar por su direccion fe80:: y escribir la cabecera que quiera.
 * Una peticion de este mismo equipo sin cabecera es del propio equipo: null (sin limites).
 */
public final class Origen {
    private Origen() { }

    /** ¿Corre en Azure App Service? (la plataforma pone WEBSITE_SITE_NAME a cada Web App). */
    private static final boolean EN_AZURE = System.getenv("WEBSITE_SITE_NAME") != null;

    /** Direcciones de enlace de las que se creen las cabeceras de reenvio: solo en Azure. */
    static boolean enlaceConfiable(InetAddress a) { return EN_AZURE && a != null && a.isLinkLocalAddress(); }

    /**
     * La direccion publica de este servidor tal como la ve quien llama (p. ej. https://nomcolab.dpdns.org):
     * NOM_NUBE si esta puesta; si no, el esquema (el que dice el tunel si la peticion viene de este
     * equipo) y el nombre del Host. Va dentro de los pases que firma este servidor.
     */
    public static String publico(HttpServletRequest req) {
        String fija = System.getenv("NOM_NUBE");
        if (fija != null && !fija.isBlank()) return fija.trim().replaceAll("/+$", "");
        String remota = req.getRemoteAddr();
        boolean confiable;
        try {
            InetAddress a = remota == null ? null : InetAddress.getByName(remota);
            confiable = a == null || a.isLoopbackAddress() || enlaceConfiable(a);
        } catch (Exception e) { confiable = false; }
        String esquema = req.getScheme();
        String fp = req.getHeader("X-Forwarded-Proto");
        if (confiable && fp != null && fp.matches("(?i)https?")) esquema = fp.toLowerCase(java.util.Locale.ROOT);
        String host = req.getHeader("Host");
        if (host == null || host.isBlank()) host = req.getServerName() + ":" + req.getServerPort();
        return esquema + "://" + host.trim().toLowerCase(java.util.Locale.ROOT);
    }

    /** La IP de la persona si la peticion viene del reenvio de Vercel con la clave correcta; si no, null. */
    static String deVercel(HttpServletRequest req) {
        String secreto = System.getenv("NOM_PROXY_SECRETO");
        String clave = req.getHeader("X-Nom-Proxy");
        if (secreto == null || secreto.length() < 32 || clave == null) return null;
        if (!java.security.MessageDigest.isEqual(secreto.getBytes(java.nio.charset.StandardCharsets.UTF_8),
                clave.getBytes(java.nio.charset.StandardCharsets.UTF_8))) return null;
        // Vercel pone la IP de la persona en X-Vercel-Forwarded-For (reemplaza la que mande cualquiera);
        // X-Real-IP tambien la trae, pero Cloudflare la quita por el camino.
        String ip = req.getHeader("X-Vercel-Forwarded-For");
        if (ip == null || ip.isBlank()) ip = req.getHeader("X-Real-IP");
        if (ip == null || ip.isBlank()) return null;
        ip = ip.contains(",") ? ip.substring(0, ip.indexOf(',')).trim() : ip.trim();
        return ip.matches("[0-9a-fA-F:.]{2,45}") ? ip : null;
    }

    /** ¿Es un origen de la red local (10.x, 192.168.x, 172.16-31.x), con cualquier puerto? */
    public static boolean deRedPrivada(String origen) {
        if (origen == null) return false;
        java.util.regex.Matcher m = java.util.regex.Pattern.compile("^https?://(\\d{1,3})\\.(\\d{1,3})\\.(\\d{1,3})\\.(\\d{1,3})(:\\d{1,5})?$").matcher(origen.trim());
        if (!m.matches()) return false;
        int a = Integer.parseInt(m.group(1)), b = Integer.parseInt(m.group(2));
        if (a > 255 || b > 255 || Integer.parseInt(m.group(3)) > 255 || Integer.parseInt(m.group(4)) > 255) return false;
        return a == 10 || (a == 192 && b == 168) || (a == 172 && b >= 16 && b <= 31);
    }

    public static String de(HttpServletRequest req) {
        String remota = req.getRemoteAddr();
        boolean local, enlace = false;
        try {
            InetAddress a = remota == null ? null : InetAddress.getByName(remota);
            local = a == null || a.isLoopbackAddress();
            enlace = enlaceConfiable(a);
        } catch (Exception e) { local = false; }
        if (!local && !enlace) return remota;
        // El nombre corto (nom.qd.je) llega por Vercel, que reenvia a la nube por Cloudflare: la
        // direccion de la persona viene en X-Real-IP, y solo se cree si trae la clave que Vercel
        // agrega en cada peticion (NOM_PROXY_SECRETO; Vercel reemplaza la que mande cualquiera).
        String ip = deVercel(req);
        if (ip != null) return ip;
        String r = req.getHeader("CF-Connecting-IP");
        if (r == null || r.isBlank()) {
            r = req.getHeader("X-Forwarded-For");
            if (r != null && r.contains(",")) r = r.substring(0, r.indexOf(','));
            // Azure la manda con el puerto ("203.0.113.5:51234"): se queda solo la direccion.
            if (r != null && r.trim().matches("\\d+(\\.\\d+){3}:\\d+")) r = r.substring(0, r.lastIndexOf(':'));
        }
        if (r == null || r.isBlank()) return local ? null : remota;
        return r.trim();
    }
}
