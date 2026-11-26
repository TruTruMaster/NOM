package mx.wack.nom.servidor.seguridad;

import java.net.InetAddress;

import jakarta.servlet.http.HttpServletRequest;

/**
 * De donde viene una peticion (su direccion de red), para los limites contra el abuso: intentos de
 * contrasena y cuentas nuevas por hora.
 *
 * Detras de un tunel o de un balanceador todas las peticiones llegan desde el mismo lugar y la
 * direccion real viene en una cabecera. Esa cabecera solo se cree cuando quien la entrega es de
 * confianza (si no, cualquiera podria escribirla): solo este mismo equipo (el tunel de Cloudflare o
 * Funnel corren aqui). Ningun otro equipo de la red, ni por su direccion de enlace (fe80::).
 * Una peticion de este mismo equipo sin cabecera es del propio equipo: null (sin limites).
 */
public final class Origen {
    private Origen() { }

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
            confiable = a == null || a.isLoopbackAddress();
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
        boolean local;
        try {
            InetAddress a = remota == null ? null : InetAddress.getByName(remota);
            local = a == null || a.isLoopbackAddress();
        } catch (Exception e) { local = false; }
        if (!local) return agrupar(remota);
        // El nombre corto (nom.qd.je) llega por Vercel, que reenvia a la nube por Cloudflare: la
        // direccion de la persona viene en X-Real-IP, y solo se cree si trae la clave que Vercel
        // agrega en cada peticion (NOM_PROXY_SECRETO; Vercel reemplaza la que mande cualquiera).
        String ip = deVercel(req);
        if (ip != null) return agrupar(ip);
        // CF-Connecting-IP la pone Cloudflare, que entra por el túnel de este mismo equipo.
        String r = req.getHeader("CF-Connecting-IP");
        if (r == null || r.isBlank()) {
            r = req.getHeader("X-Forwarded-For");
            // Cada intermediario agrega al final la dirección de donde le llegó: se toma la última.
            if (r != null && r.contains(",")) r = r.substring(r.lastIndexOf(',') + 1);
            // Si llega con el puerto ("203.0.113.5:51234"), se queda solo la direccion.
            if (r != null && r.trim().matches("\\d+(\\.\\d+){3}:\\d+")) r = r.trim().substring(0, r.trim().lastIndexOf(':'));
        }
        if (r == null || r.isBlank()) return null;
        return agrupar(r.trim());
    }

    /**
     * Para los topes (contraseñas, registros, códigos), una IPv6 de Internet cuenta por su red /64 (la que tiene cada
     * casa o servidor). Las IPv4, las de la red local (fc00::/7, fe80::/10) y lo que no es una IP se quedan como llegan.
     */
    static String agrupar(String ip) {
        if (ip == null || ip.indexOf(':') < 0 || !ip.matches("[0-9a-fA-F:.]{2,45}")) return ip;
        try {
            InetAddress a = InetAddress.getByName(ip);
            byte[] b = a.getAddress();
            if (!(a instanceof java.net.Inet6Address)) return a.getHostAddress();   // ::ffff:a.b.c.d → a.b.c.d
            if (a.isLoopbackAddress() || a.isLinkLocalAddress() || a.isAnyLocalAddress() || (b[0] & 0xfe) == 0xfc) return ip;
            return String.format(java.util.Locale.ROOT, "%x:%x:%x:%x::/64",
                    ((b[0] & 255) << 8) | (b[1] & 255), ((b[2] & 255) << 8) | (b[3] & 255),
                    ((b[4] & 255) << 8) | (b[5] & 255), ((b[6] & 255) << 8) | (b[7] & 255));
        } catch (Exception e) {
            return ip;
        }
    }
}
