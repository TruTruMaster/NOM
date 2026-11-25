package mx.wack.nom.servidor.seguridad;

import java.io.IOException;

import org.springframework.core.Ordered;
import org.springframework.core.annotation.Order;
import org.springframework.stereotype.Component;
import org.springframework.web.filter.OncePerRequestFilter;

import jakarta.servlet.FilterChain;
import jakarta.servlet.ServletException;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;

/**
 * Cabeceras de seguridad en todas las respuestas.
 *
 * - La pagina solo corre los programas que vienen de este servidor (ni pegados dentro de un
 *   documento ni de otro sitio): si algun dia se colara HTML ajeno, no se ejecuta.
 * - No se puede meter dentro de otra pagina (nadie la disfraza para robar clics).
 * - El navegador no adivina tipos de archivo ni manda de donde viene a otros sitios.
 * - Por HTTPS, que siempre vuelva por HTTPS.
 */
@Component
@Order(Ordered.HIGHEST_PRECEDENCE)
public class FiltroCabeceras extends OncePerRequestFilter {

    /** La misma politica va en index.html (para la app, que carga la interfaz desde sus archivos). */
    static final String CSP = "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; "
            + "img-src 'self' data: blob: https: http:; font-src 'self' data:; media-src 'self' data: blob:; "
            + "connect-src 'self' https: http:; worker-src 'self'; manifest-src 'self'; object-src 'none'; "
            + "base-uri 'self'; form-action 'self'; frame-ancestors 'self'";

    /** ¿El navegador ve esta página en un origen de confianza (https o este mismo equipo)? */
    static boolean confiable(HttpServletRequest req) {
        if ("https".equalsIgnoreCase(req.getHeader("X-Forwarded-Proto")) || req.isSecure()) return true;
        String host = req.getHeader("Host");
        if (host == null) return false;
        host = host.trim().toLowerCase(java.util.Locale.ROOT).replaceAll(":\\d+$", "");
        return host.equals("localhost") || host.endsWith(".localhost") || host.startsWith("127.") || host.equals("[::1]");
    }

    @Override
    protected void doFilterInternal(HttpServletRequest req, HttpServletResponse res, FilterChain cadena) throws ServletException, IOException {
        res.setHeader("X-Content-Type-Options", "nosniff");
        res.setHeader("X-Frame-Options", "SAMEORIGIN");
        res.setHeader("Referrer-Policy", "no-referrer");
        // mudanza.html es la única que se comunica con una ventana de la otra dirección de NOM (js/mudanza.js).
        boolean mudanza = "/mudanza.html".equals(req.getRequestURI());
        // El navegador solo aplica esta cabecera en https o en este mismo equipo; por http en la red del salón la
        // ignora y deja un aviso en la consola, así que ahí no se manda.
        if (confiable(req)) res.setHeader("Cross-Origin-Opener-Policy", mudanza ? "unsafe-none" : "same-origin");
        res.setHeader("Cross-Origin-Resource-Policy", "same-site");
        res.setHeader("Permissions-Policy", "camera=(self), microphone=(), geolocation=(), payment=()");
        if (!req.getRequestURI().startsWith("/api/")) res.setHeader("Content-Security-Policy", CSP);
        // Los datos de la API nunca se guardan en la caché del navegador: ahí quedarían en claro (lo del
        // equipo vive cifrado en la bóveda) y, sin red, el navegador podría dar una copia vieja como si
        // fuera la del servidor (Firefox lo hacía al recargar sin conexión).
        else res.setHeader("Cache-Control", "no-store");
        // Las direcciones de un solo uso (unirse a un espacio, ver una proyección) no van en los buscadores.
        if (mudanza || req.getRequestURI().startsWith("/unirse") || req.getRequestURI().startsWith("/proyectar")) res.setHeader("X-Robots-Tag", "noindex");
        if ("https".equalsIgnoreCase(req.getHeader("X-Forwarded-Proto")) || req.isSecure())
            res.setHeader("Strict-Transport-Security", "max-age=15552000");
        cadena.doFilter(req, res);
    }
}
