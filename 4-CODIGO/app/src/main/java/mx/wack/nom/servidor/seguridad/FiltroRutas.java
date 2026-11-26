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
 * Antes que todo: solo se atienden direcciones escritas de forma simple (sin «;», «.» ni «..» como partes, «//», ni lo
 * codificado como «%2e», «%2f» o «%5c»), como hace el cortafuegos estricto de Spring Security; ninguna dirección de NOM
 * las necesita (400). Y solo los nombres de servidor de NOM (ver Hosts; 421).
 */
@Component
@Order(Ordered.HIGHEST_PRECEDENCE)
public class FiltroRutas extends OncePerRequestFilter {

    /** ¿La dirección (tal como llegó, sin la consulta) es de las que se entienden igual en todos lados? */
    static boolean valida(String uri) { return mx.wack.nom.core.servicio.Hosts.rutaValida(uri); }

    /** Los nombres de la nube que atiende este servidor (NOM_NUBE, NOM_CANONICO y NOM_NOMBRES, con y sin «www.»). */
    static final java.util.Set<String> CONOCIDOS = conocidos(System.getenv("NOM_NUBE"), System.getenv("NOM_CANONICO"), System.getenv("NOM_NOMBRES"));

    static java.util.Set<String> conocidos(String... fuentes) { return mx.wack.nom.core.servicio.Hosts.conocidos(fuentes); }

    static boolean hostPermitido(String host, java.util.Set<String> conocidos) { return mx.wack.nom.core.servicio.Hosts.permitido(host, conocidos); }

    @Override
    protected void doFilterInternal(HttpServletRequest req, HttpServletResponse res, FilterChain cadena) throws ServletException, IOException {
        if (!valida(req.getRequestURI())) {
            responder(res, 400, "Esa dirección no es válida");
            return;
        }
        if (!hostPermitido(req.getHeader("Host"), CONOCIDOS)) {
            responder(res, 421, "Este servidor no atiende ese nombre");
            return;
        }
        cadena.doFilter(req, res);
    }

    private static void responder(HttpServletResponse res, int codigo, String mensaje) throws IOException {
        res.setStatus(codigo);
        res.setContentType("application/json;charset=UTF-8");
        res.setHeader("Cache-Control", "no-store");
        res.getOutputStream().write(("{\"error\":\"" + mensaje + "\"}").getBytes(java.nio.charset.StandardCharsets.UTF_8));
    }
}
