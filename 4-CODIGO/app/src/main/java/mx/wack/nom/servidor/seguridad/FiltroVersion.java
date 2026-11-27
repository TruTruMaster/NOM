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
import mx.wack.nom.core.Version;

/**
 * Solo la misma version trabaja con este servidor (Version.misma): la pagina, la app de Android y otro NOM que copia
 * espacios mandan su version en la cabecera X-NOM-Version (el canal en vivo, en ?v=). Con otra version: 426 y el aviso;
 * la pagina lo trata como «sin conexion» y guarda lo hecho en el equipo. Despues de CORS, para que la app lea el aviso.
 */
@Component
@Order(Ordered.HIGHEST_PRECEDENCE + 3)
public class FiltroVersion extends OncePerRequestFilter {

    @Override
    protected void doFilterInternal(HttpServletRequest req, HttpServletResponse res, FilterChain cadena) throws ServletException, IOException {
        String ruta = req.getRequestURI();
        if (!ruta.startsWith("/api/") || Version.sinVersion(req.getMethod(), ruta)) { cadena.doFilter(req, res); return; }
        String suya = req.getHeader(Version.CABECERA);
        if (suya == null) suya = req.getParameter("v");
        if (Version.misma(suya)) { cadena.doFilter(req, res); return; }
        res.setStatus(426);
        res.setHeader("Cache-Control", "no-store");
        res.setContentType("application/json;charset=UTF-8");
        java.util.Map<String, Object> m = Version.distinta(suya);
        // (los textos son fijos y la versión de quien pide ya viene limpia en distinta(): sin comillas que escapar)
        String json = "{\"error\":\"" + m.get("error") + "\",\"version\":\"" + m.get("version") + "\",\"actualizar\":true}";
        res.getOutputStream().write(json.getBytes(java.nio.charset.StandardCharsets.UTF_8));
    }
}
