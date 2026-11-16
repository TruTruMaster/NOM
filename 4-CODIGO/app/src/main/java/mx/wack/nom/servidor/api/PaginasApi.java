package mx.wack.nom.servidor.api;

import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.nio.charset.StandardCharsets;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import jakarta.servlet.http.HttpServletRequest;
import mx.wack.nom.core.servicio.PaginasLimpias;

/**
 * Direcciones que se abren desde un codigo QR con la camara del telefono: /unirse?c=CODIGO y
 * /proyectar?c=CODIGO. Las dos muestran la misma pagina; ella lee el codigo.
 *
 * Fase 14: el sitio de ayuda y legal (/ayuda/… y /legal/…, HTML estatico). Cada pagina se muestra
 * en su direccion limpia, sin diagonal final y sin index.html (/legal/terminos); las formas viejas
 * (/legal/terminos/ y /legal/terminos/index.html) redirigen a ella. Los archivos guardan enlaces
 * relativos para la app; al servirlos se cambian por direcciones completas (PaginasLimpias).
 */
@Controller
public class PaginasApi {

    private static final MediaType HTML = new MediaType("text", "html", StandardCharsets.UTF_8);

    private final ResourceLoader archivos;
    private final String[] lugares;

    public PaginasApi(ResourceLoader archivos,
                      @Value("${spring.web.resources.static-locations:classpath:/static/}") String[] lugares) {
        this.archivos = archivos;
        this.lugares = lugares;
    }

    @GetMapping({ "/unirse", "/proyectar" })
    public String pagina() {
        return "forward:/index.html";
    }

    @GetMapping({ "/ayuda", "/legal", "/ayuda/{a:[a-z0-9-]+}", "/legal/{a:[a-z0-9-]+}", "/ayuda/{a:[a-z0-9-]+}/{b:[a-z0-9-]+}" })
    public ResponseEntity<String> limpia(HttpServletRequest req) throws IOException {
        String pagina = req.getRequestURI().substring(1);
        for (String lugar : lugares) {
            Resource r = archivos.getResource((lugar.endsWith("/") ? lugar : lugar + "/") + pagina + "/index.html");
            if (!r.isReadable()) continue;
            String html;
            try (InputStream en = r.getInputStream()) { html = new String(en.readAllBytes(), StandardCharsets.UTF_8); }
            return ResponseEntity.ok().contentType(HTML).body(PaginasLimpias.limpiar(html, pagina));
        }
        return ResponseEntity.status(HttpStatus.NOT_FOUND).contentType(HTML)
                .body("<!doctype html><html lang=\"es\"><meta charset=\"utf-8\"><title>No existe</title><p>Esa página no existe. <a href=\"/ayuda\">Ir a la ayuda</a></p></html>");
    }

    @GetMapping({ "/ayuda/", "/legal/", "/ayuda/{a:[a-z0-9-]+}/", "/legal/{a:[a-z0-9-]+}/", "/ayuda/{a:[a-z0-9-]+}/{b:[a-z0-9-]+}/",
                  "/ayuda/index.html", "/legal/index.html", "/ayuda/{a:[a-z0-9-]+}/index.html", "/legal/{a:[a-z0-9-]+}/index.html",
                  "/ayuda/{a:[a-z0-9-]+}/{b:[a-z0-9-]+}/index.html" })
    public ResponseEntity<Void> vieja(HttpServletRequest req) {
        String limpia = PaginasLimpias.limpiaDe(req.getRequestURI().substring(1));
        String consulta = req.getQueryString();
        return ResponseEntity.status(HttpStatus.MOVED_PERMANENTLY)
                .location(URI.create("/" + limpia + (consulta == null ? "" : "?" + consulta))).build();
    }
}
