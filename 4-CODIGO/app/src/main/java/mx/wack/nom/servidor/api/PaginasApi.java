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
import org.springframework.web.context.request.ServletWebRequest;

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
 *
 * Buscadores: la portada y cada pagina de ayuda y legal salen con su direccion principal, lo que
 * muestran las redes al compartirlas y sus datos estructurados; robots.txt y sitemap.xml se arman
 * con la direccion principal (NOM_CANONICO). Ver Buscadores.
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

    /** Atajos cortos para compartir (nom.qd.je/app, nom.qd.je/codigo): ver Atajos. */
    @GetMapping({ "/app", "/apk", "/codigo" })
    public ResponseEntity<Void> atajo(HttpServletRequest req) {
        String destino = mx.wack.nom.core.servicio.Atajos.destino(req.getRequestURI());
        if (destino == null) return ResponseEntity.status(HttpStatus.NOT_FOUND).build();
        return ResponseEntity.status(HttpStatus.FOUND).location(URI.create(destino))
                .cacheControl(org.springframework.http.CacheControl.noStore()).build();
    }

    /** La portada (y /index.html): la pagina de la app con lo que leen los buscadores y las redes. */
    @GetMapping({ "/", "/index.html" })
    public ResponseEntity<String> portada(HttpServletRequest req) throws IOException {
        String html = leer("index.html");
        if (html == null) return ResponseEntity.status(HttpStatus.NOT_FOUND).build();
        html = mx.wack.nom.core.servicio.PaginasLimpias.sinPoliticaEnMeta(html);   // (la política va en la cabecera)
        html = mx.wack.nom.core.servicio.PaginasLimpias.conVersion(html, mx.wack.nom.core.Version.APP);
        html = Buscadores.cabecera(html, Buscadores.base(req), "");
        // Débil (W/): Tomcat no comprime una respuesta con etiqueta fuerte, y la portada pesa ~70 KB.
        String etiqueta = "W/\"" + Integer.toHexString(html.hashCode()) + "\"";
        if (new ServletWebRequest(req).checkNotModified(etiqueta)) return ResponseEntity.status(HttpStatus.NOT_MODIFIED).eTag(etiqueta).build();
        return ResponseEntity.ok().contentType(HTML).eTag(etiqueta).cacheControl(org.springframework.http.CacheControl.noCache().noTransform()).body(html);
    }

    // robots.txt y sitemap.xml cambian con NOM_CANONICO: sin esto Cloudflare los guardaba 4 horas.
    @GetMapping(value = "/robots.txt", produces = "text/plain;charset=UTF-8")
    public ResponseEntity<String> robots(HttpServletRequest req) {
        return ResponseEntity.ok().cacheControl(org.springframework.http.CacheControl.noCache().noTransform()).body(Buscadores.robots(Buscadores.base(req)));
    }

    @GetMapping(value = "/sitemap.xml", produces = "application/xml;charset=UTF-8")
    public ResponseEntity<String> sitemap(HttpServletRequest req) throws IOException {
        String escrito = leer("sitemap.xml");
        if (escrito == null) return ResponseEntity.status(HttpStatus.NOT_FOUND).build();
        return ResponseEntity.ok().cacheControl(org.springframework.http.CacheControl.noCache().noTransform()).body(Buscadores.sitemap(escrito, Buscadores.base(req)));
    }

    // security.txt con la direccion principal (antes decia siempre la de dpdns aunque se entrara por otro nombre)
    @GetMapping(value = "/.well-known/security.txt", produces = "text/plain;charset=UTF-8")
    public ResponseEntity<String> seguridad(HttpServletRequest req) throws IOException {
        String escrito = leer(".well-known/security.txt");
        if (escrito == null) return ResponseEntity.status(HttpStatus.NOT_FOUND).build();
        return ResponseEntity.ok().cacheControl(org.springframework.http.CacheControl.noCache().noTransform()).body(escrito.replace(Buscadores.ESCRITA + "/", Buscadores.base(req) + "/"));
    }

    /** Un archivo de static/ (o de donde digan las ubicaciones configuradas) como texto; null si no esta. */
    private String leer(String archivo) throws IOException {
        for (String lugar : lugares) {
            Resource r = archivos.getResource((lugar.endsWith("/") ? lugar : lugar + "/") + archivo);
            if (!r.isReadable()) continue;
            try (InputStream en = r.getInputStream()) { return new String(en.readAllBytes(), StandardCharsets.UTF_8); }
        }
        return null;
    }

    @GetMapping({ "/ayuda", "/legal", "/ayuda/{a:[a-z0-9-]+}", "/legal/{a:[a-z0-9-]+}", "/ayuda/{a:[a-z0-9-]+}/{b:[a-z0-9-]+}" })
    public ResponseEntity<String> limpia(HttpServletRequest req) throws IOException {
        String pagina = req.getRequestURI().substring(1);
        for (String lugar : lugares) {
            Resource r = archivos.getResource((lugar.endsWith("/") ? lugar : lugar + "/") + pagina + "/index.html");
            if (!r.isReadable()) continue;
            String html;
            try (InputStream en = r.getInputStream()) { html = new String(en.readAllBytes(), StandardCharsets.UTF_8); }
            html = PaginasLimpias.limpiar(html, pagina);
            // La busqueda de la ayuda no es una pagina para los buscadores (robots.txt tampoco la deja leer).
            if (!pagina.equals("ayuda/buscar")) html = Buscadores.cabecera(html, Buscadores.base(req), pagina);
            return ResponseEntity.ok().contentType(HTML).cacheControl(org.springframework.http.CacheControl.noCache().noTransform()).body(html);
        }
        return ResponseEntity.status(HttpStatus.NOT_FOUND).contentType(HTML).cacheControl(org.springframework.http.CacheControl.noCache().noTransform())
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
