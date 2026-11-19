package mx.wack.nom.servidor.api;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.springframework.http.CacheControl;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.PutMapping;
import org.springframework.web.bind.annotation.RequestAttribute;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import jakarta.servlet.http.HttpServletRequest;
import mx.wack.nom.core.modelo.Usuario;
import mx.wack.nom.core.servicio.ErrorNom;
import mx.wack.nom.core.servicio.ServicioBinarios;
import mx.wack.nom.servidor.seguridad.FiltroSesion;

/**
 * Almacén de binarios de cada espacio (ver ServicioBinarios). Los bytes viajan tal cual (sin JSON ni base64):
 *   PUT  /api/espacios/{id}/binarios/{huella}       subir (cuerpo = los bytes; la huella es su SHA-256)
 *   GET  /api/espacios/{id}/binarios/{huella}       bajar (con caché permanente: un binario nunca cambia)
 *   POST /api/espacios/{id}/binarios/faltan         { huellas: [...] } → { faltan: [...] }
 *   GET  /api/ver/{token}/binarios/{huella}          sin sesión, con el enlace para ver
 *   GET  /api/proyeccion/{codigo}/binarios/{huella}  sin sesión, con el código del proyector
 */
@RestController
public class BinariosApi {

    private final ServicioBinarios binarios;

    public BinariosApi(ServicioBinarios binarios) { this.binarios = binarios; }

    @PutMapping(value = "/api/espacios/{id}/binarios/{huella}", consumes = MediaType.ALL_VALUE)
    public Map<String, Object> subir(@RequestAttribute(FiltroSesion.ATRIBUTO) Usuario u, @PathVariable String id, @PathVariable String huella,
                                     HttpServletRequest req) throws IOException {
        long max = binarios.maxBytes();
        if (req.getContentLengthLong() > max) throw new ErrorNom(ErrorNom.LLENO, "El archivo pasa de " + (max / (1024 * 1024)) + " MB", null, Map.of("tope", max));
        // (se lee hasta un byte más del tope: si llega más, es que pasa)
        byte[] datos = req.getInputStream().readNBytes((int) Math.min(max + 1, Integer.MAX_VALUE - 16));
        if (datos.length > max) throw new ErrorNom(ErrorNom.LLENO, "El archivo pasa de " + (max / (1024 * 1024)) + " MB", null, Map.of("tope", max));
        return binarios.subir(u, id, huella, datos);
    }

    @GetMapping("/api/espacios/{id}/binarios/{huella}")
    public ResponseEntity<byte[]> bajar(@RequestAttribute(FiltroSesion.ATRIBUTO) Usuario u, @PathVariable String id, @PathVariable String huella) {
        return conCache(binarios.bajar(u, id, huella), huella);
    }

    @PostMapping("/api/espacios/{id}/binarios/faltan")
    public Map<String, Object> faltan(@RequestAttribute(FiltroSesion.ATRIBUTO) Usuario u, @PathVariable String id, @RequestBody Map<String, Object> c) {
        Object hs = c == null ? null : c.get("huellas");
        List<String> lista = hs instanceof List<?> l ? l.stream().map(String::valueOf).toList() : List.of();
        return Map.of("faltan", binarios.faltan(u, id, lista), "max", binarios.maxBytes());
    }

    @GetMapping("/api/ver/{token}/binarios/{huella}")
    public ResponseEntity<byte[]> ver(@PathVariable String token, @PathVariable String huella) {
        return conCache(binarios.bajarPublico("ver", token, huella), huella);
    }

    @GetMapping("/api/proyeccion/{codigo}/binarios/{huella}")
    public ResponseEntity<byte[]> proyeccion(@PathVariable String codigo, @PathVariable String huella) {
        return conCache(binarios.bajarPublico("proyeccion", codigo, huella), huella);
    }

    /** Un binario nunca cambia (su nombre es su huella): se puede guardar para siempre, pero solo en ese equipo. */
    private static ResponseEntity<byte[]> conCache(byte[] b, String huella) {
        String h = huella.contains(".") ? huella.substring(0, huella.indexOf('.')) : huella;
        return ResponseEntity.ok()
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .cacheControl(CacheControl.maxAge(java.time.Duration.ofDays(365)).cachePrivate().immutable())
                .eTag("\"" + h + "\"")
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment")
                .header("X-Content-Type-Options", "nosniff")
                .body(b);
    }
}
