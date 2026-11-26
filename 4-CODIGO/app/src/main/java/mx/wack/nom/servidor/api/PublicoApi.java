package mx.wack.nom.servidor.api;

import java.util.Map;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestAttribute;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.mvc.method.annotation.SseEmitter;

import mx.wack.nom.core.modelo.Usuario;
import mx.wack.nom.core.servicio.ServicioDocumentos;
import mx.wack.nom.core.servicio.ServicioEspacios;
import mx.wack.nom.servidor.seguridad.FiltroSesion;

/**
 * Rutas que se usan sin sesion (ver ConfigWeb): lo que dice un enlace para unirse, entrar con
 * apodo, comprobar un recibo y ver una presentacion con su codigo. "Entrar" con un enlace si
 * pide sesion.
 */
@RestController
public class PublicoApi {

    public record Apodo(String apodo) { }
    public record Alias(String id, String llave) { }

    private final ServicioEspacios espacios;
    private final ServicioDocumentos documentos;
    private final mx.wack.nom.core.servicio.Topes topes;

    public PublicoApi(ServicioEspacios espacios, ServicioDocumentos documentos, mx.wack.nom.core.servicio.Topes topes) {
        this.espacios = espacios;
        this.documentos = documentos;
        this.topes = topes;
    }

    private <T> T conTope(jakarta.servlet.http.HttpServletRequest req, java.util.function.Supplier<T> f) {
        return topes.conTopeNoExiste(mx.wack.nom.servidor.seguridad.Origen.de(req), f);
    }

    @GetMapping("/api/enlaces/{codigo}")
    public Map<String, Object> enlace(@PathVariable String codigo, jakarta.servlet.http.HttpServletRequest req) {
        return conTope(req, () -> espacios.infoEnlace(codigo));
    }

    @PostMapping("/api/enlaces/{codigo}/invitado")
    public Map<String, Object> invitado(@PathVariable String codigo, @RequestBody Apodo a, jakarta.servlet.http.HttpServletRequest req) {
        return conTope(req, () -> espacios.entrarComoInvitado(codigo, a == null ? null : a.apodo(), mx.wack.nom.servidor.seguridad.Origen.de(req)));
    }

    /** Fase 13: entrar con un alias al azar (o volver a ser el mismo, con la llave de este equipo). */
    @PostMapping("/api/enlaces/{codigo}/alias")
    public Map<String, Object> alias(@PathVariable String codigo, @RequestBody(required = false) Alias a, jakarta.servlet.http.HttpServletRequest req) {
        return conTope(req, () -> espacios.entrarConAlias(codigo, a == null ? null : a.id(), a == null ? null : a.llave(), mx.wack.nom.servidor.seguridad.Origen.de(req)));
    }

    /** Con sesion: la cuenta abierta entra al espacio con el permiso del enlace. */
    @PostMapping("/api/enlaces/{codigo}/entrar")
    public Map<String, Object> entrar(@RequestAttribute(FiltroSesion.ATRIBUTO) Usuario u, @PathVariable String codigo, jakarta.servlet.http.HttpServletRequest req) {
        return conTope(req, () -> espacios.entrarConEnlace(u, codigo));
    }

    @PostMapping("/api/recibos/verificar")
    public Map<String, Object> verificar(@RequestBody Map<String, Object> recibo) {
        return documentos.verificarRecibo(recibo);
    }

    /** 8.5 Ver un solo documento con el enlace que creo el propietario (sin cuenta, solo lectura). */
    @GetMapping("/api/ver/{token}")
    public Map<String, Object> ver(@PathVariable String token, jakarta.servlet.http.HttpServletRequest req) {
        return conTope(req, () -> documentos.verPublico(token));
    }

    @GetMapping("/api/proyeccion/{codigo}")
    public Map<String, Object> proyeccion(@PathVariable String codigo, jakarta.servlet.http.HttpServletRequest req) {
        return conTope(req, () -> documentos.verProyeccion(codigo));
    }

    @GetMapping(value = "/api/proyeccion/{codigo}/eventos", produces = "text/event-stream")
    public SseEmitter proyeccionEventos(@PathVariable String codigo, jakarta.servlet.http.HttpServletRequest req) {
        CanalSse canal = new CanalSse();
        conTope(req, () -> { documentos.suscribirProyeccion(codigo, canal); return null; });
        return canal.emisor;
    }
}
