package mx.wack.nom.core.servicio;

import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.stream.Collectors;


import mx.wack.nom.core.modelo.Usuario;

/**
 * Canal en tiempo real (CanalSync del lado web). Cada cliente abre un
 * EventSource por espacio y, si edita, otro por documento. Por aqui salen las
 * operaciones de los demas, la presencia y los avisos de accesos y versiones.
 *
 * Tambien aplica el RNF "maximo 50 personas a la vez por espacio" (antes 5) contando usuarios
 * distintos con un canal abierto en el espacio.
 *
 * El canal concreto (Canal) lo pone quien sirve la API: Spring en la computadora, el servidor de
 * la app en el teléfono. latido() se llama cada 20 s.
 */
public class Difusor {

    public static class Suscriptor {
        final Canal canal;
        final String usuarioId, nombre, color, espacioId, docId, cliente;
        /** Pantalla de proyeccion (5.5): recibe los cambios del documento, pero no es una persona
            del espacio (no cuenta para el maximo, ni sale en la presencia). */
        final String proyector;
        /** De una pantalla de proyeccion: quien presenta (si pierde el acceso, la pantalla se cierra). */
        final String presentador;
        /** La sesión con la que se abrió el canal (null en las pantallas de proyección): ver latido(). */
        final String token;
        volatile String bloque;
        volatile String modo;
        volatile Map<String, Object> cursor;

        Suscriptor(Canal c, Usuario u, String token, String espacioId, String docId, String cliente) {
            this.canal = c;
            this.token = token;
            this.usuarioId = u.id;
            this.nombre = u.nombre;
            this.color = u.color;
            this.espacioId = espacioId;
            this.docId = docId;
            this.cliente = cliente;
            this.proyector = null;
            this.presentador = null;
        }

        Suscriptor(Canal c, String espacioId, String docId, String codigo, String presentador) {
            this.canal = c;
            this.usuarioId = "proyector";
            this.nombre = "Proyector";
            this.color = "#6B6875";
            this.espacioId = espacioId;
            this.docId = docId;
            this.cliente = "proyector-" + codigo;
            this.proyector = codigo;
            this.presentador = presentador;
            this.token = null;
        }
    }

    /** RNF Espacios 1: personas conectadas a la vez por espacio (antes 5; 50 desde el 13-nov-2026). */
    public static final int MAX_PERSONAS = 50;
    /** Canales abiertos a la vez: por cuenta (varias pestañas y equipos), por pantalla de proyección y en todo el servidor. */
    static final int MAX_CANALES_CUENTA = 40, MAX_PROYECTORES = 20, MAX_CANALES = 5_000;

    private void exigirLugar(java.util.function.Predicate<Suscriptor> mismos, int tope, String mensaje) {
        if (subs.size() >= MAX_CANALES) throw new ErrorNom(ErrorNom.DEMASIADOS, "El servidor tiene demasiadas conexiones abiertas: intenta en un momento");
        int n = 0;
        for (Suscriptor x : subs) if (mismos.test(x) && ++n >= tope) throw new ErrorNom(ErrorNom.DEMASIADOS, mensaje);
    }

    private final List<Suscriptor> subs = new CopyOnWriteArrayList<>();
    private final int maxUsuarios;
    /** Quien deja de ver un documento (docId, usuarioId): ahi termina su visita (3.4). */
    private volatile java.util.function.BiConsumer<String, String> alSalirDeDocumento;

    public void alSalirDeDocumento(java.util.function.BiConsumer<String, String> f) { this.alSalirDeDocumento = f; }

    /** Quienes tienen abierto un documento ahora (sin contar proyectores). */
    public Set<String> usuariosEnDocumento(String docId) { return usuariosEnDocumento(docId, null); }

    /** Igual, sin contar el canal de un cliente (el que avisa que se va). */
    public Set<String> usuariosEnDocumento(String docId, String exceptoCliente) {
        return subs.stream().filter(s -> s.proyector == null && docId.equals(s.docId) && (exceptoCliente == null || !exceptoCliente.equals(s.cliente)))
                .map(s -> s.usuarioId).collect(Collectors.toSet());
    }

    public Difusor(int maxUsuarios) {
        this.maxUsuarios = maxUsuarios;
    }

    /** Abre el canal de un cliente en un espacio (y, si edita, en un documento). */
    public void suscribir(Usuario u, String espacioId, String docId, String cliente, Canal canal) {
        suscribir(u, null, espacioId, docId, cliente, canal);
    }

    /** Igual, con la sesión que lo abre: si esa sesión se cierra (salir, cambiar la contraseña, borrar la cuenta), el
        canal se cierra en el siguiente latido. */
    public void suscribir(Usuario u, String token, String espacioId, String docId, String cliente, Canal canal) {
        Set<String> conectados = usuariosEn(espacioId);
        if (!conectados.contains(u.id) && conectados.size() >= maxUsuarios) {
            throw new ErrorNom(ErrorNom.CONFLICTO, "El espacio está lleno: ya hay " + maxUsuarios
                    + " personas conectadas. Intenta en un momento.", null, Map.of("max", maxUsuarios));
        }
        exigirLugar(x -> x.proyector == null && u.id.equals(x.usuarioId), MAX_CANALES_CUENTA,
                "Tienes demasiadas pestañas o equipos abiertos a la vez: cierra alguno e intenta de nuevo");
        Suscriptor s = new Suscriptor(canal, u, token, espacioId, docId, cliente == null ? "" : cliente);
        subs.add(s);
        canal.alCerrar(() -> quitar(s));
        enviar(s, "hola", Map.of("ok", true));
        if (docId != null) presencia(docId);
        avisarEspacio(espacioId, "conectados", Map.of("usuarios", conectadosPublico(espacioId)));
    }

    /** Quien se va (cerró la pestaña o se cerró su canal): sale de la lista y los demás lo ven (una sola vez). */
    private void quitar(Suscriptor s) {
        if (!subs.remove(s)) return;
        if (s.docId != null) presencia(s.docId);
        avisarEspacio(s.espacioId, "conectados", Map.of("usuarios", conectadosPublico(s.espacioId)));
        java.util.function.BiConsumer<String, String> f = alSalirDeDocumento;
        if (s.docId != null && f != null) {
            try { f.accept(s.docId, s.usuarioId); } catch (RuntimeException ex) { /* la visita es solo informativa */ }
        }
    }

    /** Una pantalla de proyeccion para un documento (5.5). */
    public void suscribirProyector(String espacioId, String docId, String codigo, String presentador, Canal canal) {
        exigirLugar(x -> codigo.equals(x.proyector), MAX_PROYECTORES, "Ya hay demasiadas pantallas siguiendo esta presentación");
        Suscriptor s = new Suscriptor(canal, espacioId, docId, codigo, presentador);
        subs.add(s);
        canal.alCerrar(() -> subs.remove(s));
        enviar(s, "hola", Map.of("ok", true));
    }

    /** A las pantallas que siguen una presentacion: que diapositiva va, o que termino. */
    public void avisarProyectores(String codigo, String evento, Object datos) {
        for (Suscriptor s : subs) if (codigo.equals(s.proyector)) enviar(s, evento, datos);
    }

    public Set<String> usuariosEn(String espacioId) {
        return subs.stream().filter(s -> s.proyector == null && s.espacioId.equals(espacioId)).map(s -> s.usuarioId).collect(Collectors.toSet());
    }

    public boolean otroEditando(String docId, String usuarioId) {
        return subs.stream().anyMatch(s -> s.proyector == null && docId.equals(s.docId) && !s.usuarioId.equals(usuarioId));
    }

    public List<Map<String, Object>> conectadosPublico(String espacioId) {
        Map<String, Map<String, Object>> porUsuario = new LinkedHashMap<>();
        for (Suscriptor s : subs) {
            if (s.proyector != null || !s.espacioId.equals(espacioId)) continue;
            porUsuario.computeIfAbsent(s.usuarioId, k -> {
                Map<String, Object> m = new LinkedHashMap<>();
                m.put("id", s.usuarioId);
                m.put("nombre", s.nombre);
                m.put("color", s.color);
                return m;
            });
        }
        return new ArrayList<>(porUsuario.values());
    }

    /** Operaciones nuevas para todos los que tienen abierto el documento (incluido el autor: ignora las suyas por id de nodo). */
    public void ops(String docId, Object datos) { ops(docId, datos, datos); }

    /** ¿Alguien (una persona o una pantalla de proyección) tiene abierto este documento? */
    public boolean hayCanales(String docId) {
        for (Suscriptor s : subs) if (docId.equals(s.docId)) return true;
        return false;
    }

    /** ¿Hay alguna pantalla de proyección siguiendo este documento? */
    public boolean hayProyectores(String docId) {
        for (Suscriptor s : subs) if (s.proyector != null && docId.equals(s.docId)) return true;
        return false;
    }

    /** Igual, con lo que pueden ver las pantallas de proyección (sin comentarios ni notas; null: nada). */
    public void ops(String docId, Object datos, Object paraProyectores) {
        for (Suscriptor s : subs) {
            if (!docId.equals(s.docId)) continue;
            if (s.proyector == null) enviar(s, "ops", datos);
            else if (paraProyectores != null) enviar(s, "ops", paraProyectores);
        }
    }

    public void actualizarPresencia(String docId, String cliente, String bloque, String modo, Map<String, Object> cursor) {
        for (Suscriptor s : subs) {
            if (docId.equals(s.docId) && s.cliente.equals(cliente)) {
                s.bloque = bloque;
                s.modo = modo;
                s.cursor = cursor != null && cursor.size() <= 4 ? cursor : null;
            }
        }
        presencia(docId);
    }

    private void presencia(String docId) {
        List<Map<String, Object>> lista = new ArrayList<>();
        for (Suscriptor s : subs) {
            if (s.proyector != null || !docId.equals(s.docId)) continue;
            Map<String, Object> m = new LinkedHashMap<>();
            m.put("cliente", s.cliente);
            m.put("id", s.usuarioId);
            m.put("nombre", s.nombre);
            m.put("color", s.color);
            m.put("bloque", s.bloque);
            m.put("modo", s.modo);
            m.put("cursor", s.cursor);
            lista.add(m);
        }
        for (Suscriptor s : subs) if (s.proyector == null && docId.equals(s.docId)) enviar(s, "presencia", Map.of("usuarios", lista));
    }

    public void avisarDocumento(String docId, String evento, Object datos) {
        for (Suscriptor s : subs) if (docId.equals(s.docId)) enviar(s, evento, datos);
    }

    public void avisarEspacio(String espacioId, String evento, Object datos) {
        for (Suscriptor s : subs) if (s.proyector == null && s.espacioId.equals(espacioId) && s.docId == null) enviar(s, evento, datos);
    }

    /** Para todos los canales del espacio, incluidos los de documentos. */
    public void avisarTodoEspacio(String espacioId, String evento, Object datos) {
        for (Suscriptor s : subs) if (s.proyector == null && s.espacioId.equals(espacioId)) enviar(s, evento, datos);
    }

    /** 8.5 Un aviso para una persona, en todos los canales que tenga abiertos (espacios y documentos). */
    public void avisarUsuario(String usuarioId, String evento, Object datos) {
        for (Suscriptor s : subs) if (s.proyector == null && usuarioId.equals(s.usuarioId)) enviar(s, evento, datos);
    }

    /**
     * Cierra los canales abiertos en un documento de quienes ya no deben recibirlo (p. ej. al moverlo a un
     * espacio del que no son miembros). Las pantallas de proyección se cierran también: siguen un código
     * de quien presentaba en el espacio de antes.
     */
    public void cerrarDocumento(String docId, java.util.function.Predicate<String> sinAcceso) {
        for (Suscriptor s : subs) {
            if (!docId.equals(s.docId)) continue;
            if (s.proyector != null || sinAcceso.test(s.usuarioId)) {
                // (como expulsar: al cerrarse, el canal sale de la lista y se avisa la presencia)
                enviar(s, s.proyector != null ? "fin" : "expulsado", Map.of("motivo", "Este documento se movió a un espacio en el que no estás"));
                try { s.canal.cerrar(); } catch (RuntimeException ignorado) { /* ya estaba cerrado */ }
            }
        }
    }

    /** CU-17 3b: al revocar un acceso se cierra el canal de ese usuario (y las pantallas donde proyectaba). */
    public void expulsar(String espacioId, String usuarioId) {
        for (Suscriptor s : subs) {
            if (!s.espacioId.equals(espacioId)) continue;
            if (s.proyector == null && s.usuarioId.equals(usuarioId)) {
                enviar(s, "expulsado", Map.of("motivo", "El propietario retiró tu acceso a este espacio"));
                s.canal.cerrar();
            } else if (s.proyector != null && usuarioId.equals(s.presentador)) {
                enviar(s, "fin", Map.of("ok", true));
                s.canal.cerrar();
            }
        }
    }

    /** Sesiones que se revisan en un latido, como mucho, y el tiempo que se les da (comparten el hilo con el guardado). */
    static final int MAX_REVISIONES = 300;
    static final long MAX_REVISION_NS = 200_000_000L;

    /**
     * De las sesiones con canales abiertos, las que ya terminaron. Si son más de MAX_REVISIONES se revisan algunas al
     * azar (las demás, en los latidos siguientes) y nunca más de MAX_REVISION_NS: el latido no detiene el guardado.
     */
    private Set<String> sesionesTerminadas(java.util.function.Predicate<String> vigente) {
        List<String> tokens = new java.util.ArrayList<>(subs.stream()
                .filter(s -> s.proyector == null && s.token != null).map(s -> s.token).collect(Collectors.toSet()));
        if (tokens.size() > MAX_REVISIONES) { java.util.Collections.shuffle(tokens); tokens = tokens.subList(0, MAX_REVISIONES); }
        long fin = System.nanoTime() + MAX_REVISION_NS;
        Set<String> r = new java.util.HashSet<>();
        for (String t : tokens) {
            if (System.nanoTime() > fin) break;
            if (!sigueAbierta(vigente, t)) r.add(t);
        }
        return r;
    }

    private static boolean sigueAbierta(java.util.function.Predicate<String> vigente, String token) {
        // (si no se puede revisar, p. ej. la base tarda, el canal se queda: se revisa otra vez en el siguiente latido)
        try { return vigente.test(token); } catch (RuntimeException e) { return true; }
    }

    private void enviar(Suscriptor s, String evento, Object datos) {
        try {
            s.canal.enviar(evento, datos);
        } catch (IOException | IllegalStateException ex) {
            subs.remove(s);
        }
    }

    /** Cómo saber si una sesión sigue abierta (lo pone quien sirve la API; sin esto no se revisa). */
    private volatile java.util.function.Predicate<String> sesionVigente;

    public void revisarSesionesCon(java.util.function.Predicate<String> vigente) { this.sesionVigente = vigente; }

    /**
     * Latido: mantiene vivas las conexiones a traves de proxies y detecta clientes muertos. Además cierra los canales
     * cuya sesión ya terminó (salió, cambió la contraseña en otro equipo o borró la cuenta); la página, al cerrarse el
     * canal, pregunta y pide entrar otra vez.
     */
    public void latido() {
        java.util.function.Predicate<String> vigente = sesionVigente;
        Set<String> terminadas = vigente == null ? Set.of() : sesionesTerminadas(vigente);
        for (Suscriptor s : subs) {
            if (s.token != null && s.proyector == null && terminadas.contains(s.token)) {
                try { s.canal.cerrar(); } catch (RuntimeException ignorado) { /* ya estaba cerrado */ }
                quitar(s);
                continue;
            }
            try {
                s.canal.latido();
            } catch (IOException | IllegalStateException ex) {
                subs.remove(s);
            }
        }
    }
}
