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
        /** Lo que falta mandarle, en orden (ver encolar); quién lo está mandando y desde cuándo escribe. */
        final java.util.concurrent.ConcurrentLinkedQueue<Object[]> fila = new java.util.concurrent.ConcurrentLinkedQueue<>();
        final java.util.concurrent.atomic.AtomicInteger enFila = new java.util.concurrent.atomic.AtomicInteger();
        final java.util.concurrent.atomic.AtomicBoolean vaciando = new java.util.concurrent.atomic.AtomicBoolean();
        volatile long escribiendoDesde;
        volatile boolean muerto;

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

    /** ¿Quien tiene este canal solo puede ver en el espacio (espacioId, usuarioId)? Esas personas reciben lo mismo que
        una pantalla de proyección: lo visible, sin comentarios, notas ni el original de Office. */
    private volatile java.util.function.BiPredicate<String, String> soloMira = (e, u) -> false;

    public void quienSoloMira(java.util.function.BiPredicate<String, String> f) { this.soloMira = f; }

    private boolean mira(Suscriptor s) {
        if (s.proyector != null) return true;
        try { return soloMira.test(s.espacioId, s.usuarioId); }
        catch (RuntimeException e) { return true; }   // (sin poder saberlo, la copia limpia)
    }

    /** Quienes tienen abierto un documento ahora (sin contar proyectores). */
    public Set<String> usuariosEnDocumento(String docId) { return usuariosEnDocumento(docId, null); }

    /** Igual, sin contar el canal de un cliente (el que avisa que se va). */
    public Set<String> usuariosEnDocumento(String docId, String exceptoCliente) {
        return subs.stream().filter(s -> s.proyector == null && docId.equals(s.docId) && (exceptoCliente == null || !exceptoCliente.equals(s.cliente)))
                .map(s -> s.usuarioId).collect(Collectors.toSet());
    }

    public Difusor(int maxUsuarios) { this(maxUsuarios, 8); }

    /**
     * hilos: cuántos hilos mandan lo de los canales (0 = se manda en el mismo hilo de quien avisa; solo para pruebas).
     * Cada canal tiene su propia fila: si alguien deja de leer (se cerró la tapa de la laptop, el celular perdió el
     * Wi-Fi), solo se atora su fila; los demás siguen recibiendo, quien escribe no espera y el guardado tampoco.
     */
    public Difusor(int maxUsuarios, int hilos) {
        this.maxUsuarios = maxUsuarios;
        this.envios = hilos <= 0 ? null : java.util.concurrent.Executors.newFixedThreadPool(hilos, r -> {
            Thread t = new Thread(r, "nom-envios");
            t.setDaemon(true);
            return t;
        });
    }

    private final java.util.concurrent.ExecutorService envios;
    /** Lo más que se le junta a un canal sin leer; y lo más que puede tardar una escritura atorada. Después, se cierra
        (la página vuelve a conectarse y recarga lo que le faltó). */
    static final int MAX_FILA = 400;
    static volatile long MAX_ESCRITURA_MS = 30_000L;
    private static final Object LATIDO = new Object(), CERRAR = new Object();

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
        Suscriptor s = new Suscriptor(canal, u, token, espacioId, docId, clienteValido(cliente));
        subs.add(s);
        canal.alCerrar(() -> { s.muerto = true; quitar(s); });
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
        canal.alCerrar(() -> { s.muerto = true; subs.remove(s); });
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

    /** ¿Alguien que solo mira (una pantalla de proyección o quien solo puede ver) tiene abierto este documento? */
    public boolean hayMirones(String docId) {
        for (Suscriptor s : subs) if (docId.equals(s.docId) && mira(s)) return true;
        return false;
    }

    /** Igual, con lo que puede ver quien solo mira (sin comentarios ni notas; null: nada). Si algo borrado vuelve a
        verse, a quien solo mira no le llega en el lote: se le pide que vuelva a cargar su copia limpia. */
    public void ops(String docId, Object datos, Object paraMirones) {
        for (Suscriptor s : subs) {
            if (!docId.equals(s.docId)) continue;
            if (!mira(s)) enviar(s, "ops", datos);
            else if (paraMirones instanceof Map<?, ?> m && Boolean.TRUE.equals(m.get("recargar")))
                enviar(s, s.proyector != null ? "version" : "recargar", Map.of("ok", true));
            else if (paraMirones != null) enviar(s, "ops", paraMirones);
        }
    }

    /** El identificador de una pestaña (lo pone la página: «d1a2b3c-x9y8»): letras, números y . _ : - (hasta 64). */
    public static String clienteValido(String cliente) {
        return cliente != null && cliente.matches("[A-Za-z0-9_.:-]{1,64}") ? cliente : "";
    }

    /** Lo que se reparte de un texto que manda la página (bloque, modo, partes del cursor): corto, o nada. */
    private static String corto(Object o, int max, String patron) {
        return o instanceof String t && t.length() <= max && t.matches(patron) ? t : null;
    }

    /** Presencia de un canal de usuarioId (solo los suyos: nadie mueve el cursor de otra persona). */
    private final Topes topes = new Topes();

    public void actualizarPresencia(String usuarioId, String docId, String cliente, String bloque, String modo, Map<String, Object> cursor) {
        String c = clienteValido(cliente);
        if (usuarioId == null || c.isEmpty()) return;
        // (cada aviso se reparte a todos los canales del documento: la página manda unos pocos por segundo)
        topes.exigir("p|" + usuarioId, 900, "Demasiados avisos de presencia");
        String b = corto(bloque, 64, "[^\\p{Cntrl}]+"), m = corto(modo, 24, "[A-Za-z_-]+");
        Map<String, Object> cur = null;
        if (cursor != null) {
            Object a = cursor.get("a"), f = cursor.get("f");
            String ca = a == null ? "" : corto(a, 200, "[^\\p{Cntrl}]*"), cf = f == null ? "" : corto(f, 200, "[^\\p{Cntrl}]*");
            if (ca != null && cf != null) { cur = new LinkedHashMap<>(); cur.put("a", ca); cur.put("f", cf); }
        }
        boolean cambio = false;
        for (Suscriptor s : subs) {
            if (docId.equals(s.docId) && usuarioId.equals(s.usuarioId) && s.cliente.equals(c)) {
                s.bloque = b;
                s.modo = m;
                s.cursor = cur;
                cambio = true;
            }
        }
        if (cambio) presencia(docId);
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
                cerrarEnOrden(s);
            }
        }
    }

    /** CU-17 3b: al revocar un acceso se cierra el canal de ese usuario (y las pantallas donde proyectaba). */
    public void expulsar(String espacioId, String usuarioId) {
        for (Suscriptor s : subs) {
            if (!s.espacioId.equals(espacioId)) continue;
            if (s.proyector == null && s.usuarioId.equals(usuarioId)) {
                enviar(s, "expulsado", Map.of("motivo", "El propietario retiró tu acceso a este espacio"));
                cerrarEnOrden(s);
            } else if (s.proyector != null && usuarioId.equals(s.presentador)) {
                enviar(s, "fin", Map.of("ok", true));
                cerrarEnOrden(s);
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

    private void enviar(Suscriptor s, String evento, Object datos) { encolar(s, new Object[] { evento, datos }); }

    /** Se cierra el canal después de mandarle lo que ya tenía en su fila (p. ej. el aviso de por qué se cierra). */
    private void cerrarEnOrden(Suscriptor s) { encolar(s, new Object[] { CERRAR, null }); }

    /** A la fila del canal; un hilo de envío la vacía en orden. Si la fila ya está llena, el canal no está leyendo: se corta. */
    private void encolar(Suscriptor s, Object[] x) {
        if (s.muerto) return;
        if (envios == null) { entregar(s, x); return; }
        if (s.enFila.incrementAndGet() > MAX_FILA) { s.enFila.decrementAndGet(); cortar(s); return; }
        s.fila.add(x);
        programarVaciado(s);
    }

    private void programarVaciado(Suscriptor s) {
        if (!s.vaciando.compareAndSet(false, true)) return;
        try { envios.execute(() -> vaciar(s)); }
        catch (java.util.concurrent.RejectedExecutionException e) { s.vaciando.set(false); }
    }

    private void vaciar(Suscriptor s) {
        try {
            Object[] x;
            while (!s.muerto && (x = s.fila.poll()) != null) {
                s.enFila.decrementAndGet();
                if (!entregar(s, x)) break;
            }
            if (s.muerto) { s.fila.clear(); s.enFila.set(0); }
        } finally {
            s.vaciando.set(false);
            if (!s.muerto && !s.fila.isEmpty()) programarVaciado(s);
        }
    }

    /** Escribe en el canal (en el hilo de envío). false si el canal ya no sirve: sale de la lista. */
    private boolean entregar(Suscriptor s, Object[] x) {
        if (x[0] == CERRAR) {
            s.muerto = true;
            try { s.canal.cerrar(); } catch (RuntimeException ignorado) { /* ya estaba cerrado */ }
            return false;
        }
        s.escribiendoDesde = System.currentTimeMillis();
        try {
            if (x[0] == LATIDO) s.canal.latido(); else s.canal.enviar((String) x[0], x[1]);
            return true;
        } catch (IOException | RuntimeException ex) {
            s.muerto = true;
            quitar(s);
            return false;
        } finally {
            s.escribiendoDesde = 0;
        }
    }

    /**
     * Un canal que no lee (su fila se llenó o lleva mucho escribiendo una sola cosa): sale de la lista y de la presencia
     * ya, y se cierra desde un hilo de envío (cerrarlo puede esperar a que termine la escritura atorada).
     */
    private void cortar(Suscriptor s) {
        if (s.muerto) return;
        s.muerto = true;
        quitar(s);
        Runnable cerrar = () -> { try { s.canal.cerrar(); } catch (RuntimeException ignorado) { /* ya estaba cerrado */ } };
        if (envios == null) cerrar.run();
        else try { envios.execute(cerrar); } catch (java.util.concurrent.RejectedExecutionException e) { cerrar.run(); }
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
        long ahora = System.currentTimeMillis();
        for (Suscriptor s : subs) {
            if (s.token != null && s.proyector == null && terminadas.contains(s.token)) {
                cortar(s);
                continue;
            }
            long desde = s.escribiendoDesde;
            if (desde > 0 && ahora - desde > MAX_ESCRITURA_MS) { cortar(s); continue; }
            encolar(s, new Object[] { LATIDO, null });
        }
    }

    /** Cierra todos los canales abiertos y detiene los hilos de envío al apagar el servidor. */
    public void detener() {
        for (Suscriptor s : subs) {
            s.muerto = true;
            try { s.canal.cerrar(); } catch (RuntimeException ignorado) { /* ya cerrado */ }
        }
        subs.clear();
        if (envios != null) envios.shutdownNow();
    }
}
