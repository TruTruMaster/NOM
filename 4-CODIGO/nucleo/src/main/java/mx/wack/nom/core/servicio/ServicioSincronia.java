package mx.wack.nom.core.servicio;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.InetAddress;
import java.net.URL;
import java.net.UnknownHostException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

import mx.wack.nom.core.crdt.Operacion;
import mx.wack.nom.core.modelo.Documento;
import mx.wack.nom.core.modelo.Espacio;
import mx.wack.nom.core.modelo.Permiso;
import mx.wack.nom.core.modelo.Usuario;

/**
 * 6.5 Copia de un espacio en otro servidor: el telefono del salon (que trabajo sin Internet) con
 * el servidor propio, y al reves. Como los dos usan el mismo CRDT, sincronizar es mandar
 * todo lo que uno tiene al otro (cada registro con su reloj y su nodo originales) y traer lo que
 * tiene el otro: los dos quedan con la union, sin importar el orden. No se borra nada en ningun
 * lado y los accesos de cada servidor no se tocan.
 *
 * Usa la misma API que los navegadores (con la cuenta del propietario en el otro servidor) y
 * solo el JDK para HTTP, asi corre igual en la computadora y en el telefono.
 */
public class ServicioSincronia {

    private static final Bitacora log = Bitacora.de(ServicioSincronia.class);
    private static final int LOTE = 300;
    /** Lo que copia la sincronizacion no lo tecleo quien sincroniza: quien escribio que va en la bitacora aparte. */
    private static final Map<String, Object> REACOMODO = Map.of("o", "r");

    private final ServicioEspacios espacios;
    private final ServicioDocumentos documentos;
    private final CodecJson json;
    private final Nodo nodo;

    /** El almacén de binarios de este servidor (las imágenes y el Office original de los documentos): lo que un documento
        menciona tiene que estar también en el otro servidor. null = sin almacén. */
    private volatile mx.wack.nom.core.almacen.RepositorioBinarios binarios;

    public void usarBinarios(mx.wack.nom.core.almacen.RepositorioBinarios repo) { this.binarios = repo; }

    public ServicioSincronia(ServicioEspacios espacios, ServicioDocumentos documentos, CodecJson json, Nodo nodo) {
        this.espacios = espacios;
        this.documentos = documentos;
        this.json = json;
        this.nodo = nodo;
    }

    /** Error del otro servidor (con su codigo HTTP). */
    static final class ErrorRemoto extends IOException {
        final int codigo;
        ErrorRemoto(int codigo, String m) { super(m); this.codigo = codigo; }
    }

    /** La dirección (o alguna de sus IP) queda fuera del alcance de quien configuró la copia. */
    static final class DireccionInterna extends IOException {
        DireccionInterna() { super("Esa dirección es de una red interna: la copia no se conecta ahí"); }
    }

    // ------------------------------------------------------------ a dónde se puede conectar

    /*
     * A dónde se conecta una copia: hasta donde llega quien la configura. Desde Internet, solo a Internet; desde una
     * red privada (el salón, la casa), también a esa red; y solo desde este mismo equipo, también a este equipo. El
     * alcance se guarda con la copia y se revisa en cada conexión, y las redirecciones no se siguen (se le dice a la
     * persona qué dirección escribir).
     */
    public static final String EQUIPO = "equipo", RED = "red", INTERNET = "internet";

    /** Hasta dónde llega quien pide, por su dirección de red (null = la petición salió de este mismo equipo). */
    public static String alcanceDe(String origen) {
        if (origen == null || origen.isBlank()) return EQUIPO;
        // (solo una IP escrita: es la de quien pide; nunca se busca un nombre)
        if (!origen.trim().matches("[0-9a-fA-F:.]{2,45}")) return INTERNET;
        try {
            // Una IP de este mismo equipo o de enlace local solo puede venir escrita en una cabecera de reenvío
            // (una petición que de verdad sale de aquí no la trae): no se le cree.
            return distancia(InetAddress.getByName(origen.trim())) == 1 ? RED : INTERNET;
        } catch (UnknownHostException | RuntimeException e) {
            return INTERNET;
        }
    }

    /**
     * Qué tan lejos queda una IP: 0 = este mismo equipo o una dirección especial (localhost, sin especificar,
     * enlace local como 169.254.169.254, multidifusión), 1 = una red privada (LAN, CGNAT/Tailscale, IPv6 ULA),
     * 2 = Internet.
     */
    static int distancia(InetAddress a) {
        byte[] b = a.getAddress();
        if (b.length == 16) {
            // ::a.b.c.d, ::ffff:a.b.c.d y 64:ff9b::a.b.c.d (NAT64): cuenta la IPv4 de adentro.
            boolean doce = true, diez = true;
            for (int i = 0; i < 12; i++) if (b[i] != 0) { doce = false; break; }
            for (int i = 0; i < 10; i++) if (b[i] != 0) { diez = false; break; }
            boolean mapeada = diez && (b[10] & 255) == 0xff && (b[11] & 255) == 0xff;
            boolean nat64 = (b[0] & 255) == 0 && (b[1] & 255) == 0x64 && (b[2] & 255) == 0xff && (b[3] & 255) == 0x9b;
            for (int i = 4; nat64 && i < 12; i++) if (b[i] != 0) nat64 = false;
            if (doce || mapeada || nat64) {
                try { return distancia(InetAddress.getByAddress(java.util.Arrays.copyOfRange(b, 12, 16))); }
                catch (UnknownHostException e) { return 0; }
            }
            if ((b[0] & 255) == 0x20 && (b[1] & 255) == 0x02) return 0;
        }
        if (a.isAnyLocalAddress() || a.isLoopbackAddress() || a.isLinkLocalAddress() || a.isMulticastAddress()) return 0;
        if (b.length == 4 && ((b[0] & 255) == 0 || (b[0] & 255) >= 224)) return 0;
        return ServicioFuentes.esPublica(a) ? 2 : 1;
    }

    /** ¿Puede una copia con ese alcance conectarse a esa IP? */
    static boolean permitida(InetAddress a, String alcance) {
        if (EQUIPO.equals(alcance)) return true;
        int d = distancia(a);
        return RED.equals(alcance) ? d >= 1 : d == 2;
    }

    /** El alcance de la copia que se hace en este hilo; sin poner, el más estricto (Internet). */
    private static final ThreadLocal<String> ALCANCE = new ThreadLocal<>();

    /** Lo más que se lee de una respuesta del otro servidor: un documento (25 MB como máximo) o un binario (48 MB). */
    static final int MAX_RESPUESTA = 32 * 1024 * 1024, MAX_BINARIO = 50 * 1024 * 1024;
    /** Tiempo total de cada petición y de la copia de un espacio entero (lo que falte se copia la próxima vez). */
    static final long MAX_MS_PETICION = 120_000, MAX_MS_BINARIO = 300_000, MAX_MS_ESPACIO = 15 * 60_000;
    /** Cuando la persona espera la respuesta (conectar o «copiar ahora»): lo demás sigue en la copia de cada 5 minutos. */
    static final long MAX_MS_ESPERA = 100_000;
    /** Documentos de allá que se emparejan en cada copia. */
    static final int MAX_DOCS_REMOTOS = 2_000;
    /** Hasta cuándo puede seguir la copia de este hilo. */
    private static final ThreadLocal<Long> HASTA = new ThreadLocal<>();

    /** Sin tiempo para seguir: la copia se corta y sigue la próxima vez. */
    static final class SinTiempo extends IOException {
        SinTiempo() { super("El otro NOM tardó demasiado en responder; se vuelve a intentar solo"); }
    }

    private static long limite(long porPeticion) {
        long l = System.currentTimeMillis() + porPeticion;
        Long h = HASTA.get();
        return h == null ? l : Math.min(l, h);
    }

    /** Abre la conexión solo si TODAS las IP del nombre quedan al alcance, y sin seguir redirecciones. */
    private static HttpURLConnection abrir(String direccion) throws IOException {
        Long h = HASTA.get();
        if (h != null && System.currentTimeMillis() > h) throw new SinTiempo();
        URL url = new URL(direccion);
        String host = url.getHost();
        if (host == null || host.isEmpty()) throw new IOException("La dirección no tiene servidor");
        if (host.startsWith("[") && host.endsWith("]")) host = host.substring(1, host.length() - 1);
        String alcance = ALCANCE.get() == null ? INTERNET : ALCANCE.get();
        InetAddress[] ips = InetAddress.getAllByName(host);
        if (ips.length == 0) throw new DireccionInterna();
        for (InetAddress a : ips) if (!permitida(a, alcance)) throw new DireccionInterna();
        // (Java guarda unos segundos lo que contestó el DNS: la conexión usa las mismas IP que se revisaron)
        HttpURLConnection c = (HttpURLConnection) url.openConnection();
        c.setInstanceFollowRedirects(false);
        return c;
    }

    /** A dónde manda una redirección (no se sigue: se le dice a la persona qué dirección escribir). */
    private static String destino(HttpURLConnection c) {
        String l = c.getHeaderField("Location");
        if (l == null || l.isBlank()) return "";
        return " (" + (l.length() > 120 ? l.substring(0, 120) + "…" : l) + "): escribe esa";
    }

    // ------------------------------------------------------------ configurar

    static String normalizar(String url) {
        String u = url == null ? "" : url.trim().replaceAll("/+$", "");
        if (u.isEmpty()) throw ErrorNom.invalido("Escribe la dirección del otro NOM", "servidor");
        if (!u.matches("(?i)^https?://.*")) {
            String host = u.split("[/:]")[0].toLowerCase(Locale.ROOT);
            boolean red = host.matches("\\d+(\\.\\d+){3}") || host.equals("localhost") || !host.contains(".") || host.endsWith(".local");
            u = (red ? "http://" : "https://") + u;
            if (red && !u.replaceFirst("^https?://", "").matches(".*:\\d+$")) u += ":8080";
        }
        return u;
    }

    /** El propietario conecta el espacio con su cuenta en otro servidor y se hace la primera copia. */
    public Map<String, Object> configurar(Usuario u, String espacioId, String servidor, String usuario, String contrasena) {
        // El teléfono que comparte: quien lo pide está en su misma red (el salón).
        return configurar(u, espacioId, servidor, usuario, contrasena, RED);
    }

    /** alcance: hasta dónde llega quien la configura (ver alcanceDe); la copia no llega más lejos. */
    @SuppressWarnings("unchecked")
    public Map<String, Object> configurar(Usuario u, String espacioId, String servidor, String usuario, String contrasena, String alcance) {
        Espacio e = espacios.requerir(espacioId);
        espacios.exigir(e, u, Permiso.PROPIETARIO);
        topes.exigir("conf|" + u.id, 10, "Muchos intentos de conectar seguidos: espera un minuto");
        String base = normalizar(servidor);
        if (base.length() > 200) throw ErrorNom.invalido("Esa dirección es demasiado larga", "servidor");
        if (usuario == null || usuario.isBlank() || contrasena == null || contrasena.isEmpty())
            throw ErrorNom.invalido("Escribe tu usuario y tu contraseña de allá", "usuario");
        if (usuario.trim().length() > 64 || contrasena.length() > 1024) throw ErrorNom.invalido("Usuario o contraseña demasiado largos", "usuario");
        String alc = EQUIPO.equals(alcance) || RED.equals(alcance) ? alcance : INTERNET;
        ALCANCE.set(alc);
        long ini = System.currentTimeMillis();
        HASTA.set(ini + MAX_MS_ESPERA / 2);
        try {
            Object info = pedir("GET", base + "/api/red/info", null, null);
            if (!(info instanceof Map<?, ?> mi) || !(mi.get("nom") instanceof Number nom) || nom.intValue() != 1)
                throw ErrorNom.invalido("En esa dirección no hay un NOM", "servidor");
            if (String.valueOf(mi.get("servidor")).equals(nodo.id())) throw ErrorNom.invalido("Esa dirección es este mismo NOM", "servidor");
            Map<String, Object> ses = (Map<String, Object>) pedir("POST", base + "/api/cuentas/sesion", null, Map.of("nombre", usuario.trim(), "contrasena", contrasena));
            String token = String.valueOf(ses == null ? null : ses.get("token"));
            if (!token.matches("[A-Za-z0-9_-]{16,128}")) throw ErrorNom.invalido("El otro NOM respondió algo que no se entiende", "servidor");
            // Un espacio con el mismo nombre (si ya existe y lo puedo editar) o uno nuevo.
            String remoto = null;
            for (Object o : (List<Object>) pedir("GET", base + "/api/espacios", token, null)) {
                Map<String, Object> x = (Map<String, Object>) o;
                if (e.nombre.equalsIgnoreCase(String.valueOf(x.get("nombre"))) && !"LECTURA".equals(String.valueOf(x.get("permiso")))) { remoto = String.valueOf(x.get("id")); break; }
            }
            if (remoto == null) remoto = String.valueOf(((Map<String, Object>) pedir("POST", base + "/api/espacios", token, Map.of("nombre", e.nombre))).get("id"));
            if (!remoto.matches("[A-Za-z0-9_-]{1,64}")) throw ErrorNom.invalido("El otro NOM respondió algo que no se entiende", "servidor");
            Espacio.Sincronia s = new Espacio.Sincronia();
            s.servidor = base;
            s.usuario = usuario.trim();
            s.token = token;
            s.espacioRemoto = remoto;
            s.alcance = alc;
            e.sincronia = s;
            espacios.guardar(e);
        } catch (DireccionInterna ex) {
            throw new ErrorNom(ErrorNom.PROHIBIDO, ex.getMessage(), "servidor");
        } catch (ErrorRemoto ex) {
            throw ErrorNom.invalido(ex.codigo == 401 || ex.codigo == 423 ? "El otro NOM no aceptó tu usuario o contraseña: " + ex.getMessage() : "El otro NOM respondió: " + ex.getMessage(), ex.codigo == 401 ? "contrasena" : null);
        } catch (IOException ex) {
            throw ErrorNom.invalido("No se pudo conectar con " + base + " (¿hay Internet o estás en el mismo Wi-Fi?)", "servidor");
        } finally {
            ALCANCE.remove();
            HASTA.remove();
        }
        return sincronizarEspacio(e, Math.max(10_000, ini + MAX_MS_ESPERA - System.currentTimeMillis()));
    }

    private final Topes topes = new Topes();

    public Map<String, Object> sincronizar(Usuario u, String espacioId) {
        Espacio e = espacios.requerir(espacioId);
        espacios.exigir(e, u, Permiso.PROPIETARIO);
        if (e.sincronia == null) throw ErrorNom.invalido("Este espacio no tiene copia en otro NOM", null);
        // «Copiar ahora» trae y manda todo el espacio: unas pocas veces por minuto (la copia sola va cada 5 minutos).
        topes.exigir(espacioId, 6, "Se acaba de copiar este espacio: espera unos segundos");
        return sincronizarEspacio(e, MAX_MS_ESPERA);
    }

    public Map<String, Object> quitar(Usuario u, String espacioId) {
        Espacio e = espacios.requerir(espacioId);
        espacios.exigir(e, u, Permiso.PROPIETARIO);
        e.sincronia = null;
        espacios.guardar(e);
        return espacios.resumen(e, u);
    }

    /** Cada 5 minutos: los espacios conectados se copian si hay red (si no, se intenta luego). */
    public void cadaRato() {
        for (Espacio e : espacios.todos()) {
            if (e.sincronia == null || e.sincronia.token == null) continue;
            // (todas las copias comparten un hilo: cada espacio tiene a lo más 2 minutos por vuelta; lo que falte, en la siguiente)
            try { sincronizarEspacio(e, Math.min(MAX_MS_ESPACIO, 120_000)); } catch (RuntimeException ex) { log.aviso("Sincronización del espacio {}: {}", e.id, ex.getMessage()); }
        }
    }

    // ------------------------------------------------------------ copiar

    /** Fase 13, parte B: el reloj mas alto de cada nodo en un estado tal como llega en JSON. */
    @SuppressWarnings("unchecked")
    static Map<String, Long> vector(Map<String, Object> estado) {
        Map<String, Long> v = new LinkedHashMap<>();
        if (estado == null) return v;
        if (estado.get("bloques") instanceof Map<?, ?> bl)
            for (Object campos : bl.values()) if (campos instanceof Map<?, ?> cs)
                for (Object reg : cs.values()) if (reg instanceof Map<?, ?> rg && rg.get("n") != null && rg.get("t") instanceof Number t)
                    v.merge(String.valueOf(rg.get("n")), t.longValue(), Math::max);
        if (estado.get("seq") instanceof Map<?, ?> sq)
            for (Object reg : sq.values()) if (reg instanceof Map<?, ?> rg)
                for (String k : new String[] { "ins", "vis", "fmt" }) if (rg.get(k) instanceof List<?> l)
                    for (Object o : l) if (o instanceof Map<?, ?> en && en.get("n") != null && en.get("t") instanceof Number t) {
                        long fin = t.longValue();
                        if ("ins".equals(k) && en.get("e") == null && en.get("x") instanceof String x) fin += Math.max(0, x.codePointCount(0, x.length()) - 1);
                        v.merge(String.valueOf(en.get("n")), fin, Math::max);
                    }
        return v;
    }

    /** B.2/B.6: el telefono estuvo sin conexion con el otro servidor mas de 2 minutos. */
    static final long UMBRAL_RAMA = Long.getLong("nom.umbralRama", 2 * 60_000L);   // solo las pruebas lo cambian

    /** Los alias (y cuentas) que escribieron en un documento de aqui desde la ultima copia. */
    private String quienes(List<Map<String, Object>> filas) {
        java.util.LinkedHashSet<String> n = new java.util.LinkedHashSet<>();
        for (Map<String, Object> f : filas) if (f.get("usuario") != null && n.size() < 6) n.add(String.valueOf(f.get("usuario")));
        return String.join(", ", n);
    }

    /** Manda a alla quien escribio que en este documento (desde la ultima vez). */
    private void mandarActividad(Espacio.Sincronia s, Documento d, String rid) throws IOException {
        long hasta = s.actividadHasta.getOrDefault(d.id, 0L);
        List<Map<String, Object>> filas = documentos.actividadDesde(d.id, hasta);
        if (filas.isEmpty()) return;
        Map<String, Object> c = new LinkedHashMap<>();
        c.put("filas", filas);
        c.put("via", "teléfono de " + (s.usuario == null ? "su dueño" : s.usuario));
        try {
            pedir("POST", s.servidor + "/api/documentos/" + rid + "/actividad-externa", s.token, c);
        } catch (ErrorRemoto ex) {
            if (ex.codigo != 404) throw ex;   // un servidor de antes no tiene esta ruta: se omite
        }
        long ultima = hasta;
        for (Map<String, Object> f : filas) if (f.get("t") instanceof Number t) ultima = Math.max(ultima, t.longValue());
        s.actividadHasta.put(d.id, ultima);
    }

    /** Una copia a la vez por espacio (si ya hay una en curso, no se espera: se avisa). */
    private final Map<String, java.util.concurrent.locks.ReentrantLock> enCurso = new java.util.concurrent.ConcurrentHashMap<>();

    private Map<String, Object> sincronizarEspacio(Espacio e, long maxMs) {
        java.util.concurrent.locks.ReentrantLock c = enCurso.computeIfAbsent(e.id, k -> new java.util.concurrent.locks.ReentrantLock());
        if (!c.tryLock()) {
            Map<String, Object> r = new LinkedHashMap<>();
            r.put("estado", "en-curso");
            r.put("error", "Ya se está copiando este espacio: espera un momento");
            r.put("ultima", e.sincronia == null ? 0 : e.sincronia.ultima);
            r.put("subidas", 0);
            r.put("bajadas", 0);
            r.put("documentos", e.sincronia == null || e.sincronia.docs == null ? 0 : e.sincronia.docs.size());
            return r;
        }
        try { return copiarEspacio(e, maxMs); } finally { c.unlock(); }
    }

    @SuppressWarnings("unchecked")
    private Map<String, Object> copiarEspacio(Espacio e, long maxMs) {
        Espacio.Sincronia s = e.sincronia;
        Map<String, Object> r = new LinkedHashMap<>();
        int subidas = 0, bajadas = 0, nuevosAqui = 0, nuevosAlla = 0, ramas = 0;
        if (s.bases == null) s.bases = new LinkedHashMap<>();
        if (s.actividadHasta == null) s.actividadHasta = new LinkedHashMap<>();
        long ahora = System.currentTimeMillis();
        boolean estuvoSinRed = s.sinRedDesde > 0 && ahora - s.sinRedDesde > UMBRAL_RAMA;
        // Cada conexión de esta copia se revisa con el alcance de quien la configuró (una copia que no lo guarda se
        // toma como de la red local, que es donde se usaban).
        ALCANCE.set(s.alcance == null ? RED : s.alcance);
        HASTA.set(ahora + maxMs);
        try {
            String base = s.servidor, tok = s.token, rem = s.espacioRemoto;
            List<Object> remotos = (List<Object>) pedir("GET", base + "/api/espacios/" + rem + "/documentos", tok, null);
            Map<String, Map<String, Object>> porId = new LinkedHashMap<>();
            if (remotos == null) remotos = List.of();
            for (Object o : remotos) {
                if (porId.size() >= MAX_DOCS_REMOTOS) break;
                if (o instanceof Map<?, ?> x0 && String.valueOf(x0.get("id")).matches("[A-Za-z0-9_-]{1,64}")) { Map<String, Object> x = (Map<String, Object>) x0; porId.put(String.valueOf(x.get("id")), x); }
            }
            Set<String> emparejados = new HashSet<>(s.docs.values());
            // Documentos nuevos de alla: se crean aqui.
            for (var x : porId.values()) {
                String rid = String.valueOf(x.get("id"));
                if (emparejados.contains(rid)) continue;
                Documento d = documentos.crearExterno(e.id, String.valueOf(x.get("nombre")), String.valueOf(x.get("modo")));
                s.docs.put(d.id, rid);
                emparejados.add(rid);
                nuevosAqui++;
            }
            for (Documento d : documentos.listarInterno(e.id)) {
                String rid = s.docs.get(d.id);
                if (rid == null) {
                    // Documento nuevo de aqui: se crea alla.
                    Map<String, Object> cuerpo = new LinkedHashMap<>();
                    cuerpo.put("nombre", d.nombre);
                    cuerpo.put("modo", d.modo == null ? "TEXTO" : d.modo.name());
                    cuerpo.put("ops", List.of());
                    rid = String.valueOf(((Map<String, Object>) pedir("POST", base + "/api/espacios/" + rem + "/documentos", tok, cuerpo)).get("id"));
                    s.docs.put(d.id, rid);
                    nuevosAlla++;
                } else if (!porId.containsKey(rid)) {
                    // Lo borraron alla: aqui se conserva, ya sin pareja.
                    s.docs.remove(d.id);
                    s.copiados.remove(d.id);
                    continue;
                }
                long remMod = porId.containsKey(rid) ? ((Number) porId.get(rid).getOrDefault("modificado", 0L)).longValue() : -1;
                boolean subir = !Long.valueOf(d.modificado).equals(s.copiados.get(d.id));
                boolean bajar = !Long.valueOf(remMod).equals(s.remotos.get(rid));
                if (!subir && !bajar) continue;
                // Fase 13, B.6: el telefono estuvo sin conexion y alla tambien cambio: rama del telefono
                // (el original de alla no se toca; este documento sigue en la rama).
                if (subir && bajar && estuvoSinRed && s.bases.containsKey(d.id)) {
                    Map<String, Object> cd = null;
                    try { cd = (Map<String, Object>) pedir("POST", base + "/api/documentos/" + rid + "/cambios-desde", tok, Map.of("base", s.bases.get(d.id), "nodos", List.of(), "desde", s.sinRedDesde)); }
                    catch (ErrorRemoto ex) { if (ex.codigo != 404) throw ex; }
                    if (cd != null && Boolean.TRUE.equals(cd.get("otros"))) {
                        List<Map<String, Object>> filas = documentos.actividadDesde(d.id, s.actividadHasta.getOrDefault(d.id, 0L));
                        String quien = quienes(filas);
                        String nombre = d.nombre + " — rama del teléfono de " + (s.usuario == null ? "su dueño" : s.usuario) + (quien.isEmpty() ? "" : " (" + quien + ")");
                        if (nombre.length() > 120) nombre = nombre.substring(0, 119) + "…";
                        Map<String, Object> cuerpo = new LinkedHashMap<>();
                        cuerpo.put("nombre", nombre);
                        cuerpo.put("ops", operaciones((Map<String, Object>) json.leer(new String(json.escribir(documentos.estadoParaCopiar(d.id)), StandardCharsets.UTF_8))));
                        cuerpo.put("base", s.bases.get(d.id));
                        cuerpo.put("nodos", List.of());
                        cuerpo.put("motivo", "telefono");
                        cuerpo.put("desde", s.sinRedDesde);
                        Map<String, Object> rama = (Map<String, Object>) pedir("POST", base + "/api/documentos/" + rid + "/ramas", tok, cuerpo);
                        String ramaId = String.valueOf(rama.get("id"));
                        mandarActividad(s, d, ramaId);
                        s.docs.put(d.id, ramaId);
                        s.copiados.put(d.id, documentos.requerir(d.id).modificado);
                        s.remotos.put(ramaId, rama.get("modificado") instanceof Number n ? n.longValue() : 0L);
                        s.bases.remove(d.id);
                        ramas++;
                        continue;
                    }
                }
                if (subir) {
                    Map<String, Object> estadoAqui = (Map<String, Object>) json.leer(new String(json.escribir(documentos.estadoParaCopiar(d.id)), StandardCharsets.UTF_8));
                    subirBinarios(e.id, base, rem, tok, estadoAqui);
                    List<Operacion> ops = operaciones(estadoAqui);
                    for (int i = 0; i < ops.size(); i += LOTE) {
                        Map<String, Object> lote = new LinkedHashMap<>();
                        lote.put("cliente", "sincronia-" + nodo.id());
                        lote.put("ops", ops.subList(i, Math.min(ops.size(), i + LOTE)));
                        lote.put("dispositivo", "Sincronización");
                        Map<String, Object> res = (Map<String, Object>) pedir("POST", base + "/api/documentos/" + rid + "/ops", tok, lote);
                        subidas += ((Number) res.getOrDefault("aplicadas", 0)).intValue();
                    }
                }
                if (subir) mandarActividad(s, d, rid);
                Map<String, Object> abierto = (Map<String, Object>) pedir("GET", base + "/api/documentos/" + rid, tok, null);
                bajarBinarios(e.id, base, rem, tok, abierto.get("estado"));
                bajadas += documentos.fusionarExterno(d.id, operaciones((Map<String, Object>) abierto.get("estado")), "Sincronización con " + s.servidor.replaceFirst("^https?://", ""));
                s.copiados.put(d.id, documentos.requerir(d.id).modificado);
                Object doc = abierto.get("documento");
                s.remotos.put(rid, doc instanceof Map<?, ?> dm && dm.get("modificado") instanceof Number n ? n.longValue() : remMod);
                // La base para la proxima vez: lo que alla tiene ahora (ya con lo de aqui).
                s.bases.put(d.id, ServicioDocumentos.baseAcotada(vector((Map<String, Object>) abierto.get("estado"))));
            }
            s.ultima = System.currentTimeMillis();
            s.estado = "ok";
            s.error = null;
            s.sinRedDesde = 0;
        } catch (DireccionInterna ex) {
            s.estado = "error";
            s.error = ex.getMessage();
        } catch (ErrorRemoto ex) {
            s.estado = ex.codigo == 401 ? "sesion" : "error";
            s.error = ex.codigo == 401 ? "La sesión en el otro NOM venció: vuelve a conectar el espacio." : ex.getMessage();
        } catch (IOException ex) {
            s.estado = "sin-red";
            s.error = "Sin conexión con el otro NOM; se vuelve a intentar solo.";
            if (s.sinRedDesde == 0) s.sinRedDesde = ahora;
        } finally {
            ALCANCE.remove();
            HASTA.remove();
        }
        if (s.error != null) s.error = Textos.cortar(Textos.limpio(s.error), 300);
        espacios.guardar(e);
        r.put("estado", s.estado);
        r.put("error", s.error);
        r.put("ultima", s.ultima);
        r.put("subidas", subidas);
        r.put("bajadas", bajadas);
        r.put("nuevosAqui", nuevosAqui);
        r.put("nuevosAlla", nuevosAlla);
        r.put("ramas", ramas);
        r.put("documentos", s.docs.size());
        return r;
    }

    /**
     * Operaciones que recrean un estado (como lo da GET /api/documentos/{id}): un registro por
     * campo con su reloj y nodo originales, y cada tramo de texto, visibilidad y formato. Sus ids
     * son fijos, asi mandarlos dos veces no cambia nada.
     */
    @SuppressWarnings("unchecked")
    static List<Operacion> operaciones(Map<String, Object> estado) {
        List<Operacion> ops = new ArrayList<>();
        if (estado == null) return ops;
        Object bl = estado.get("bloques");
        if (bl instanceof Map<?, ?> bloques) {
            for (var b : ((Map<String, Object>) bloques).entrySet()) {
                if (!(b.getValue() instanceof Map<?, ?> campos)) continue;
                Map<String, Map<String, Object>> grupos = new LinkedHashMap<>();
                Map<String, long[]> relojes = new LinkedHashMap<>();
                Map<String, String> nodos = new LinkedHashMap<>();
                for (var c : ((Map<String, Object>) campos).entrySet()) {
                    if (!(c.getValue() instanceof Map<?, ?> reg) || !(reg.get("t") instanceof Number t) || reg.get("n") == null) continue;
                    String k = t.longValue() + "|" + reg.get("n");
                    grupos.computeIfAbsent(k, z -> new LinkedHashMap<>()).put(c.getKey(), reg.get("v"));
                    relojes.put(k, new long[] { t.longValue() });
                    nodos.put(k, String.valueOf(reg.get("n")));
                }
                for (var g : grupos.entrySet()) {
                    List<String> ks = new ArrayList<>(g.getValue().keySet());
                    for (int i = 0, parte = 0; i < ks.size(); i += 300, parte++) {
                        Map<String, Object> trozo = new LinkedHashMap<>();
                        for (String k : ks.subList(i, Math.min(ks.size(), i + 300))) trozo.put(k, g.getValue().get(k));
                        long t = relojes.get(g.getKey())[0];
                        String n = nodos.get(g.getKey());
                        ops.add(new Operacion("sx:" + n + ":" + t + ":" + b.getKey() + (parte > 0 ? ":" + parte : ""), b.getKey(), trozo, t, n, null, REACOMODO));
                    }
                }
            }
        }
        Object sq = estado.get("seq");
        if (sq instanceof Map<?, ?> seqs) {
            for (var b : ((Map<String, Object>) seqs).entrySet()) {
                if (!(b.getValue() instanceof Map<?, ?> reg)) continue;
                for (String lista : new String[] { "ins", "vis", "fmt" }) {
                    if (!(reg.get(lista) instanceof List<?> l)) continue;
                    String k = lista.equals("ins") ? "i" : lista.equals("vis") ? "v" : "f";
                    for (Object o : l) {
                        if (!(o instanceof Map<?, ?> en) || !(en.get("t") instanceof Number t) || en.get("n") == null) continue;
                        Map<String, Object> s = new LinkedHashMap<>((Map<String, Object>) en);
                        s.remove("t");
                        s.remove("n");
                        s.put("k", k);
                        String n = String.valueOf(en.get("n"));
                        ops.add(new Operacion("sx:" + k + ":" + n + ":" + t.longValue() + ":" + b.getKey(), b.getKey(), Map.of(), t.longValue(), n, s, REACOMODO));
                    }
                }
            }
        }
        ops.sort(java.util.Comparator.comparingLong(Operacion::t));
        return ops;
    }

    // ------------------------------------------------------------ HTTP (solo el JDK)

    // ------------------------------------------------------------ almacén de binarios
    private static final java.util.regex.Pattern REF = java.util.regex.Pattern.compile("nom:bin/([0-9a-f]{64})");

    private static void refsEn(Object v, Set<String> s) {
        if (v instanceof CharSequence cs) { if (cs.length() >= 72 && cs.toString().contains("nom:bin/")) { var m = REF.matcher(cs); while (m.find()) s.add(m.group(1)); } }
        else if (v instanceof Map<?, ?> m) m.values().forEach(x -> refsEn(x, s));
        else if (v instanceof Iterable<?> it) it.forEach(x -> refsEn(x, s));
    }

    /** Antes de mandar un documento: los binarios que menciona y que el otro servidor no tiene, se le suben. */
    @SuppressWarnings("unchecked")
    private void subirBinarios(String aqui, String base, String rem, String tok, Object estado) throws IOException {
        var repo = binarios;
        if (repo == null || estado == null) return;
        Set<String> refs = new java.util.TreeSet<>();
        refsEn(estado, refs);
        refs.removeIf(h -> !repo.existe(aqui, h));
        if (refs.isEmpty()) return;
        Object r;
        try { r = pedir("POST", base + "/api/espacios/" + rem + "/binarios/faltan", tok, Map.of("huellas", new ArrayList<>(refs))); }
        catch (ErrorRemoto ex) { if (ex.codigo == 404 || ex.codigo == 405) return; throw ex; }   // el otro no tiene almacén
        Object f = r instanceof Map<?, ?> m ? m.get("faltan") : null;
        if (!(f instanceof List<?> faltan)) return;
        for (Object h : faltan) {
            byte[] b = repo.leer(aqui, String.valueOf(h));
            if (b != null) pedirBytes("PUT", base + "/api/espacios/" + rem + "/binarios/" + h, tok, b);
        }
    }

    /** Después de traer un documento: los binarios que menciona y que aquí no están, se bajan del otro servidor. */
    private void bajarBinarios(String aqui, String base, String rem, String tok, Object estado) throws IOException {
        var repo = binarios;
        if (repo == null || estado == null) return;
        Set<String> refs = new java.util.TreeSet<>();
        refsEn(estado, refs);
        Espacio esp = espacios.requerir(aqui);
        for (String h : refs) {
            if (repo.existe(aqui, h)) continue;
            documentos.revisarCuotaDe(esp);
            try {
                byte[] b = pedirBytes("GET", base + "/api/espacios/" + rem + "/binarios/" + h, tok, null);
                // (se guarda solo si llegó lo que dice su huella)
                if (b != null && b.length > 0 && ServicioBinarios.huellaDe(b).equals(h)) { repo.guardar(aqui, h, b); documentos.olvidarUso(esp.propietarioId); }
            } catch (ErrorRemoto ex) { if (ex.codigo != 404) throw ex; }
        }
    }

    private byte[] pedirBytes(String metodo, String url, String token, byte[] cuerpo) throws IOException {
        HttpURLConnection c = abrir(url);
        try {
            c.setRequestMethod(metodo);
            c.setConnectTimeout(8_000);
            c.setReadTimeout(60_000);
            if (token != null) c.setRequestProperty("Authorization", "Bearer " + token);
            if (cuerpo != null) {
                c.setDoOutput(true);
                c.setRequestProperty("Content-Type", "application/octet-stream");
                c.setFixedLengthStreamingMode(cuerpo.length);
                try (OutputStream o = c.getOutputStream()) { o.write(cuerpo); }
            }
            int codigo = c.getResponseCode();
            if (codigo >= 300 && codigo < 400) throw new ErrorRemoto(codigo, "El otro servidor manda a otra dirección" + destino(c));
            long hasta = limite(MAX_MS_BINARIO);
            if (codigo >= 400) { leer(c.getErrorStream(), 64 * 1024, hasta); throw new ErrorRemoto(codigo, "Error " + codigo); }
            return leerBytes(c.getInputStream(), MAX_BINARIO, hasta);
        } finally {
            c.disconnect();
        }
    }

    private Object pedir(String metodo, String url, String token, Object cuerpo) throws IOException {
        HttpURLConnection c = abrir(url);
        try {
            c.setRequestMethod(metodo);
            c.setConnectTimeout(8_000);
            c.setReadTimeout(30_000);
            c.setRequestProperty("Accept", "application/json");
            if (token != null) c.setRequestProperty("Authorization", "Bearer " + token);
            if (cuerpo != null) {
                byte[] b = json.escribir(cuerpo);
                c.setDoOutput(true);
                c.setRequestProperty("Content-Type", "application/json;charset=UTF-8");
                c.setFixedLengthStreamingMode(b.length);
                try (OutputStream o = c.getOutputStream()) { o.write(b); }
            }
            int codigo = c.getResponseCode();
            if (codigo >= 300 && codigo < 400) throw new ErrorRemoto(codigo, "El otro servidor manda a otra dirección" + destino(c));
            String t = leer(codigo >= 400 ? c.getErrorStream() : c.getInputStream(), codigo >= 400 ? 64 * 1024 : MAX_RESPUESTA, limite(MAX_MS_PETICION));
            if (codigo >= 400) {
                String m = "Error " + codigo;
                try { Object e = json.leer(t); if (e instanceof Map<?, ?> em && em.get("error") != null) m = String.valueOf(em.get("error")); } catch (RuntimeException ignorado) { }
                throw new ErrorRemoto(codigo, m);
            }
            return t.isBlank() ? null : json.leer(t);
        } finally {
            c.disconnect();
        }
    }

    private static String leer(InputStream in, int max, long hasta) throws IOException {
        return new String(leerBytes(in, max, hasta), StandardCharsets.UTF_8);
    }

    /** Lee la respuesta con tope de tamaño y de tiempo (el otro servidor puede ser lento o mandar de más). */
    private static byte[] leerBytes(InputStream in, int max, long hasta) throws IOException {
        if (in == null) return new byte[0];
        try (InputStream i = in) {
            ByteArrayOutputStream b = new ByteArrayOutputStream();
            byte[] buf = new byte[32 * 1024];
            int n;
            while ((n = i.read(buf)) > 0) {
                if (b.size() + n > max) throw new IOException("La respuesta del otro NOM es demasiado grande");
                if (System.currentTimeMillis() > hasta) throw new SinTiempo();
                b.write(buf, 0, n);
            }
            return b.toByteArray();
        }
    }
}
