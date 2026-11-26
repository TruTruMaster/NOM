package mx.wack.nom.core.servicio;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import mx.wack.nom.core.almacen.RepositorioEspacios;
import mx.wack.nom.core.modelo.Espacio;
import mx.wack.nom.core.modelo.Permiso;
import mx.wack.nom.core.modelo.Usuario;

/**
 * Fase 15, parte E: el chat de cada espacio, con imagenes y documentos.
 *
 * El espacio es cerrado (solo por invitacion), asi que el chat tambien: no hay chat abierto ni
 * busqueda de personas. En un espacio cifrado de extremo a extremo, el texto y los adjuntos llegan
 * cifrados desde el equipo de quien escribe y el servidor solo los guarda y los reparte; en los
 * demas espacios van cifrados en reposo en el disco, como todo lo demas.
 *
 * Se guardan los ultimos 2000 mensajes; cada adjunto (hasta 2 MB) va aparte y cuenta en el espacio
 * del propietario. Borrar un mensaje lo puede su autor o el propietario del espacio.
 */
public class ServicioChat {

    public static final int MAX_MENSAJES = 2000;
    public static final int MAX_TEXTO = 8000;
    public static final int MAX_ADJUNTO = 2 * 1024 * 1024 * 4 / 3 + 1024;   // 2 MB en base64
    /** Lo más que ocupan los textos del chat de un espacio (se reescribe completo con cada mensaje): se van los más viejos. */
    static final long MAX_BYTES_CHAT = 4L * 1024 * 1024;
    private final Topes topes = new Topes();

    private final RepositorioEspacios repo;
    private final ServicioEspacios espacios;
    private final Difusor difusor;
    private final Cifrado cifrado;

    /** DEF-42: con cuota (en la nube), los adjuntos cuentan en el espacio del propietario y se revisa el limite. */
    private volatile ServicioDocumentos cuota;

    public void usarCuota(ServicioDocumentos documentos) {
        this.cuota = documentos;
        if (documentos != null) documentos.contarChat(this::bytes);
    }

    public ServicioChat(RepositorioEspacios repo, ServicioEspacios espacios, Difusor difusor, Cifrado cifrado) {
        this.repo = repo;
        this.espacios = espacios;
        this.difusor = difusor;
        this.cifrado = cifrado;
    }

    /** Los mensajes despues de 'desde' (epoch ms), sin el contenido de los adjuntos. */
    public Map<String, Object> mensajes(Usuario u, String espacioId, long desde) {
        Espacio e = espacios.requerir(espacioId);
        espacios.exigir(e, u, Permiso.LECTURA);
        List<Map<String, Object>> r = new ArrayList<>();
        for (Map<String, Object> m : repo.chat(espacioId)) if (num(m.get("t")) > desde) r.add(m);
        Map<String, Object> x = new LinkedHashMap<>();
        x.put("e2e", e.e2e);
        x.put("mensajes", r.size() > 500 ? r.subList(r.size() - 500, r.size()) : r);
        return x;
    }

    private static long num(Object o) { return o instanceof Number n ? n.longValue() : 0; }

    /** Un mensaje nuevo: texto (o cifrado) y, si trae, un adjunto { tipo, nombre, mime, datos (base64) }. */
    @SuppressWarnings("unchecked")
    public synchronized Map<String, Object> enviar(Usuario u, String espacioId, Map<String, Object> c) {
        Espacio e = espacios.requerir(espacioId);
        espacios.exigir(e, u, Permiso.LECTURA);
        topes.exigir("u|" + u.id, 30, "Muchos mensajes seguidos: espera un momento");
        topes.exigir("e|" + espacioId, 120, "Hay muchos mensajes a la vez en este espacio: espera un momento");
        Object texto = c.get("texto");
        Object adj = c.get("adjunto");
        String t = texto == null ? "" : String.valueOf(texto);
        if (t.length() > MAX_TEXTO * 2) throw ErrorNom.invalido("El mensaje es demasiado largo", "texto");
        if (e.e2e && !t.isEmpty() && !ServicioEspacios.cifrado(t)) throw ErrorNom.invalido("En un espacio cifrado los mensajes se cifran en tu equipo (actualiza la app)", "texto");
        if (t.isBlank() && !(adj instanceof Map)) throw ErrorNom.invalido("Escribe un mensaje", "texto");
        Map<String, Object> m = new LinkedHashMap<>();
        m.put("id", cifrado.id());
        m.put("t", System.currentTimeMillis());
        m.put("usuarioId", u.id);
        m.put("usuario", u.nombre);
        m.put("color", u.color);
        m.put("texto", t);
        if (adj instanceof Map<?, ?> a) {
            Object datos = a.get("datos");
            if (!(datos instanceof String d) || d.isEmpty()) throw ErrorNom.invalido("El archivo llegó vacío", "adjunto");
            if (e.e2e && !ServicioEspacios.cifrado(d)) throw ErrorNom.invalido("En un espacio cifrado los archivos se cifran en tu equipo (actualiza la app)", "adjunto");
            if (!(e.e2e ? base64(d, 4, true) : base64(d, 0, false))) throw ErrorNom.invalido("El archivo llegó dañado: vuelve a adjuntarlo", "adjunto");
            // Cifrado, el archivo crece un tercio más (base64 del base64 cifrado).
            if (d.length() > (e.e2e ? MAX_ADJUNTO / 3 * 4 + 256 : MAX_ADJUNTO)) throw new ErrorNom(ErrorNom.LLENO, "El archivo pasa de 2 MB: compártelo como documento del espacio", null, Map.of("tope", 2 * 1024 * 1024));
            ServicioDocumentos q = cuota;
            if (q != null) q.revisarCuotaDe(e);   // DEF-42: sin lugar en la cuenta del propietario no se suben archivos
            String tipo = "imagen".equals(a.get("tipo")) ? "imagen" : "archivo";
            Map<String, Object> info = new LinkedHashMap<>();
            info.put("tipo", tipo);
            info.put("nombre", recorta(a.get("nombre"), 400));
            info.put("mime", recorta(a.get("mime"), 100));
            info.put("bytes", d.length() * 3L / 4);
            m.put("adjunto", info);
            repo.guardarAnexoEspacio(espacioId, "chat-" + m.get("id"), Map.of("datos", d));
        }
        if (c.get("docId") instanceof String doc && doc.matches("[A-Za-z0-9_-]{4,40}")) m.put("docId", doc);
        List<Map<String, Object>> l = new ArrayList<>(repo.chat(espacioId));
        l.add(m);
        long total = 0;
        for (Map<String, Object> x : l) total += String.valueOf(x.get("texto")).length() + 200;
        while (l.size() > MAX_MENSAJES || (total > MAX_BYTES_CHAT && l.size() > 1)) {
            Map<String, Object> viejo = l.remove(0);
            total -= String.valueOf(viejo.get("texto")).length() + 200;
            if (viejo.get("adjunto") != null) repo.guardarAnexoEspacio(espacioId, "chat-" + viejo.get("id"), null);
        }
        repo.guardarChat(espacioId, l);
        ServicioDocumentos q = cuota;
        if (q != null) q.olvidarUso(e.propietarioId);
        difusor.avisarTodoEspacio(espacioId, "chat", m);
        return m;
    }

    /** ¿Es texto en base64 (desde `desde`)? Normal: A-Z, a-z, 0-9, + y / con hasta dos = al final; para URL: - y _ sin =. */
    static boolean base64(String d, int desde, boolean url) {
        int n = d.length(), fin = n;
        if (!url) while (fin > desde && n - fin < 2 && d.charAt(fin - 1) == '=') fin--;
        for (int i = desde; i < fin; i++) {
            char c = d.charAt(i);
            boolean ok = (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9')
                    || (url ? c == '-' || c == '_' : c == '+' || c == '/');
            if (!ok) return false;
        }
        return true;
    }

    private static String recorta(Object o, int max) {
        String s = o == null ? "" : String.valueOf(o);
        return Textos.cortar(Textos.limpio(s), max);
    }

    /** El contenido de un adjunto (base64, cifrado si el espacio lo esta). */
    @SuppressWarnings("unchecked")
    public Map<String, Object> adjunto(Usuario u, String espacioId, String mensajeId) {
        Espacio e = espacios.requerir(espacioId);
        espacios.exigir(e, u, Permiso.LECTURA);
        if (mensajeId == null || !mensajeId.matches("[A-Za-z0-9_-]{4,40}")) throw ErrorNom.noEncontrado("El archivo");
        Object a = repo.anexoEspacio(espacioId, "chat-" + mensajeId);
        if (!(a instanceof Map<?, ?> m) || m.get("datos") == null) throw ErrorNom.noEncontrado("El archivo");
        return Map.of("datos", m.get("datos"));
    }

    /** Borra un mensaje (su autor o el propietario del espacio). */
    public synchronized Map<String, Object> borrar(Usuario u, String espacioId, String mensajeId) {
        Espacio e = espacios.requerir(espacioId);
        Permiso p = espacios.exigir(e, u, Permiso.LECTURA);
        List<Map<String, Object>> l = new ArrayList<>(repo.chat(espacioId));
        Map<String, Object> m = l.stream().filter(x -> String.valueOf(x.get("id")).equals(mensajeId)).findFirst().orElseThrow(() -> ErrorNom.noEncontrado("El mensaje"));
        if (!u.id.equals(m.get("usuarioId")) && p != Permiso.PROPIETARIO) throw ErrorNom.prohibido("Solo quien lo escribió o el propietario del espacio lo borra");
        l.remove(m);
        if (m.get("adjunto") != null) repo.guardarAnexoEspacio(espacioId, "chat-" + mensajeId, null);
        repo.guardarChat(espacioId, l);
        difusor.avisarTodoEspacio(espacioId, "chat", Map.of("borrado", mensajeId));
        return Map.of("ok", true);
    }

    /** Lo que ocupa el chat de un espacio (cuenta en la cuota del propietario). */
    public long bytes(String espacioId) {
        long n = 0;
        for (Map<String, Object> m : repo.chat(espacioId)) {
            n += String.valueOf(m.get("texto")).length() + 200;
            if (m.get("adjunto") instanceof Map<?, ?> a) n += num(a.get("bytes"));
        }
        return n;
    }
}
