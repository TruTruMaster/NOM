package mx.wack.nom.core.servicio;

import java.security.MessageDigest;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import mx.wack.nom.core.almacen.RepositorioBinarios;
import mx.wack.nom.core.modelo.Documento;
import mx.wack.nom.core.modelo.Espacio;
import mx.wack.nom.core.modelo.Permiso;
import mx.wack.nom.core.modelo.Usuario;

/**
 * Almacén de binarios por espacio (imágenes, el Office original de un documento importado; después video y audio).
 *
 * Por qué: dentro del documento las imágenes iban como texto base64 (un tercio más pesadas) y repetidas (una vez en
 * el original de Office y otra en cada bloque que las usa): un .pptx de 10 MB pesaba 30 MB dentro de NOM y no cabía
 * en el tope por documento. Ahora el documento guarda solo una referencia corta, «nom:bin/HUELLA.ext», y el binario
 * va aparte, una sola vez por espacio.
 *
 * - La huella es el SHA-256 (en hexadecimal) de los bytes que llegan; se comprueba al subir, así que un binario no
 *   se puede cambiar por otro. En un espacio cifrado de extremo a extremo los bytes llegan YA cifrados en el equipo
 *   con la llave del documento: el servidor no ve el contenido ni su huella real.
 * - Permisos: subir, quien edita en el espacio; bajar, quien lo puede ver. Sin sesión, solo con el enlace para ver
 *   o el código del proyector y solo lo que ese documento usa (ver bajarPublico).
 * - Cuenta en la cuota del propietario del espacio, como los documentos y el chat.
 */
public class ServicioBinarios {

    public static final Pattern HUELLA = Pattern.compile("^[0-9a-f]{64}$");
    /** Lo que se borra si nadie lo usa: no antes de esto desde que se subió (un documento puede tardar en llegar). */
    static final long GRACIA_MS = 30L * 24 * 3600 * 1000;

    private final ServicioEspacios espacios;
    private final ServicioDocumentos documentos;
    private final RepositorioBinarios repo;
    private volatile long maxBytes;

    public ServicioBinarios(ServicioEspacios espacios, ServicioDocumentos documentos, RepositorioBinarios repo, long maxBytes) {
        this.espacios = espacios;
        this.documentos = documentos;
        this.repo = repo;
        this.maxBytes = maxBytes > 0 ? maxBytes : 48L * 1024 * 1024;
        // lo que ocupa el almacén de cada espacio cuenta en la cuota de su propietario
        if (documentos != null) {
            documentos.contarBinarios(repo::bytesDe);
            // al copiar o mover un documento a otro espacio, sus binarios van con él
            documentos.copiarBinariosCon(this::copiarEntreEspacios);
        }
    }

    /**
     * Copia al espacio destino los binarios (huellas) que usa un documento que se copia o se mueve allá. Lo que el origen
     * no tenga (todavía no terminaba de subirse) se omite: el equipo que lo tenga lo subirá. Devuelve cuántos copió.
     */
    int copiarEntreEspacios(String deEspacio, String aEspacio, java.util.Set<String> huellas) {
        if (huellas == null || huellas.isEmpty() || deEspacio.equals(aEspacio)) return 0;
        int n = 0;
        for (String h : huellas) if (HUELLA.matcher(h).matches() && repo.copiar(deEspacio, aEspacio, h)) n++;
        return n;
    }

    public long maxBytes() { return maxBytes; }

    static String huellaDe(byte[] datos) {
        try { return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(datos)); }
        catch (java.security.NoSuchAlgorithmException e) { throw new IllegalStateException(e); }
    }

    private static String validar(String huella) {
        String h = huella == null ? "" : huella.trim().toLowerCase(java.util.Locale.ROOT);
        // (la referencia puede traer la extensión: «huella.png»)
        int p = h.indexOf('.');
        if (p > 0) h = h.substring(0, p);
        if (!HUELLA.matcher(h).matches()) throw ErrorNom.invalido("Esa no es la huella de un archivo", "huella");
        return h;
    }

    /** Subir un binario. Si ya estaba, no se vuelve a guardar. */
    public Map<String, Object> subir(Usuario u, String espacioId, String huella, byte[] datos) {
        Espacio e = espacios.requerir(espacioId);
        espacios.exigir(e, u, Permiso.EDICION);
        String h = validar(huella);
        if (datos == null || datos.length == 0) throw ErrorNom.invalido("El archivo llegó vacío", "datos");
        if (datos.length > maxBytes)
            throw new ErrorNom(ErrorNom.LLENO, "El archivo pasa de " + (maxBytes / (1024 * 1024)) + " MB", null, Map.of("tope", maxBytes));
        if (!huellaDe(datos).equals(h)) throw ErrorNom.invalido("El archivo llegó distinto de como salió (su huella no coincide): vuelve a intentarlo", "huella");
        Map<String, Object> r = new LinkedHashMap<>();
        r.put("huella", h);
        r.put("bytes", datos.length);
        // (ya estaba: se marca como recién subido, para que la limpieza no lo borre antes de que llegue quien lo usa)
        if (repo.existe(e.id, h)) { repo.tocar(e.id, h); r.put("nuevo", false); return r; }
        documentos.revisarCuotaDe(e);
        repo.guardar(e.id, h, datos);
        documentos.olvidarUso(e.propietarioId);
        r.put("nuevo", true);
        return r;
    }

    /** Bajar un binario del espacio (quien lo puede ver). */
    public byte[] bajar(Usuario u, String espacioId, String huella) {
        Espacio e = espacios.requerir(espacioId);
        espacios.exigir(e, u, Permiso.LECTURA);
        byte[] b = repo.leer(e.id, validar(huella));
        if (b == null) throw new ErrorNom(ErrorNom.NO_EXISTE, "Ese archivo no está en el espacio (todavía no termina de subirse o se borró)");
        return b;
    }

    /** De una lista de huellas, las que el espacio todavía no tiene (para subir solo lo que falta). */
    public List<String> faltan(Usuario u, String espacioId, List<String> huellas) {
        Espacio e = espacios.requerir(espacioId);
        espacios.exigir(e, u, Permiso.EDICION);
        List<String> r = new ArrayList<>();
        if (huellas == null) return r;
        if (huellas.size() > 5000) throw ErrorNom.invalido("Demasiadas huellas de una vez", "huellas");
        for (String x : huellas) {
            String h = validar(x);
            if (!repo.existe(e.id, h) && !r.contains(h)) r.add(h);
        }
        return r;
    }

    /**
     * Sin sesión, con el enlace para ver (tipo "ver") o el código del proyector (tipo "proyeccion"): solo un binario
     * que ese documento usa. En un documento normal se comprueba en lo que se le muestra a quien mira (sin el original
     * de Office, comentarios ni notas); en uno cifrado de extremo a extremo el servidor no puede leerlo, pero el
     * binario va cifrado con la llave del documento y su huella no se puede adivinar.
     */
    public byte[] bajarPublico(String tipo, String codigo, String huella) {
        String h = validar(huella);
        Documento d = "proyeccion".equals(tipo) ? documentos.docDeProyeccion(codigo) : documentos.docDeEnlaceVer(codigo);
        Espacio e = espacios.requerir(d.espacioId);
        if (!e.e2e && !documentos.muestraBinario(d, "proyeccion".equals(tipo), h))
            throw new ErrorNom(ErrorNom.NO_EXISTE, "Ese archivo no es parte de este documento");
        byte[] b = repo.leer(e.id, h);
        if (b == null) throw new ErrorNom(ErrorNom.NO_EXISTE, "Ese archivo no está en el espacio");
        return b;
    }

    /** Lo que ocupa el almacén de un espacio. */
    public long bytesDe(String espacioId) { return repo.bytesDe(espacioId); }

    /**
     * Borra lo que ningún documento del espacio (ni sus versiones ni la papelera) usa desde hace más de 30 días.
     * Devuelve cuántos se borraron. (En un espacio cifrado de extremo a extremo las referencias van cifradas: ahí no se
     * borra nada solo.)
     */
    /** La limpieza programada: todos los espacios (salvo los cifrados de extremo a extremo). Devuelve cuántos borró. */
    public int limpiarTodo() {
        int n = 0;
        for (Espacio e : espacios.todos()) {
            try { n += limpiar(e.id); } catch (RuntimeException ex) { /* un espacio con problemas no detiene a los demás */ }
        }
        return n;
    }

    public int limpiar(String espacioId) {
        Espacio e = espacios.requerir(espacioId);
        if (e.e2e) return 0;
        java.util.Set<String> usados = documentos.binariosUsados(espacioId);
        int n = 0;
        for (String h : repo.huellas(espacioId)) {
            if (usados.contains(h) || repo.edadMs(espacioId, h) < GRACIA_MS) continue;
            repo.borrar(espacioId, h);
            n++;
        }
        if (n > 0) documentos.olvidarUso(e.propietarioId);
        return n;
    }
}
