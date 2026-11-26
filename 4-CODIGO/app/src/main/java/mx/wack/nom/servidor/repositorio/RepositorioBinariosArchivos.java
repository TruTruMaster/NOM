package mx.wack.nom.servidor.repositorio;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Pattern;
import java.util.stream.Stream;

import org.springframework.stereotype.Component;

import mx.wack.nom.core.almacen.RepositorioBinarios;
import mx.wack.nom.core.servicio.CifradoReposo;

/**
 * El almacén de binarios en disco: un archivo por binario en {@code <datos>/binarios/<espacio>/<huella>}, cifrado en
 * reposo como todo lo demás (AES-256-GCM con la ruta autenticada: un archivo no se puede cambiar por otro), sin
 * comprimir (imágenes y archivos de Office ya vienen comprimidos). Se escribe a un temporal y se mueve de una vez:
 * si se va la luz a la mitad no queda un binario a medias con el nombre de su huella.
 */
@Component
public class RepositorioBinariosArchivos implements RepositorioBinarios {

    private static final Pattern ID = Pattern.compile("^[A-Za-z0-9_-]{1,80}$");
    private static final Pattern HUELLA = Pattern.compile("^[0-9a-f]{64}$");
    private final CifradoReposo cifrado = mx.wack.nom.servidor.ArranqueCifrado.cifrado();
    private final Path raiz;
    /** Lo que ocupa cada espacio (se recalcula al cambiar algo). */
    private final Map<String, Long> bytesCache = new ConcurrentHashMap<>();

    public RepositorioBinariosArchivos(AlmacenJson almacen) throws IOException {
        this.raiz = almacen.raiz().resolve("binarios");
        Files.createDirectories(raiz);
    }

    private Path carpeta(String espacioId) {
        if (espacioId == null || !ID.matcher(espacioId).matches()) throw new IllegalArgumentException("Espacio no válido");
        return raiz.resolve(espacioId);
    }

    private Path archivo(String espacioId, String huella) {
        if (huella == null || !HUELLA.matcher(huella).matches()) throw new IllegalArgumentException("Huella no válida");
        return carpeta(espacioId).resolve(huella);
    }

    /** La ruta que va autenticada en el cifrado (la misma que usa la migración de AlmacenJson). */
    private static String relativa(String espacioId, String huella) { return "binarios/" + espacioId + "/" + huella; }

    @Override
    public boolean existe(String espacioId, String huella) { return Files.isRegularFile(archivo(espacioId, huella)); }

    @Override
    public void guardar(String espacioId, String huella, byte[] datos) {
        Path f = archivo(espacioId, huella);
        if (Files.isRegularFile(f)) return;
        try {
            Files.createDirectories(f.getParent());
            byte[] sale = cifrado.activo() ? cifrado.cifrar(datos, false, datos.length, relativa(espacioId, huella)) : datos;
            // (termina en .tmp: la migración a cifrado de AlmacenJson no lo toca mientras se escribe)
            Path tmp = f.resolveSibling(huella + "." + Long.toHexString(System.nanoTime()) + ".tmp");
            try {
                Files.write(tmp, sale);
                try { Files.move(tmp, f, StandardCopyOption.ATOMIC_MOVE, StandardCopyOption.REPLACE_EXISTING); }
                catch (java.nio.file.AtomicMoveNotSupportedException ex) { Files.move(tmp, f, StandardCopyOption.REPLACE_EXISTING); }
            } catch (IOException e) {
                try { Files.deleteIfExists(tmp); } catch (IOException ignorado) { }
                throw e;
            }
            bytesCache.remove(espacioId);
        } catch (IOException e) { throw new UncheckedIOException("No se pudo guardar el archivo", e); }
    }

    @Override
    public byte[] leer(String espacioId, String huella) {
        Path f = archivo(espacioId, huella);
        if (!Files.isRegularFile(f)) return null;
        try {
            byte[] b = Files.readAllBytes(f);
            if (!CifradoReposo.cifrado(b)) return b;
            byte[] d = cifrado.descifrar(b, relativa(espacioId, huella));
            // (si lo cifró la migración de datos viejos, puede venir comprimido)
            if (CifradoReposo.conGzip(b)) try (var in = new java.util.zip.GZIPInputStream(new ByteArrayInputStream(d))) { d = in.readAllBytes(); }
            return d;
        } catch (IOException e) { throw new UncheckedIOException("No se pudo leer el archivo", e); }
    }

    @Override
    public long bytesDe(String espacioId) {
        Long c = bytesCache.get(espacioId);
        if (c != null) return c;
        long n = 0;
        Path d = carpeta(espacioId);
        if (Files.isDirectory(d)) {
            try (Stream<Path> s = Files.list(d)) {
                for (Path p : (Iterable<Path>) s::iterator) if (HUELLA.matcher(p.getFileName().toString()).matches()) n += Files.size(p);
            } catch (IOException e) { /* se cuenta lo que se pudo */ }
        }
        bytesCache.put(espacioId, n);
        return n;
    }

    @Override
    public List<String> huellas(String espacioId) {
        List<String> r = new ArrayList<>();
        Path d = carpeta(espacioId);
        if (!Files.isDirectory(d)) return r;
        try (Stream<Path> s = Files.list(d)) {
            s.map(p -> p.getFileName().toString()).filter(n -> HUELLA.matcher(n).matches()).forEach(r::add);
        } catch (IOException e) { /* nada */ }
        return r;
    }

    @Override
    public void borrar(String espacioId, String huella) {
        try { Files.deleteIfExists(archivo(espacioId, huella)); } catch (IOException e) { /* se intenta la siguiente vez */ }
        bytesCache.remove(espacioId);
    }

    @Override
    public void tocar(String espacioId, String huella) {
        try { Files.setLastModifiedTime(archivo(espacioId, huella), java.nio.file.attribute.FileTime.fromMillis(System.currentTimeMillis())); }
        catch (IOException e) { /* nada: solo da más tiempo antes de la limpieza */ }
    }

    @Override
    public long edadMs(String espacioId, String huella) {
        try { return System.currentTimeMillis() - Files.getLastModifiedTime(archivo(espacioId, huella)).toMillis(); }
        catch (IOException e) { return 0; }
    }
}
