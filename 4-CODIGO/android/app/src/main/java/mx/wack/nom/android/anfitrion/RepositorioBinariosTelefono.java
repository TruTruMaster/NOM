package mx.wack.nom.android.anfitrion;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Pattern;

import mx.wack.nom.core.almacen.RepositorioBinarios;

/**
 * El almacén de binarios del teléfono que comparte: un archivo por binario en la carpeta privada de la app
 * ({@code servidor/binarios/<espacio>/<huella>}), igual que en la computadora pero sin cifrado en reposo (la carpeta
 * privada de Android ya no la lee otra app; lo de un espacio cifrado de extremo a extremo llega cifrado igual).
 */
final class RepositorioBinariosTelefono implements RepositorioBinarios {

    private static final Pattern ID = Pattern.compile("^[A-Za-z0-9_-]{1,80}$");
    private static final Pattern HUELLA = Pattern.compile("^[0-9a-f]{64}$");
    private final File raiz;
    private final Map<String, Long> bytesCache = new ConcurrentHashMap<>();

    RepositorioBinariosTelefono(File raiz) {
        this.raiz = raiz;
        if (!raiz.isDirectory() && !raiz.mkdirs()) throw new IllegalStateException("No se pudo crear " + raiz);
    }

    private File carpeta(String espacioId) {
        if (espacioId == null || !ID.matcher(espacioId).matches()) throw new IllegalArgumentException("Espacio no válido");
        return new File(raiz, espacioId);
    }

    private File archivo(String espacioId, String huella) {
        if (huella == null || !HUELLA.matcher(huella).matches()) throw new IllegalArgumentException("Huella no válida");
        return new File(carpeta(espacioId), huella);
    }

    @Override public boolean existe(String espacioId, String huella) { return archivo(espacioId, huella).isFile(); }

    @Override public void guardar(String espacioId, String huella, byte[] datos) {
        File f = archivo(espacioId, huella);
        if (f.isFile()) return;
        File d = f.getParentFile();
        if (!d.isDirectory() && !d.mkdirs()) throw new IllegalStateException("No se pudo crear " + d);
        File tmp = new File(d, huella + "." + Long.toHexString(System.nanoTime()) + ".tmp");
        try (FileOutputStream o = new FileOutputStream(tmp)) { o.write(datos); o.getFD().sync(); }
        catch (IOException e) { tmp.delete(); throw new IllegalStateException("No se pudo guardar el archivo", e); }
        if (!tmp.renameTo(f)) { tmp.delete(); throw new IllegalStateException("No se pudo guardar el archivo"); }
        bytesCache.remove(espacioId);
    }

    @Override public byte[] leer(String espacioId, String huella) {
        File f = archivo(espacioId, huella);
        if (!f.isFile()) return null;
        try { return Files.readAllBytes(f.toPath()); } catch (IOException e) { throw new IllegalStateException("No se pudo leer el archivo", e); }
    }

    @Override public long bytesDe(String espacioId) {
        Long c = bytesCache.get(espacioId);
        if (c != null) return c;
        long n = 0;
        File[] fs = carpeta(espacioId).listFiles();
        if (fs != null) for (File f : fs) if (HUELLA.matcher(f.getName()).matches()) n += f.length();
        bytesCache.put(espacioId, n);
        return n;
    }

    @Override public List<String> huellas(String espacioId) {
        List<String> r = new ArrayList<>();
        File[] fs = carpeta(espacioId).listFiles();
        if (fs != null) for (File f : fs) if (HUELLA.matcher(f.getName()).matches()) r.add(f.getName());
        return r;
    }

    @Override public void borrar(String espacioId, String huella) {
        archivo(espacioId, huella).delete();
        bytesCache.remove(espacioId);
    }

    @Override public long edadMs(String espacioId, String huella) {
        File f = archivo(espacioId, huella);
        return f.isFile() ? System.currentTimeMillis() - f.lastModified() : 0;
    }
}
