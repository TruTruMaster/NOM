package mx.wack.nom.red.nodo;

import java.net.Inet4Address;
import java.net.InetAddress;
import java.net.NetworkInterface;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.UUID;

import org.springframework.core.env.Environment;
import mx.wack.nom.core.servicio.Nodo;
import org.springframework.stereotype.Component;

/**
 * El nodo de esta instancia en la LAN: su id, el nombre del equipo, el puerto
 * HTTP y las direcciones por las que otros pueden entrar (CU-14).
 */
@Component
public class NodoLocal implements Nodo {

    private final String id;
    private final Environment env;

    public NodoLocal(Environment env) {
        this.env = env;
        this.id = idGuardado(env);
    }

    /**
     * El id de este servidor no cambia al reiniciarlo: la app guarda con él los alias de cada espacio (para «volver a
     * ser el mismo») y la sincronía lo usa para no copiarse a sí misma. Vive junto a las llaves («<datos>-llaves»),
     * fuera de los datos cifrados. Si no se puede guardar, uno al azar (como antes).
     */
    static String idGuardado(Environment env) {
        String nuevo = UUID.randomUUID().toString().substring(0, 8);
        try {
            java.nio.file.Path dir = mx.wack.nom.servidor.Configuracion.carpetaLlaves(
                    env.getProperty("nom.datos", "./datos-nom"), env.getProperty("nom.llaves", ""));
            java.nio.file.Path f = dir.resolve("nodo.id");
            if (java.nio.file.Files.isRegularFile(f)) {
                String g = java.nio.file.Files.readString(f, java.nio.charset.StandardCharsets.UTF_8).trim();
                if (g.matches("[0-9a-f]{8}")) return g;
            }
            java.nio.file.Files.createDirectories(dir);
            java.nio.file.Files.writeString(f, nuevo + "\n", java.nio.charset.StandardCharsets.UTF_8);
        } catch (Exception e) { /* sin carpeta de llaves: al azar */ }
        return nuevo;
    }

    public String id() { return id; }

    public int puerto() {
        String p = env.getProperty("local.server.port", env.getProperty("server.port", "8080"));
        return Integer.parseInt(p);
    }

    public String equipo() {
        try {
            return InetAddress.getLocalHost().getHostName();
        } catch (Exception e) {
            return "equipo-" + id;
        }
    }

    /** Direcciones IPv4 privadas de interfaces activas (no loopback, no virtuales). */
    public List<String> direcciones() {
        List<String> r = new ArrayList<>();
        try {
            for (NetworkInterface ni : Collections.list(NetworkInterface.getNetworkInterfaces())) {
                if (!ni.isUp() || ni.isLoopback() || ni.isVirtual()) continue;
                String n = ni.getName().toLowerCase();
                if (n.startsWith("docker") || n.startsWith("veth") || n.startsWith("br-") || n.startsWith("virbr")) continue;
                for (InetAddress a : Collections.list(ni.getInetAddresses())) {
                    if (a instanceof Inet4Address && a.isSiteLocalAddress()) r.add(a.getHostAddress());
                }
            }
        } catch (Exception ignorado) {
            // Sin interfaces legibles: se trata como "sin red".
        }
        return r;
    }

    public List<InetAddress> difusion() {
        List<InetAddress> r = new ArrayList<>();
        try {
            for (NetworkInterface ni : Collections.list(NetworkInterface.getNetworkInterfaces())) {
                if (!ni.isUp() || ni.isLoopback()) continue;
                ni.getInterfaceAddresses().forEach(ia -> {
                    if (ia.getBroadcast() != null) r.add(ia.getBroadcast());
                });
            }
        } catch (Exception ignorado) { }
        return r;
    }
}
