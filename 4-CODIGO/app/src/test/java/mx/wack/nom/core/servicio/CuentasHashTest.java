package mx.wack.nom.core.servicio;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.Base64;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import javax.crypto.SecretKeyFactory;
import javax.crypto.spec.PBEKeySpec;

import org.junit.jupiter.api.Test;

import mx.wack.nom.core.almacen.RepositorioUsuarios;
import mx.wack.nom.core.modelo.Sesion;
import mx.wack.nom.core.modelo.Usuario;

/** Una cuenta con el hash de antes entra igual y queda con el formato nuevo. */
class CuentasHashTest {

    static final class Memoria implements RepositorioUsuarios {
        final Map<String, Usuario> porId = new HashMap<>();
        final Map<String, Sesion> sesiones = new HashMap<>();
        public Optional<Usuario> porId(String id) { return Optional.ofNullable(porId.get(id)); }
        public Optional<Usuario> porNombre(String n) { return porId.values().stream().filter(u -> u.nombre.equalsIgnoreCase(n)).findFirst(); }
        public List<Usuario> todos() { return new ArrayList<>(porId.values()); }
        public void guardar(Usuario u) { porId.put(u.id, u); }
        public void borrar(String id) { porId.remove(id); }
        public Optional<Sesion> sesion(String t) { return Optional.ofNullable(sesiones.get(t)); }
        public void guardarSesion(Sesion s) { sesiones.put(s.token, s); }
        public void borrarSesion(String t) { sesiones.remove(t); }
        public void borrarSesionesDe(String id) { sesiones.values().removeIf(s -> id.equals(s.usuarioId)); }
    }

    @Test
    void laCuentaViejaEntraYSeActualiza() throws Exception {
        Memoria repo = new Memoria();
        Cifrado cifrado = new Cifrado(130_000);
        ServicioCuentas cuentas = new ServicioCuentas(repo, cifrado, 5, 15, 24);
        Usuario u = cuentas.registrar("maestra", "contraseña segura", "WACK");
        // Como lo dejaba la versión anterior: base64 de 120 000 iteraciones, sin prefijo.
        PBEKeySpec spec = new PBEKeySpec("contraseña segura".toCharArray(), Base64.getDecoder().decode(u.sal), 120_000, 256);
        u.hashContrasena = Base64.getEncoder().encodeToString(SecretKeyFactory.getInstance("PBKDF2WithHmacSHA256").generateSecret(spec).getEncoded());
        repo.guardar(u);
        String antes = u.hashContrasena;

        assertThrows(ErrorNom.class, () -> cuentas.iniciarSesion("maestra", "otra contraseña"));
        assertTrue(cifrado.necesitaActualizar(repo.porNombre("maestra").get().hashContrasena), "un fallo no la cambia");

        assertNotNull(cuentas.iniciarSesion("maestra", "contraseña segura"));
        Usuario hoy = repo.porNombre("maestra").get();
        assertNotEquals(antes, hoy.hashContrasena);
        assertTrue(hoy.hashContrasena.startsWith("pbkdf2$130000$"), hoy.hashContrasena);
        assertFalse(cifrado.necesitaActualizar(hoy.hashContrasena));
        assertNotNull(cuentas.iniciarSesion("maestra", "contraseña segura"), "entra con el formato nuevo");
        assertThrows(ErrorNom.class, () -> cuentas.iniciarSesion("maestra", "contraseña insegura"));
    }
}
