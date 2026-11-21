package mx.wack.nom.core.servicio;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Base64;

import javax.crypto.SecretKeyFactory;
import javax.crypto.spec.PBEKeySpec;

import org.junit.jupiter.api.Test;

/** El hash de contraseñas con versión: el nuevo, el de antes y lo que no se entiende. */
class CifradoTest {

    /** Lo que guardaba la versión anterior: solo el base64 de 120 000 iteraciones. */
    private static String hashDeAntes(String contrasena, String sal) throws Exception {
        PBEKeySpec spec = new PBEKeySpec(contrasena.toCharArray(), Base64.getDecoder().decode(sal), 120_000, 256);
        return Base64.getEncoder().encodeToString(SecretKeyFactory.getInstance("PBKDF2WithHmacSHA256").generateSecret(spec).getEncoded());
    }

    @Test
    void elNuevoLlevaSuVersionYCabeEnLaColumna() {
        Cifrado c = new Cifrado();
        String sal = c.nuevaSal();
        String h = c.hash("una contraseña larga".toCharArray(), sal);
        assertTrue(h.startsWith("pbkdf2$600000$"), h);
        assertTrue(h.length() <= 64, "la columna hash_contrasena es VARCHAR(64)");
        assertTrue(c.coincide("una contraseña larga".toCharArray(), sal, h));
        assertFalse(c.coincide("otra".toCharArray(), sal, h));
        assertFalse(c.necesitaActualizar(h));
    }

    @Test
    void elDeAntesSigueEntrandoYPideActualizarse() throws Exception {
        Cifrado c = new Cifrado(130_000);
        String sal = c.nuevaSal();
        String viejo = hashDeAntes("contraseña de 2025", sal);
        assertTrue(c.coincide("contraseña de 2025".toCharArray(), sal, viejo));
        assertFalse(c.coincide("contraseña de 2024".toCharArray(), sal, viejo));
        assertTrue(c.necesitaActualizar(viejo));
        assertEquals(120_000, Cifrado.iteracionesDe(viejo));
    }

    @Test
    void loQueNoSeEntiendeNuncaCoincide() {
        Cifrado c = new Cifrado(130_000);
        String sal = c.nuevaSal();
        for (String raro : new String[] { "pbkdf2$", "pbkdf2$abc$xyz", "pbkdf2$5$xyz", "pbkdf2$999999999$xyz" }) {
            assertEquals(0, Cifrado.iteracionesDe(raro), raro);
            assertFalse(c.coincide("x".toCharArray(), sal, raro), raro);
            assertFalse(c.necesitaActualizar(raro), raro);
        }
        assertFalse(c.coincide("x".toCharArray(), null, "pbkdf2$130000$xyz"));
    }

    @Test
    void borraLaContrasenaDelArreglo() {
        Cifrado c = new Cifrado(130_000);
        char[] a = "secreto".toCharArray();
        String sal = c.nuevaSal();
        String h = c.hash(a, sal);
        assertEquals("\0\0\0\0\0\0\0", new String(a));
        char[] b = "secreto".toCharArray();
        assertTrue(c.coincide(b, sal, h));
        assertEquals("\0\0\0\0\0\0\0", new String(b));
    }
}
