package mx.wack.nom.core.servicio;

import java.security.MessageDigest;
import java.security.SecureRandom;
import java.security.spec.KeySpec;
import java.util.Arrays;
import java.util.Base64;

import javax.crypto.SecretKeyFactory;
import javax.crypto.spec.PBEKeySpec;


/**
 * CU-05 Cifrar credenciales. PBKDF2-HMAC-SHA256 con sal aleatoria de 16 bytes.
 * Viene en el JDK, asi que no agrega dependencias.
 *
 * Formato guardado: {@code pbkdf2$<iteraciones>$<base64>} (600 000 iteraciones: ~0.2 s en la Raspberry Pi 5,
 * ~0.8 s en una laptop sin instrucciones SHA). Lo guardado antes de este formato es solo el base64 de 120 000
 * iteraciones: se sigue aceptando, tarda lo mismo que el nuevo (se completa el trabajo que falta) y la cuenta se
 * pasa al formato nuevo la siguiente vez que entra ({@link #necesitaActualizar}).
 */
public class Cifrado {

    static final int ITERACIONES = 600_000;
    static final int ITERACIONES_ANTES = 120_000;
    private static final String PREFIJO = "pbkdf2$";
    private static final int BITS = 256;
    private final SecureRandom azar = new SecureRandom();
    private final int iteraciones;

    public Cifrado() { this(ITERACIONES); }

    /** Para pruebas: menos iteraciones (nunca menos que las de antes). */
    public Cifrado(int iteraciones) { this.iteraciones = Math.max(ITERACIONES_ANTES, iteraciones); }

    public String nuevaSal() {
        byte[] s = new byte[16];
        azar.nextBytes(s);
        return Base64.getEncoder().encodeToString(s);
    }

    /** El hash para guardar, en el formato con versión. Borra la contraseña del arreglo al terminar. */
    public String hash(char[] contrasena, String sal) {
        try {
            return PREFIJO + iteraciones + "$" + Base64.getEncoder().encodeToString(pbkdf2(contrasena, sal, iteraciones));
        } finally {
            Arrays.fill(contrasena, '\0');   // paso 4: se descarta el texto claro
        }
    }

    private static byte[] pbkdf2(char[] contrasena, String sal, int vueltas) {
        try {
            KeySpec spec = new PBEKeySpec(contrasena, Base64.getDecoder().decode(sal), vueltas, BITS);
            return SecretKeyFactory.getInstance("PBKDF2WithHmacSHA256").generateSecret(spec).getEncoded();
        } catch (Exception e) {
            throw new IllegalStateException("PBKDF2 no disponible", e);
        }
    }

    /** Iteraciones de un hash guardado (0 si no se entiende). */
    static int iteracionesDe(String guardado) {
        if (guardado == null) return 0;
        if (!guardado.startsWith(PREFIJO)) return ITERACIONES_ANTES;
        int fin = guardado.indexOf('$', PREFIJO.length());
        if (fin < 0 || fin - PREFIJO.length() > 8) return 0;
        try {
            int n = Integer.parseInt(guardado.substring(PREFIJO.length(), fin));
            return n >= ITERACIONES_ANTES && n <= 10_000_000 ? n : 0;
        } catch (NumberFormatException e) { return 0; }
    }

    /** Comparacion en tiempo constante (y el mismo tiempo con un hash de antes que con uno nuevo). */
    public boolean coincide(char[] contrasena, String sal, String hashGuardado) {
        try {
            int n = iteracionesDe(hashGuardado);
            if (n == 0 || sal == null) {
                pbkdf2(contrasena, nuevaSal(), iteraciones);   // mismo trabajo, nunca coincide
                return false;
            }
            byte[] calculado = pbkdf2(contrasena, sal, n);
            String guardado = hashGuardado.startsWith(PREFIJO) ? hashGuardado.substring(hashGuardado.indexOf('$', PREFIJO.length()) + 1) : hashGuardado;
            if (n < iteraciones) pbkdf2("relleno".toCharArray(), sal, iteraciones - n);   // tarda lo mismo que uno nuevo
            return MessageDigest.isEqual(Base64.getEncoder().encodeToString(calculado).getBytes(), guardado.getBytes());
        } finally {
            Arrays.fill(contrasena, '\0');
        }
    }

    /** Si un hash guardado es de antes (o con menos iteraciones) y conviene rehacerlo al entrar. */
    public boolean necesitaActualizar(String hashGuardado) {
        int n = iteracionesDe(hashGuardado);
        return n != 0 && (n < iteraciones || !hashGuardado.startsWith(PREFIJO));
    }

    public String token() {
        byte[] t = new byte[32];
        azar.nextBytes(t);
        return Base64.getUrlEncoder().withoutPadding().encodeToString(t);
    }

    public String id() {
        byte[] t = new byte[9];
        azar.nextBytes(t);
        return Base64.getUrlEncoder().withoutPadding().encodeToString(t);
    }
}
