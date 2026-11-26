package mx.wack.nom.android.anfitrion;

import android.content.Context;
import android.os.Build;
import android.util.Base64;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.SecureRandom;

/**
 * Llave local con la que la app entra a su propio servidor en el teléfono.
 * Vive en la carpeta privada sin respaldo para que no salga en respaldos ni sea conocida fuera de la app.
 */
public final class LlaveDueno {

    private LlaveDueno() { }

    public static synchronized String obtener(Context ctx) {
        File dir = Build.VERSION.SDK_INT >= 21 ? ctx.getNoBackupFilesDir() : ctx.getFilesDir();
        File f = new File(dir, "llave-dueno");
        try {
            if (f.isFile() && f.length() >= 32 && f.length() <= 256) {
                try (FileInputStream in = new FileInputStream(f)) {
                    byte[] b = new byte[(int) f.length()];
                    int n = in.read(b);
                    String leida = new String(b, 0, Math.max(0, n), StandardCharsets.US_ASCII).trim();
                    if (leida.length() >= 32) return leida;
                }
            }
            byte[] r = new byte[32];
            new SecureRandom().nextBytes(r);
            String llave = Base64.encodeToString(r, Base64.URL_SAFE | Base64.NO_PADDING | Base64.NO_WRAP);
            File tmp = new File(dir, "llave-dueno.tmp");
            try (FileOutputStream o = new FileOutputStream(tmp)) {
                o.write(llave.getBytes(StandardCharsets.US_ASCII));
                o.getFD().sync();
            }
            if (!tmp.renameTo(f)) {
                tmp.delete();
                try (FileOutputStream o = new FileOutputStream(f)) { o.write(llave.getBytes(StandardCharsets.US_ASCII)); }
            }
            return llave;
        } catch (Exception e) {
            return "";
        }
    }

    static boolean coincide(String esperada, String recibida) {
        if (esperada == null || esperada.length() < 32 || recibida == null || recibida.length() < 32) return false;
        return MessageDigest.isEqual(esperada.getBytes(StandardCharsets.UTF_8), recibida.getBytes(StandardCharsets.UTF_8));
    }
}
