package mx.wack.nom.core.servicio;

import java.util.Map;

/**
 * Atajos cortos para compartir: nom.qd.je/app baja la app de Android y nom.qd.je/codigo abre el codigo
 * fuente. La redireccion de nom.qd.je conserva la ruta, asi que llegan aqui como /app y /codigo; el
 * servidor (Spring) y el telefono que comparte (ServidorHttp) responden igual, con una redireccion
 * temporal (302): el destino se puede cambiar despues sin que el navegador recuerde el de antes.
 */
public final class Atajos {

    /** La app de Android de la ultima version publicada (la sirve la red de GitHub, no el servidor). */
    public static final String APK = "https://github.com/TruTruMaster/NOM-COLAB/releases/latest/download/NOM-COLAB.apk";
    /** El codigo fuente. */
    public static final String CODIGO = "https://github.com/TruTruMaster/NOM-COLAB";

    private static final Map<String, String> DESTINOS = Map.of("/app", APK, "/apk", APK, "/codigo", CODIGO);

    private Atajos() { }

    /** A donde lleva la ruta (con la diagonal inicial), o null si no es un atajo. */
    public static String destino(String ruta) {
        return ruta == null ? null : DESTINOS.get(ruta);
    }
}
