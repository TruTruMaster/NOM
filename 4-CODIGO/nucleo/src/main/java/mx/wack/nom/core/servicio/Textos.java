package mx.wack.nom.core.servicio;

/**
 * Nombres que escribe la gente (espacios, documentos, carpetas, escuela o empresa, archivos del chat): sin caracteres
 * de control ni marcas que cambian la dirección del texto (con ellas un nombre se ve distinto de lo que es), y
 * recortados sin partir un emoji ni otro carácter de dos unidades.
 */
public final class Textos {

    private Textos() { }

    /** ¿Es una marca que cambia la dirección del texto (incrustar, forzar o aislar) o un carácter de control? */
    static boolean quitar(int c) {
        return Character.getType(c) == Character.CONTROL
                || (c >= 0x202A && c <= 0x202E)    // incrustar y forzar izquierda/derecha
                || (c >= 0x2066 && c <= 0x2069);   // aislar
    }

    /** El texto sin caracteres de control ni marcas de dirección (las marcas sencillas LRM/RLM se quedan). */
    public static String limpio(String s) {
        if (s == null) return null;
        int i = 0, n = s.length();
        while (i < n) {
            int c = s.codePointAt(i);
            if (quitar(c)) break;
            i += Character.charCount(c);
        }
        if (i >= n) return s;
        StringBuilder sb = new StringBuilder(n);
        sb.append(s, 0, i);
        while (i < n) {
            int c = s.codePointAt(i);
            if (!quitar(c)) sb.appendCodePoint(c);
            i += Character.charCount(c);
        }
        return sb.toString();
    }

    /** Los primeros `max` caracteres (unidades de Java) sin dejar medio carácter al final. */
    public static String cortar(String s, int max) {
        if (s == null || s.length() <= max) return s;
        if (max <= 0) return "";
        int fin = max;
        if (Character.isHighSurrogate(s.charAt(fin - 1))) fin--;
        return s.substring(0, fin);
    }
}
