package mx.wack.nom.core.crdt;

import java.util.Map;

/**
 * Unica forma de cambiar un documento: asignar campos de un bloque.
 * Crear un bloque es asignar sus campos; borrarlo es asignar eliminado=true
 * (borrado logico, nunca DELETE: ver regla 3c de CU-21).
 *
 * El texto letra por letra va en 's' (insertar, mostrar/ocultar o dar formato a
 * caracteres, ver Secuencia); en ese caso 'campos' va vacio.
 *
 * @param id     idNodo + ":" + contador. Hace la sincronizacion idempotente.
 * @param bloque id del bloque; "_doc" para los campos del documento (nombre, modo, paginas).
 * @param campos campo -> valor nuevo
 * @param t      reloj de Lamport del nodo que la genero
 * @param n      id del nodo que la genero
 * @param s      operacion de secuencia (opcional)
 * @param m      datos de la bitacora que no cambian el documento (opcional): { o: "p" } si el
 *               texto vino de pegar, con g = id del pegado y gn = cuantas letras tenia;
 *               { o: "c" } si es una cita insertada con su fuente (fase 9) y { o: "r" } si solo
 *               reacomoda lo que ya estaba (insertar o borrar filas, ordenar): no es tecleado.
 */
public record Operacion(String id, String bloque, Map<String, Object> campos, long t, String n, Map<String, Object> s,
                        Map<String, Object> m) {

    public static final String DOC = "_doc";

    public Operacion(String id, String bloque, Map<String, Object> campos, long t, String n) {
        this(id, bloque, campos, t, n, null, null);
    }

    public Operacion(String id, String bloque, Map<String, Object> campos, long t, String n, Map<String, Object> s) {
        this(id, bloque, campos, t, n, s, null);
    }

    public boolean esSecuencia() { return s != null; }

    /** El contenido vino de pegar (Ctrl+V, "Pegar" de la ruedita o el teclado del celular). */
    public boolean pegada() { return m != null && "p".equals(m.get("o")); }

    /** Cita textual insertada desde una fuente (panel Investigar o fichas). */
    public boolean citada() { return m != null && "c".equals(m.get("o")); }

    /** Solo mueve lo que ya estaba (insertar o borrar filas y columnas, ordenar) o lo armó la app
        (portada, preguntas guía y referencias del reporte): no lo escribió nadie. */
    public boolean reacomodo() { return m != null && "r".equals(m.get("o")); }

    /** Nombres que no pueden ser de un bloque ni de un campo (en la página son parte de todos los objetos). */
    public static final java.util.Set<String> RESERVADOS = java.util.Set.of("__proto__", "constructor", "prototype");

    public static boolean nombreValido(String k) { return k != null && !k.isEmpty() && k.length() <= 200 && !RESERVADOS.contains(k); }

    /** Elementos (de listas y mapas) que puede traer una operación entre todos sus valores. */
    public static final int MAX_ELEMENTOS = 100_000;

    /** Cuenta los elementos de un valor hasta pasar del tope (devuelve tope + 1 si se pasa). */
    static int elementos(Object v, int tope) {
        int n = 0;
        java.util.ArrayDeque<Object> pila = new java.util.ArrayDeque<>();
        if (v != null) pila.push(v);
        while (!pila.isEmpty() && n <= tope) {
            Object x = pila.pop();
            if (x instanceof Map<?, ?> m) { n += m.size(); for (Object y : m.values()) if (y instanceof Map || y instanceof java.util.Collection) pila.push(y); }
            else if (x instanceof java.util.Collection<?> c) { n += c.size(); for (Object y : c) if (y instanceof Map || y instanceof java.util.Collection) pila.push(y); }
        }
        return n;
    }

    public boolean valida() {
        boolean cuerpo = s != null ? s.get("k") instanceof String && !DOC.equals(bloque)
                : campos != null && !campos.isEmpty() && campos.size() <= 400 && campos.keySet().stream().allMatch(Operacion::nombreValido);
        return cuerpo && elementos(campos, MAX_ELEMENTOS) + elementos(s, MAX_ELEMENTOS) <= MAX_ELEMENTOS
                && id != null && !id.isBlank() && id.length() <= 200 && bloque != null && !bloque.isBlank()
                && n != null && !n.isBlank() && n.length() <= 64 && t > 0 && bloque.length() <= 64 && !RESERVADOS.contains(bloque)
                && !RESERVADOS.contains(n);
    }
}
