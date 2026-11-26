package mx.wack.nom.core.servicio;

import java.util.ArrayDeque;
import java.util.HashMap;
import java.util.Map;

/**
 * Cuántas veces por minuto puede alguien hacer algo que le cuesta al servidor (escribir en el chat, buscar dentro de
 * los documentos, copiar un espacio a otro servidor). Una persona normal nunca llega; un programa que lo repite sin
 * parar se detiene aquí en vez de ocupar el disco o el procesador de todos.
 */
public final class Topes {

    private static final long MINUTO = 60_000L;
    private final Map<String, ArrayDeque<Long>> usos = new HashMap<>();

    /** Cuenta un uso de `clave`; si ya van `porMinuto` en el último minuto, lanza DEMASIADOS (429) con el mensaje. */
    public synchronized void exigir(String clave, int porMinuto, String mensaje) {
        long ahora = System.currentTimeMillis();
        ArrayDeque<Long> l = usos.computeIfAbsent(clave, k -> new ArrayDeque<>());
        while (!l.isEmpty() && ahora - l.peekFirst() > MINUTO) l.pollFirst();
        if (l.size() >= porMinuto) throw new ErrorNom(ErrorNom.DEMASIADOS, mensaje);
        l.addLast(ahora);
        if (usos.size() > 20_000) usos.values().removeIf(q -> q.isEmpty() || ahora - q.peekLast() > MINUTO);
    }

    /** Tiempo de procesador que ya gastó `clave` en el último minuto (anotado con gastar). */
    private final Map<String, ArrayDeque<long[]>> gastos = new HashMap<>();

    /** Si `clave` ya gastó `maxMs` de procesador en el último minuto, lanza DEMASIADOS (429): se reintenta después. */
    public synchronized void exigirTiempo(String clave, long maxMs, String mensaje) {
        long ahora = System.currentTimeMillis(), suma = 0;
        ArrayDeque<long[]> l = gastos.get(clave);
        if (l == null) return;
        while (!l.isEmpty() && ahora - l.peekFirst()[0] > MINUTO) l.pollFirst();
        for (long[] g : l) suma += g[1];
        if (suma >= maxMs) throw new ErrorNom(ErrorNom.DEMASIADOS, mensaje);
    }

    public synchronized void gastar(String clave, long ms) {
        if (ms <= 0) return;
        long ahora = System.currentTimeMillis();
        gastos.computeIfAbsent(clave, k -> new ArrayDeque<>()).addLast(new long[] { ahora, ms });
        if (gastos.size() > 20_000) gastos.values().removeIf(q -> q.isEmpty() || ahora - q.peekLast()[0] > MINUTO);
    }
}
