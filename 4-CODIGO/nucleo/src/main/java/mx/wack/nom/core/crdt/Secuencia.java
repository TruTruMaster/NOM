package mx.wack.nom.core.crdt;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.TreeMap;


/**
 * Texto letra por letra (CRDT de secuencia tipo RGA). Gemelo de static/js/secuencia.js.
 *
 * El servidor no necesita conocer el orden de los caracteres: solo guarda, sin
 * repetir, las tres clases de operaciones (insertar tramos, mostrar/ocultar y dar
 * formato). Los clientes arman el mismo orden con esas operaciones. Para restaurar
 * una version basta saber que caracteres estaban visibles y con que formato, y eso
 * se calcula con registros LWW por caracter, sin orden.
 *
 *   ins: { t, n, iz, x, e?, a? }   el caracter i del tramo tiene id (t+i)@n
 *   vis: { t, n, v, r }            r = [[nodo, t1, t2], ...]
 *   fmt: { t, n, a, r }
 *
 * Para los parrafos con dueno (candados por persona) si hace falta el orden: se arma con el
 * mismo algoritmo RGA que el cliente, solo cuando el documento tiene algun parrafo asignado, y
 * despues se actualiza con cada operacion (sin volver a recorrer todo).
 */
public class Secuencia {

    public static final String FIN = "fin";

    public List<Map<String, Object>> ins = new ArrayList<>();
    public List<Map<String, Object>> vis = new ArrayList<>();
    public List<Map<String, Object>> fmt = new ArrayList<>();

    /** nodo -> (t inicial -> tramo), para no repetir caracteres ya guardados. */
    // Privados y sin "getters": no salen en el JSON (ni con Jackson ni con el del teléfono).
    private Map<String, TreeMap<Long, Map<String, Object>>> indice;
    private Set<String> vistas;

    // ------------------------------------------------------------------ fusionar

    /** Lo más que puede tocar una operación de mostrar/ocultar o de formato: caracteres, sumando todos sus rangos. */
    public static final long MAX_ALCANCE = 4_000_000;
    /** Mayor marca entera exacta en el navegador (2^53 - 2, para que avanzar de uno en uno nunca se quede quieto). */
    public static final long MAX_ID = 9_007_199_254_740_990L;
    /** Máximo de rangos en una sola operación de visibilidad o formato. */
    static final int MAX_RANGOS = 5000;

    /** r = [[nodo, t1, t2], …] (o [fin, 0, 0]): listas de tres, nodo corto, 0 ≤ t1 ≤ t2 ≤ MAX_ID y sin pasar del tope. */
    static boolean rangosValidos(List<?> r) {
        if (r == null || r.isEmpty() || r.size() > MAX_RANGOS) return false;
        long total = 0;
        for (Object o : r) {
            if (!(o instanceof List<?> rango) || rango.size() != 3 || !(rango.get(0) instanceof String rn) || rn.isEmpty() || rn.length() > 64
                    || !(rango.get(1) instanceof Number a) || !(rango.get(2) instanceof Number b)) return false;
            if (FIN.equals(rn)) continue;
            long t1 = a.longValue(), t2 = b.longValue();
            if (t1 < 0 || t2 < t1 || t2 > MAX_ID) return false;
            total += Math.min(t2 - t1 + 1, MAX_ALCANCE + 1);
            if (total > MAX_ALCANCE) return false;
        }
        return true;
    }

    /** Cuántos caracteres abarcan los rangos de una operación de mostrar/ocultar o de formato (0 si inserta). */
    public static long alcance(Map<String, Object> s) {
        if (s == null || !(s.get("r") instanceof List<?> r)) return 0;
        long total = 0;
        for (Object o : r) {
            if (o instanceof List<?> rango && rango.size() == 3 && !FIN.equals(rango.get(0))
                    && rango.get(1) instanceof Number a && rango.get(2) instanceof Number b
                    && a.longValue() >= 0 && b.longValue() >= a.longValue() && b.longValue() <= MAX_ID)
                total += Math.min(b.longValue() - a.longValue() + 1, MAX_ALCANCE + 1);
        }
        return total;
    }

    /** Atributos de formato: pocos, con nombres cortos y sin nombres reservados. */
    static boolean atributosValidos(Object a) {
        if (a == null) return true;
        if (!(a instanceof Map<?, ?> m) || m.size() > 200) return false;
        for (Object k : m.keySet()) if (!(k instanceof String ks) || ks.isEmpty() || ks.length() > 64 || Operacion.RESERVADOS.contains(ks)) return false;
        return true;
    }

    /** Guarda una operacion (k = i, v o f) si es nueva. Devuelve true si se guardo algo. */
    public synchronized boolean aplicar(Map<String, Object> s, long t, String n) {
        Object k = s.get("k");
        if ("i".equals(k)) return insertar(s, t, n);
        if (!("v".equals(k) || "f".equals(k)) || !(s.get("r") instanceof List<?> r) || r.isEmpty() || r.size() > 5000 || !rangosValidos(r)) return false;
        if ("f".equals(k) && (!(s.get("a") instanceof Map) || !atributosValidos(s.get("a")))) return false;
        String clave = k + ":" + t + "@" + n;
        if (!vistas().add(clave)) return false;
        Map<String, Object> e = new LinkedHashMap<>();
        e.put("k", k);
        e.put("t", t);
        e.put("n", n);
        if ("v".equals(k)) e.put("v", Boolean.TRUE.equals(s.get("v")));
        else e.put("a", s.get("a"));
        e.put("r", r);
        ("v".equals(k) ? vis : fmt).add(e);
        if ("f".equals(k) && ((Map<?, ?>) e.get("a")).containsKey("dueno")) duenos = true;
        if (orden != null) aplicarRangosOrden(e);
        return true;
    }

    private boolean insertar(Map<String, Object> s, long t, String n) {
        Object e = s.get("e");
        String x = e != null ? "￼" : s.get("x") instanceof String str ? str : null;
        if (x == null || x.isEmpty() || x.length() > 1_000_000) return false;
        if (!atributosValidos(s.get("a")) || (s.get("iz") != null && !(s.get("iz") instanceof String iz0 && iz0.length() <= 100))) return false;
        int largo = x.codePointCount(0, x.length());
        TreeMap<Long, Map<String, Object>> deNodo = indice().computeIfAbsent(n, z -> new TreeMap<>());
        // Lo ya guardado de este nodo siempre es un prefijo (cada caracter depende del anterior).
        int k = 0;
        while (k < largo && contiene(deNodo, t + k)) k++;
        if (k >= largo) return false;
        if (k > 0) {
            x = x.substring(x.offsetByCodePoints(0, k));
            t += k;
            s = new LinkedHashMap<>(s);
            s.put("iz", (t - 1) + "@" + n);
            largo -= k;
        }
        Object iz = s.get("iz");
        Object a = s.get("a");
        // Letras cifradas de extremo a extremo (fase 15): la marca se guarda para que quien abra el documento sepa descifrarlas.
        Object cx = s.get("cx") != null && !Boolean.FALSE.equals(s.get("cx")) && !Integer.valueOf(0).equals(s.get("cx")) ? 1 : null;
        // Junta con el tramo anterior del mismo nodo si es su continuacion directa (escritura normal).
        Map.Entry<Long, Map<String, Object>> previo = deNodo.floorEntry(t - 1);
        if (e == null && previo != null) {
            Map<String, Object> p = previo.getValue();
            String px = (String) p.get("x");
            long pt = ((Number) p.get("t")).longValue();
            if (p.get("e") == null && pt + px.codePointCount(0, px.length()) == t
                    && ((t - 1) + "@" + n).equals(iz) && Objects.equals(p.get("a"), a) && (p.get("cx") == null) == (cx == null)) {
                p.put("x", px + x);
                alInsertar(t, n, (String) iz, x, null, a);
                return true;
            }
        }
        Map<String, Object> r = new LinkedHashMap<>();
        r.put("k", "i");
        r.put("t", t);
        r.put("n", n);
        r.put("iz", iz);
        r.put("x", x);
        if (e != null) r.put("e", e);
        if (a != null) r.put("a", a);
        if (cx != null) r.put("cx", cx);
        ins.add(r);
        deNodo.put(t, r);
        alInsertar(t, n, iz == null ? null : String.valueOf(iz), x, e, a);
        return true;
    }

    private static boolean contiene(TreeMap<Long, Map<String, Object>> deNodo, long t) {
        Map.Entry<Long, Map<String, Object>> p = deNodo.floorEntry(t);
        if (p == null) return false;
        String x = (String) p.getValue().get("x");
        return t < p.getKey() + x.codePointCount(0, x.length());
    }

    private Map<String, TreeMap<Long, Map<String, Object>>> indice() {
        if (indice == null) {
            indice = new HashMap<>();
            for (Map<String, Object> r : ins) {
                indice.computeIfAbsent((String) r.get("n"), z -> new TreeMap<>()).put(((Number) r.get("t")).longValue(), r);
            }
        }
        return indice;
    }

    private Set<String> vistas() {
        if (vistas == null) {
            vistas = new HashSet<>();
            for (Map<String, Object> e : vis) vistas.add("v:" + e.get("t") + "@" + e.get("n"));
            for (Map<String, Object> e : fmt) vistas.add("f:" + e.get("t") + "@" + e.get("n"));
        }
        return vistas;
    }

    // ------------------------------------------------------------------ orden (candados por parrafo)

    /** Un caracter en el orden del documento (incluidos los ocultos). */
    private static final class Nodo {
        final String id, n;
        final long t;
        final boolean salto;
        final Car car = new Car();

        Nodo(String id, long t, String n, boolean salto) { this.id = id; this.t = t; this.n = n; this.salto = salto; }
    }

    // Privados y sin "getters": no salen en el JSON.
    private List<Nodo> orden;
    private Map<String, Nodo> nodos;
    private Map<String, TreeMap<Long, Nodo>> nodosPorNodo;
    private Car carFin;
    private Boolean duenos;

    /** ¿Algun parrafo tiene (o tuvo) dueno? Solo entonces se revisan los candados. */
    @SuppressWarnings("unchecked")
    public synchronized boolean tieneDuenos() {
        if (duenos == null) {
            duenos = false;
            for (Map<String, Object> e : fmt) if (e.get("a") instanceof Map && ((Map<String, Object>) e.get("a")).containsKey("dueno")) { duenos = true; break; }
        }
        return duenos;
    }

    private static final Comparator<Nodo> HERMANOS = (x, y) -> x.t != y.t ? Long.compare(y.t, x.t) : y.n.compareTo(x.n);

    /** Arma el orden completo (RGA): hijos de cada caracter de mas nuevo a mas viejo, en profundidad. */
    @SuppressWarnings("unchecked")
    private void armarOrden() {
        nodos = new HashMap<>();
        nodosPorNodo = new HashMap<>();
        Map<String, List<Nodo>> hijos = new HashMap<>();
        for (Map<String, Object> r : ins) {
            long t = ((Number) r.get("t")).longValue();
            String n = (String) r.get("n");
            String x = r.get("e") != null ? "￼" : (String) r.get("x");
            Map<String, Object> a = r.get("a") instanceof Map ? (Map<String, Object>) r.get("a") : Map.of();
            int[] cps = x.codePoints().toArray();
            for (int i = 0; i < cps.length; i++) {
                String id = (t + i) + "@" + n;
                if (nodos.containsKey(id)) continue;
                Nodo nd = new Nodo(id, t + i, n, cps[i] == '\n' && r.get("e") == null);
                nd.car.vt = t + i;
                nd.car.vn = n;
                for (var en : a.entrySet()) nd.car.f.put(en.getKey(), new Object[] { en.getValue(), t + i, n });
                nodos.put(id, nd);
                String iz = i == 0 ? (r.get("iz") == null ? "" : String.valueOf(r.get("iz"))) : (t + i - 1) + "@" + n;
                hijos.computeIfAbsent(iz, z -> new ArrayList<>()).add(nd);
            }
        }
        hijos.values().forEach(l -> l.sort(HERMANOS));
        orden = new ArrayList<>(nodos.size());
        List<Nodo> pila = new ArrayList<>(hijos.getOrDefault("", List.of()));
        java.util.Collections.reverse(pila);
        while (!pila.isEmpty()) {
            Nodo nd = pila.remove(pila.size() - 1);
            orden.add(nd);
            List<Nodo> h = hijos.get(nd.id);
            if (h != null) for (int i = h.size() - 1; i >= 0; i--) pila.add(h.get(i));
        }
        // Lo que no quedo colgado de la raiz (su vecino izquierdo no llego) no cuenta.
        if (orden.size() != nodos.size()) {
            Set<String> puestos = new HashSet<>();
            for (Nodo nd : orden) puestos.add(nd.id);
            nodos.keySet().retainAll(puestos);
        }
        for (Nodo nd : orden) nodosPorNodo.computeIfAbsent(nd.n, z -> new TreeMap<>()).put(nd.t, nd);
        carFin = new Car();
        for (Map<String, Object> o : vis) aplicarRangosOrden(o);
        for (Map<String, Object> o : fmt) aplicarRangosOrden(o);
    }

    /** Un tramo nuevo entra al orden ya armado (mismo lugar que en el cliente). */
    private void alInsertar(long t, String n, String iz, String x, Object e, Object a) {
        if (orden == null) return;
        int i;
        if (iz == null) i = 0;
        else {
            Nodo izq = nodos.get(iz);
            if (izq == null) { orden = null; nodos = null; nodosPorNodo = null; return; }   // huerfano: se rearma la proxima vez
            i = orden.indexOf(izq) + 1;
        }
        while (i < orden.size() && (orden.get(i).t > t || (orden.get(i).t == t && orden.get(i).n.compareTo(n) > 0))) i++;
        int[] cps = x.codePoints().toArray();
        List<Nodo> nuevos = new ArrayList<>(cps.length);
        @SuppressWarnings("unchecked")
        Map<String, Object> at = a instanceof Map ? (Map<String, Object>) a : Map.of();
        TreeMap<Long, Nodo> delNodo = nodosPorNodo.computeIfAbsent(n, z -> new TreeMap<>());
        for (int k = 0; k < cps.length; k++) {
            Nodo nd = new Nodo((t + k) + "@" + n, t + k, n, cps[k] == '\n' && e == null);
            nd.car.vt = t + k;
            nd.car.vn = n;
            for (var en : at.entrySet()) nd.car.f.put(en.getKey(), new Object[] { en.getValue(), t + k, n });
            nodos.put(nd.id, nd);
            delNodo.put(nd.t, nd);
            nuevos.add(nd);
        }
        orden.addAll(i, nuevos);
    }

    @SuppressWarnings("unchecked")
    private void aplicarRangosOrden(Map<String, Object> o) {
        long t = ((Number) o.get("t")).longValue();
        String n = (String) o.get("n");
        boolean esV = "v".equals(o.get("k")) || o.containsKey("v");
        Map<String, Object> a = esV ? null : (Map<String, Object>) o.get("a");
        if (!(o.get("r") instanceof List<?> rs)) return;
        for (Object ob : rs) {
            if (!(ob instanceof List<?> rango) || rango.size() != 3 || !(rango.get(0) instanceof String rn)) continue;
            if (FIN.equals(rn)) { aplicarA(carFin, esV, o, a, t, n); continue; }
            if (!(rango.get(1) instanceof Number n1) || !(rango.get(2) instanceof Number n2)) continue;
            long t1 = n1.longValue(), t2 = n2.longValue();
            if (t1 < 0 || t2 < t1 || t2 > MAX_ID) continue;
            TreeMap<Long, Nodo> delNodo = nodosPorNodo.get(rn);
            if (delNodo == null || delNodo.isEmpty()) continue;
            long lim = Math.min(t2, t1 + MAX_ALCANCE - 1);
            for (Nodo nd : delNodo.subMap(t1, true, lim, true).values()) aplicarA(nd.car, esV, o, a, t, n);
        }
    }

    private static void aplicarA(Car c, boolean esV, Map<String, Object> o, Map<String, Object> a, long t, String n) {
        if (esV) { if (gana(t, n, c.vt, c.vn)) { c.v = Boolean.TRUE.equals(o.get("v")); c.vt = t; c.vn = n; } return; }
        if (a == null) return;
        for (var e : a.entrySet()) {
            Object[] reg = c.f.get(e.getKey());
            if (reg == null || gana(t, n, ((Number) reg[1]).longValue(), (String) reg[2])) c.f.put(e.getKey(), new Object[] { e.getValue(), t, n });
        }
    }

    private static String duenoDe(Car c) {
        Object[] reg = c.f.get("dueno");
        return reg == null || reg[0] == null || "".equals(reg[0]) ? null : String.valueOf(reg[0]);
    }

    /** Dueno del parrafo que empieza a buscar en la posicion i (el primer salto visible, o el final). */
    private String duenoDesde(int i) {
        for (int k = Math.max(0, i); k < orden.size(); k++) {
            Nodo nd = orden.get(k);
            if (nd.salto && nd.car.v) return duenoDe(nd.car);
        }
        return duenoDe(carFin);
    }

    /** Marca usada cuando el texto va junto a algo que este servidor no conoce. */
    public static final String DESCONOCIDO = "?";

    /**
     * Duenos de los parrafos que tocaria una operacion (insertar, mostrar u ocultar, o dar formato),
     * calculados ANTES de aplicarla. Vacio si ningun parrafo tocado tiene dueno.
     */
    public synchronized Set<String> duenos(Map<String, Object> s) {
        if (orden == null) armarOrden();
        Set<String> r = new LinkedHashSet<>();
        Object k = s.get("k");
        if ("i".equals(k)) {
            Object iz = s.get("iz");
            if (iz == null) { String d = duenoDesde(0); if (d != null) r.add(d); return r; }
            Nodo izq = nodos.get(String.valueOf(iz));
            if (izq == null) { r.add(DESCONOCIDO); return r; }
            String d = duenoDesde(orden.indexOf(izq) + 1);
            if (d != null) r.add(d);
            return r;
        }
        if (!(s.get("r") instanceof List<?> rangos) || !rangosValidos(rangos)) return r;
        Set<Nodo> tocados = new HashSet<>();
        for (Object ob : rangos) {
            if (!(ob instanceof List<?> rango) || rango.size() != 3 || !(rango.get(0) instanceof String rn)) continue;
            if (FIN.equals(rn)) { String d = duenoDesde(orden.size()); if (d != null) r.add(d); continue; }
            if (!(rango.get(1) instanceof Number n1) || !(rango.get(2) instanceof Number n2)) continue;
            long t1 = n1.longValue(), t2 = n2.longValue();
            if (t1 < 0 || t2 < t1 || t2 > MAX_ID) continue;
            TreeMap<Long, Nodo> delNodo = nodosPorNodo.get(rn);
            if (delNodo == null || delNodo.isEmpty()) continue;
            long lim = Math.min(t2, t1 + MAX_ALCANCE - 1);
            tocados.addAll(delNodo.subMap(t1, true, lim, true).values());
        }
        if (tocados.isEmpty()) return r;
        // Una sola pasada de atras hacia adelante: cada caracter pertenece al siguiente salto visible.
        String actual = duenoDe(carFin);
        int faltan = tocados.size();
        for (int i = orden.size() - 1; i >= 0 && faltan > 0; i--) {
            Nodo nd = orden.get(i);
            if (nd.salto && nd.car.v) actual = duenoDe(nd.car);
            if (tocados.contains(nd)) { faltan--; if (actual != null) r.add(actual); }
        }
        return r;
    }

    /** Ids de los caracteres de un tramo de insercion (para descartar lo que dependa de el). */
    public static List<String> idsDe(Map<String, Object> s, long t, String n) {
        List<String> l = new ArrayList<>();
        String x = s.get("e") != null ? "￼" : s.get("x") instanceof String str ? str : "";
        int largo = x.codePointCount(0, x.length());
        for (int i = 0; i < largo; i++) l.add((t + i) + "@" + n);
        return l;
    }

    // ------------------------------------------------------------------ estado por caracter (sin orden)

    /** Visibilidad y formato de un caracter, resueltos como registros LWW. */
    public static final class Car {
        boolean v = true;
        long vt;
        String vn;
        final Map<String, Object[]> f = new HashMap<>();

        public boolean visible() { return v; }

        public Map<String, Object> attrs() {
            Map<String, Object> a = new TreeMap<>();
            f.forEach((k, reg) -> { if (reg[0] != null && !Boolean.FALSE.equals(reg[0]) && !"".equals(reg[0])) a.put(k, reg[0]); });
            return a;
        }
    }

    private static boolean gana(long t, String n, long rt, String rn) {
        return t > rt || (t == rt && n.compareTo(rn == null ? "" : rn) > 0);
    }

    /** id -> estado de cada caracter (incluye "fin"). */
    @SuppressWarnings("unchecked")
    public synchronized Map<String, Car> caracteres() {
        Map<String, Car> m = new LinkedHashMap<>();
        Map<String, TreeMap<Long, Car>> porNodo = new HashMap<>();
        Car fin = new Car();
        m.put(FIN, fin);
        for (Map<String, Object> r : ins) {
            long t = ((Number) r.get("t")).longValue();
            String n = (String) r.get("n");
            String x = r.get("e") != null ? "￼" : (String) r.get("x");
            int largo = x.codePointCount(0, x.length());
            Map<String, Object> a = r.get("a") instanceof Map ? (Map<String, Object>) r.get("a") : Map.of();
            TreeMap<Long, Car> delNodo = porNodo.computeIfAbsent(n, z -> new TreeMap<>());
            for (int i = 0; i < largo; i++) {
                String id = (t + i) + "@" + n;
                if (m.containsKey(id)) continue;
                Car c = new Car();
                c.vt = t + i;
                c.vn = n;
                for (var e : a.entrySet()) c.f.put(e.getKey(), new Object[] { e.getValue(), t + i, n });
                m.put(id, c);
                delNodo.put(t + i, c);
            }
        }
        for (Map<String, Object> o : vis) {
            long t = ((Number) o.get("t")).longValue();
            String n = (String) o.get("n");
            boolean v = Boolean.TRUE.equals(o.get("v"));
            for (Car c : enRangos(fin, porNodo, (List<Object>) o.get("r"))) {
                if (gana(t, n, c.vt, c.vn)) { c.v = v; c.vt = t; c.vn = n; }
            }
        }
        for (Map<String, Object> o : fmt) {
            long t = ((Number) o.get("t")).longValue();
            String n = (String) o.get("n");
            Map<String, Object> a = (Map<String, Object>) o.get("a");
            for (Car c : enRangos(fin, porNodo, (List<Object>) o.get("r"))) {
                for (var e : a.entrySet()) {
                    Object[] reg = c.f.get(e.getKey());
                    if (reg == null || gana(t, n, ((Number) reg[1]).longValue(), (String) reg[2])) c.f.put(e.getKey(), new Object[] { e.getValue(), t, n });
                }
            }
        }
        return m;
    }

    private static List<Car> enRangos(Car fin, Map<String, TreeMap<Long, Car>> porNodo, List<Object> r) {
        List<Car> l = new ArrayList<>();
        if (r == null) return l;
        for (Object o : r) {
            if (!(o instanceof List<?> rango) || rango.size() != 3 || !(rango.get(0) instanceof String n)) continue;
            if (FIN.equals(n)) { l.add(fin); continue; }
            if (!(rango.get(1) instanceof Number n1) || !(rango.get(2) instanceof Number n2)) continue;
            long t1 = n1.longValue(), t2 = n2.longValue();
            if (t1 < 0 || t2 < t1 || t2 > MAX_ID) continue;
            TreeMap<Long, Car> delNodo = porNodo.get(n);
            if (delNodo == null || delNodo.isEmpty()) continue;
            long lim = Math.min(t2, t1 + MAX_ALCANCE - 1);
            l.addAll(delNodo.subMap(t1, true, lim, true).values());
        }
        return l;
    }

    /**
     * 8.5 El texto que se ve, en orden (para buscar dentro de los documentos). Arma el orden aparte,
     * con el mismo algoritmo que el cliente, sin guardar nada en la secuencia.
     */
    @SuppressWarnings("unchecked")
    public synchronized String textoVisible(int max) {
        Map<String, Car> cs = caracteres();
        final class Nd { final String id, n; final long t; final int cp; Nd(String id, long t, String n, int cp) { this.id = id; this.t = t; this.n = n; this.cp = cp; } }
        Map<String, List<Nd>> hijos = new HashMap<>();
        Set<String> vistos = new HashSet<>();
        for (Map<String, Object> r : ins) {
            long t = ((Number) r.get("t")).longValue();
            String n = (String) r.get("n");
            boolean objeto = r.get("e") != null;
            int[] cps = (objeto ? " " : (String) r.get("x")).codePoints().toArray();
            for (int i = 0; i < cps.length; i++) {
                String id = (t + i) + "@" + n;
                if (!vistos.add(id)) continue;
                String iz = i == 0 ? (r.get("iz") == null ? "" : String.valueOf(r.get("iz"))) : (t + i - 1) + "@" + n;
                hijos.computeIfAbsent(iz, z -> new ArrayList<>()).add(new Nd(id, t + i, n, cps[i]));
            }
        }
        Comparator<Nd> hermanos = (x, y) -> x.t != y.t ? Long.compare(y.t, x.t) : y.n.compareTo(x.n);
        hijos.values().forEach(l -> l.sort(hermanos));
        StringBuilder sb = new StringBuilder();
        java.util.ArrayDeque<Nd> pila = new java.util.ArrayDeque<>();
        List<Nd> raiz = hijos.getOrDefault("", List.of());
        for (int i = raiz.size() - 1; i >= 0; i--) pila.push(raiz.get(i));
        while (!pila.isEmpty() && sb.length() < max) {
            Nd nd = pila.pop();
            Car c = cs.get(nd.id);
            if (c == null || c.v) sb.appendCodePoint(nd.cp);
            List<Nd> hs = hijos.get(nd.id);
            if (hs != null) for (int i = hs.size() - 1; i >= 0; i--) pila.push(hs.get(i));
        }
        return sb.toString();
    }

    /** Cuantos caracteres se ven (sin la marca final). */
    public int visibles() {
        int n = 0;
        for (var e : caracteres().entrySet()) if (!FIN.equals(e.getKey()) && e.getValue().v) n++;
        return n;
    }

    /** Copia del registro (para la foto de una version). */
    public synchronized Map<String, Object> copia() {
        Map<String, Object> c = new LinkedHashMap<>();
        c.put("ins", new ArrayList<>(ins.stream().map(LinkedHashMap::new).toList()));
        c.put("vis", new ArrayList<>(vis));
        c.put("fmt", new ArrayList<>(fmt));
        return c;
    }

    /**
     * Copia para quien solo mira (enlace para ver, proyección): lo que está oculto conserva su lugar y sus ids, pero no su
     * contenido. Cada carácter oculto se vuelve un espacio y cada objeto oculto se queda solo con su tipo; lo que se ve y
     * el orden no cambian.
     */
    @SuppressWarnings("unchecked")
    public synchronized Secuencia sinOcultos() {
        Map<String, Car> cs = caracteres();
        Map<String, Object> c = copia();
        for (Map<String, Object> r : (List<Map<String, Object>>) c.get("ins")) {
            long t = ((Number) r.get("t")).longValue();
            String n = (String) r.get("n");
            if (r.get("e") != null) {
                Car car = cs.get(t + "@" + n);
                if (car != null && !car.v) {
                    Map<String, Object> e = new LinkedHashMap<>();
                    if (r.get("e") instanceof Map<?, ?> em && em.get("tipo") instanceof String tp) e.put("tipo", tp);
                    r.put("e", e);
                    r.remove("a");
                }
                continue;
            }
            String x = (String) r.get("x");
            int[] cps = x.codePoints().toArray();
            boolean alguno = false, todos = true;
            StringBuilder sb = new StringBuilder(x.length());
            for (int i = 0; i < cps.length; i++) {
                Car car = cs.get((t + i) + "@" + n);
                if (car != null && !car.v) { sb.append(' '); alguno = true; }
                else { sb.appendCodePoint(cps[i]); todos = false; }
            }
            if (alguno) {
                r.put("x", sb.toString());
                if (todos) r.remove("a");
            }
        }
        return desde(c);
    }

    @SuppressWarnings("unchecked")
    public static Secuencia desde(Map<String, Object> copia) {
        Secuencia s = new Secuencia();
        if (copia == null) return s;
        for (String k : List.of("ins", "vis", "fmt")) {
            if (!(copia.get(k) instanceof List<?> l)) continue;
            for (Object o : l) {
                if (!(o instanceof Map<?, ?> e)) continue;
                Map<String, Object> m = (Map<String, Object>) e;
                long t = ((Number) m.get("t")).longValue();
                Map<String, Object> sin = new LinkedHashMap<>(m);
                sin.put("k", k.equals("ins") ? "i" : k.equals("vis") ? "v" : "f");
                s.aplicar(sin, t, (String) m.get("n"));
            }
        }
        return s;
    }

    // ------------------------------------------------------------------ restaurar (CU-23)

    /**
     * Operaciones (s) que dejan el texto como estaba en 'version': se ocultan los
     * caracteres que no existian o estaban borrados, se revelan los que estaban, y se
     * devuelve a cada caracter su formato. El orden no cambia porque los ids no cambian.
     */
    public List<Map<String, Object>> restaurar(Secuencia version) {
        Map<String, Car> ahora = caracteres(), antes = version.caracteres();
        Map<String, Map<String, Object>> formatos = new HashMap<>();
        List<String> ocultar = new ArrayList<>(), mostrar = new ArrayList<>();
        Map<String, List<String>> porFormato = new LinkedHashMap<>();
        for (var e : ahora.entrySet()) {
            String id = e.getKey();
            Car c = e.getValue(), v = antes.get(id);
            boolean debeVerse = v != null && v.v;
            if (!FIN.equals(id)) {
                if (c.v && !debeVerse) ocultar.add(id);
                else if (!c.v && debeVerse) mostrar.add(id);
            }
            if (!debeVerse && !FIN.equals(id)) continue;
            Map<String, Object> objetivo = v == null ? Map.of() : v.attrs(), actual = c.attrs();
            if (objetivo.equals(actual)) continue;
            Map<String, Object> a = new TreeMap<>();
            for (String k : actual.keySet()) a.put(k, null);
            a.putAll(objetivo);
            porFormato.computeIfAbsent(String.valueOf(a), z -> new ArrayList<>()).add(id);
            formatos.put(String.valueOf(a), a);
        }
        List<Map<String, Object>> ops = new ArrayList<>();
        agregarPartidas(ops, "v", false, null, rangos(ocultar));
        agregarPartidas(ops, "v", true, null, rangos(mostrar));
        porFormato.forEach((k, ids) -> agregarPartidas(ops, "f", false, formatos.get(k), rangos(ids)));
        return ops;
    }

    private static void agregarPartidas(List<Map<String, Object>> ops, String k, boolean v, Map<String, Object> a, List<List<Object>> rs) {
        for (int i = 0; i < rs.size(); i += MAX_RANGOS) {
            Map<String, Object> s = new LinkedHashMap<>();
            s.put("k", k);
            if ("v".equals(k)) s.put("v", v); else s.put("a", a);
            s.put("r", new ArrayList<>(rs.subList(i, Math.min(rs.size(), i + MAX_RANGOS))));
            ops.add(s);
        }
    }

    /** ids "t@n" -> [[n, t1, t2], ...] juntando consecutivos del mismo nodo. */
    static List<List<Object>> rangos(List<String> ids) {
        Map<String, List<Long>> porNodo = new TreeMap<>();
        List<List<Object>> r = new ArrayList<>();
        for (String id : ids) {
            if (FIN.equals(id)) { r.add(List.of(FIN, 0, 0)); continue; }
            int at = id.indexOf('@');
            porNodo.computeIfAbsent(id.substring(at + 1), z -> new ArrayList<>()).add(Long.parseLong(id.substring(0, at)));
        }
        porNodo.forEach((n, ts) -> {
            ts.sort(null);
            long a = ts.get(0), b = a;
            for (int i = 1; i < ts.size(); i++) {
                if (ts.get(i) == b + 1) b++;
                else { r.add(List.of(n, a, b)); a = b = ts.get(i); }
            }
            r.add(List.of(n, a, b));
        });
        return r;
    }
}
