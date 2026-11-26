package mx.wack.nom.core.servicio;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

import mx.wack.nom.core.crdt.Operacion;
import mx.wack.nom.core.modelo.Usuario;

/** Topes y revisiones contra el abuso de quien ya tiene cuenta (presencia, reloj, metadatos, repeticiones). */
class ProteccionesTest {

    static final class Canal implements mx.wack.nom.core.servicio.Canal {
        final List<String> eventos = new ArrayList<>();
        final List<Object> datos = new ArrayList<>();
        @Override public void enviar(String evento, Object d) { eventos.add(evento); datos.add(d); }
        @Override public void latido() { }
        @Override public void cerrar() { }
        @Override public void alCerrar(Runnable r) { }
    }

    private static Usuario usuario(String id) {
        Usuario u = new Usuario();
        u.id = id; u.nombre = id; u.color = "#123456";
        return u;
    }

    @Test
    @SuppressWarnings("unchecked")
    void nadieMueveLaPresenciaDeOtraPersona() {
        Difusor d = new Difusor(50, 0);
        Canal ana = new Canal(), beto = new Canal();
        d.suscribir(usuario("ana"), "t1", "esp", "doc", "dana-0001", ana);
        d.suscribir(usuario("beto"), "t2", "esp", "doc", "dbeto-0002", beto);
        ana.eventos.clear(); beto.eventos.clear(); ana.datos.clear(); beto.datos.clear();
        // Beto usa el cliente de Ana: no cambia nada ni se reparte nada
        d.actualizarPresencia("beto", "doc", "dana-0001", "bloque-x", "TEXTO", Map.of("a", "x", "f", "y"));
        assertTrue(ana.eventos.isEmpty() && beto.eventos.isEmpty(), "sin cambio no hay envío");
        // Ana cambia la suya; un bloque enorme no se guarda
        d.actualizarPresencia("ana", "doc", "dana-0001", "b".repeat(5000), "TEXTO", Map.of("a", "c1", "f", "c2", "z", "otro"));
        assertEquals(List.of("presencia"), beto.eventos);
        List<Map<String, Object>> us = (List<Map<String, Object>>) ((Map<String, Object>) beto.datos.get(0)).get("usuarios");
        Map<String, Object> deAna = us.stream().filter(x -> "ana".equals(x.get("id"))).findFirst().orElseThrow();
        assertNull(deAna.get("bloque"), "un bloque de más de 64 caracteres no se reparte");
        assertEquals(Map.of("a", "c1", "f", "c2"), deAna.get("cursor"), "del cursor solo a y f");
        assertEquals("", Difusor.clienteValido("x\"><b>"), "un cliente con caracteres raros se descarta");
        assertEquals("", Difusor.clienteValido("a".repeat(65)));
        assertEquals("dabc123-x9y8", Difusor.clienteValido("dabc123-x9y8"));
    }

    private static Operacion op(long t) { return new Operacion("n1:" + t, "b1", Map.of("texto", "hola"), t, "n1"); }

    @Test
    void elRelojNoSeLlevaAlTope() {
        List<Operacion> ops = List.of(op(11), op(Long.MAX_VALUE));
        long lim = ServicioDocumentos.limiteReloj(10, ops, ServicioDocumentos.HOLGURA_RELOJ);
        assertTrue(ServicioDocumentos.relojValido(op(11), lim));
        assertFalse(ServicioDocumentos.relojValido(op(Long.MAX_VALUE), lim), "una operación con el reloj al máximo no entra");
        assertFalse(ServicioDocumentos.relojValido(op(10 + ServicioDocumentos.HOLGURA_RELOJ + 1000), lim));
        assertTrue(lim < ServicioDocumentos.MAX_RELOJ);
        // un texto largo sí puede adelantar el reloj lo que mide
        Map<String, Object> s = new LinkedHashMap<>();
        s.put("k", "i"); s.put("x", "a".repeat(3_000_000));
        Operacion texto = new Operacion("n1:t", "b2", Map.of(), 20, "n1", s);
        long lim2 = ServicioDocumentos.limiteReloj(10, List.of(texto), 100);
        assertTrue(ServicioDocumentos.relojValido(texto, lim2), "las letras insertadas cuentan en lo que puede avanzar");
        assertFalse(ServicioDocumentos.relojValido(texto, ServicioDocumentos.limiteReloj(10, List.of(op(11)), 100)), "sin traer esas letras, no");
        assertEquals(ServicioDocumentos.MAX_RELOJ, ServicioDocumentos.limiteReloj(Long.MAX_VALUE, ops, ServicioDocumentos.HOLGURA_REPLICA));
    }

    @Test
    void laBaseDeUnaRamaSeGuardaAcotada() {
        Map<String, Long> base = new LinkedHashMap<>();
        for (int i = 0; i < 5000; i++) base.put("d" + i + "-abcd", (long) i);
        base.put("x".repeat(500), 1L);
        base.put("malo\"<", 1L);
        Map<String, Long> r = ServicioDocumentos.baseAcotada(base);
        assertEquals(ServicioDocumentos.MAX_NODOS_BASE, r.size());
        assertFalse(r.containsKey("malo\"<"));
        List<String> nodos = new ArrayList<>(List.of("dabc-1", "y".repeat(100), "<x>"));
        assertEquals(List.of("dabc-1"), ServicioDocumentos.nodosAcotados(nodos));
    }

    @Test
    void unaHuellaDistingueEnterosGrandes() {
        Map<String, Object> a = new LinkedHashMap<>(), b = new LinkedHashMap<>();
        a.put("letras", 1500L);
        b.put("letras", new java.math.BigInteger("18446744073709553116"));   // 2^64 + 1500
        assertFalse(Huella.sha256(a).equals(Huella.sha256(b)));
        b.put("letras", 1500.0);
        assertEquals(Huella.sha256(a), Huella.sha256(b), "un 1500.0 es el mismo 1500");
    }

    @Test
    void losTopesPorMinutoFrenanLaRepeticion() {
        Topes t = new Topes();
        for (int i = 0; i < 30; i++) t.exigir("u1", 30, "espera");
        ErrorNom e = assertThrows(ErrorNom.class, () -> t.exigir("u1", 30, "espera"));
        assertEquals(ErrorNom.DEMASIADOS, e.estado);
        t.exigir("u2", 30, "espera");   // otra persona no espera
    }
}
