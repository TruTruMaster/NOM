package mx.wack.nom.core.crdt;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

/** Qué cambios acepta el servidor: nombres de bloque y de campo, y cuánto puede abarcar un cambio de texto. */
class OperacionNombresTest {

    private static Operacion op(String bloque, Map<String, Object> campos) { return new Operacion("x:1", bloque, campos, 5, "n1"); }

    @Test
    void nombresReservadosNo() {
        assertTrue(op("b1", Map.of("texto", "hola")).valida());
        for (String r : List.of("__proto__", "constructor", "prototype")) {
            assertFalse(op(r, Map.of("texto", "hola")).valida(), "bloque " + r);
            assertFalse(op("b1", Map.of(r, "hola")).valida(), "campo " + r);
            assertFalse(new Operacion("x:2", "b1", Map.of(), 5, r, Map.of("k", "i", "x", "a"), null).valida(), "nodo " + r);
        }
        assertFalse(new Operacion("x".repeat(201), "b1", Map.of("texto", "a"), 5, "n1").valida());
    }

    @Test
    void rangosDeTexto() {
        Secuencia s = new Secuencia();
        assertTrue(s.aplicar(Map.of("k", "i", "x", "hola"), 1, "n1"));
        assertTrue(s.aplicar(Map.of("k", "v", "v", false, "r", List.of(List.of("n1", 1, 2))), 10, "n2"));
        assertTrue(s.aplicar(Map.of("k", "f", "a", Map.of("b", true), "r", List.of(List.of("n1", 1, 4), List.of("fin", 0, 0))), 11, "n2"));
        // rangos mal hechos o desmedidos
        assertFalse(s.aplicar(Map.of("k", "v", "v", false, "r", List.of(List.of("n1", "1", 2))), 12, "n2"));
        assertFalse(s.aplicar(Map.of("k", "v", "v", false, "r", List.of(List.of("n1", 5, 2))), 13, "n2"));
        assertFalse(s.aplicar(Map.of("k", "v", "v", false, "r", List.of(List.of("n1", 1))), 14, "n2"));
        assertFalse(s.aplicar(Map.of("k", "v", "v", false, "r", List.of(List.of("n1", 1, 3_000_000), List.of("n1", 1, 3_000_000))), 15, "n2"));
        long t0 = System.nanoTime();
        assertFalse(s.aplicar(Map.of("k", "v", "v", false, "r", java.util.Collections.nCopies(5000, List.of("n9", 1, 2_000_000))), 16, "n2"));
        assertTrue((System.nanoTime() - t0) / 1_000_000 < 500);
        // atributos con nombres reservados
        assertFalse(s.aplicar(Map.of("k", "f", "a", Map.of("__proto__", 1), "r", List.of(List.of("n1", 1, 1))), 17, "n2"));
        assertFalse(s.aplicar(Map.of("k", "i", "x", "b", "a", Map.of("constructor", 1)), 20, "n1"));
        assertTrue(s.textoVisible(100).contains("ho") || s.textoVisible(100).contains("la"));
    }
}
