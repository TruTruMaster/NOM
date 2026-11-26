package mx.wack.nom.core.crdt;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

/** La copia para quien solo mira se ve igual, pero lo borrado ya no trae su texto. */
class SecuenciaSinOcultosTest {

    @Test
    void loBorradoNoViaja() {
        Secuencia s = new Secuencia();
        Map<String, Object> i = new LinkedHashMap<>();
        i.put("k", "i"); i.put("iz", ""); i.put("x", "hola SECRETO mundo");
        assertTrue(s.aplicar(i, 1, "ana"));
        Map<String, Object> v = new LinkedHashMap<>();
        // "SECRETO " son los caracteres 6..13 del tramo (ids 6@ana..13@ana)
        v.put("k", "v"); v.put("v", false); v.put("r", List.of(List.of("ana", 6L, 13L)));
        assertTrue(s.aplicar(v, 100, "ana"));
        Map<String, Object> j = new LinkedHashMap<>();
        j.put("k", "i"); j.put("iz", "18@ana"); j.put("x", "CLAVE");
        assertTrue(s.aplicar(j, 200, "beto"));
        Map<String, Object> w = new LinkedHashMap<>();
        w.put("k", "v"); w.put("v", false); w.put("r", List.of(List.of("beto", 200L, 204L)));
        assertTrue(s.aplicar(w, 300, "beto"));

        Secuencia m = s.sinOcultos();
        assertEquals(s.textoVisible(1000), m.textoVisible(1000));
        assertEquals("hola mundo", m.textoVisible(1000));
        String json = String.valueOf(m.copia().get("ins"));
        assertFalse(json.contains("SECRETO"), json);
        assertFalse(json.contains("CLAVE"), json);
        assertEquals(s.visibles(), m.visibles());
    }
}
