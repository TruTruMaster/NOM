package mx.wack.nom.servidor.repositorio;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** La bitácora de un documento se lee desde el final y no crece sin fin. */
class BitacoraTest {

    @TempDir Path dir;

    @Test
    void ultimasLineasYRecorte() throws Exception {
        AlmacenJson a = new AlmacenJson(dir.toString());
        for (int i = 0; i < 2000; i++) a.agregarLinea("docs/d1/actividad.jsonl", "{\"n\":" + i + ",\"x\":\"" + "z".repeat(200) + "\"}");
        List<String> u = a.ultimasLineas("docs/d1/actividad.jsonl", 3);
        assertEquals(3, u.size());
        assertTrue(u.get(2).startsWith("{\"n\":1999"), "la última es la más nueva");
        assertTrue(u.get(0).startsWith("{\"n\":1997"));
        assertEquals(2000, a.ultimasLineas("docs/d1/actividad.jsonl", 100_000).size());
        long antes = Files.size(a.ruta("docs/d1/actividad.jsonl"));
        a.recortarLineas("docs/d1/actividad.jsonl", 100_000);
        long despues = Files.size(a.ruta("docs/d1/actividad.jsonl"));
        assertTrue(antes > 100_000 && despues <= 50_000, "se queda con lo más nuevo: " + despues);
        List<String> quedan = a.ultimasLineas("docs/d1/actividad.jsonl", 100_000);
        assertTrue(quedan.get(quedan.size() - 1).startsWith("{\"n\":1999"));
        assertTrue(quedan.get(0).startsWith("{\"n\":"), "sin una línea partida al principio");
    }
}
