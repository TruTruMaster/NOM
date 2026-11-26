package mx.wack.nom.core.servicio;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

import mx.wack.nom.core.modelo.Usuario;

/** Quien solo puede ver recibe en vivo lo mismo que una pantalla de proyección; quien edita, todo. */
class DifusorLectoresTest {

    private static Usuario usuario(String id) {
        Usuario u = new Usuario();
        u.id = id; u.nombre = id; u.color = "#123456";
        return u;
    }

    @Test
    void quienSoloMiraRecibeLaCopiaLimpia() {
        Difusor d = new Difusor(50, 0);
        d.quienSoloMira((espacio, usuario) -> "lector".equals(usuario));
        ProteccionesTest.Canal editor = new ProteccionesTest.Canal(), lector = new ProteccionesTest.Canal();
        d.suscribir(usuario("editor"), "t1", "esp", "doc", "deditor-01", editor);
        d.suscribir(usuario("lector"), "t2", "esp", "doc", "dlector-02", lector);
        assertTrue(d.hayMirones("doc"));
        editor.eventos.clear(); editor.datos.clear(); lector.eventos.clear(); lector.datos.clear();

        Map<String, Object> todo = Map.of("ops", List.of("con comentario"), "cliente", "x");
        Map<String, Object> limpio = Map.of("ops", List.of("sin comentario"), "cliente", "x");
        d.ops("doc", todo, limpio);
        assertEquals(List.of("ops"), editor.eventos);
        assertEquals(todo, editor.datos.get(0));
        assertEquals(List.of("ops"), lector.eventos);
        assertEquals(limpio, lector.datos.get(0), "quien solo mira no recibe el lote completo");

        // Algo borrado vuelve a verse: quien solo mira vuelve a pedir su copia (no le llega el texto en el lote).
        editor.eventos.clear(); lector.eventos.clear(); lector.datos.clear();
        d.ops("doc", todo, Map.of("recargar", true));
        assertEquals(List.of("ops"), editor.eventos);
        assertEquals(List.of("recargar"), lector.eventos);

        // Nada visible (p. ej. solo un comentario): a quien solo mira no le llega nada.
        lector.eventos.clear();
        d.ops("doc", todo, null);
        assertTrue(lector.eventos.isEmpty());
    }

    @Test
    void siNoSeSabeElPapelVaLaCopiaLimpia() {
        Difusor d = new Difusor(50, 0);
        d.quienSoloMira((espacio, usuario) -> { throw new IllegalStateException("el espacio ya no existe"); });
        ProteccionesTest.Canal c = new ProteccionesTest.Canal();
        d.suscribir(usuario("ana"), "t1", "esp", "doc", "dana-01", c);
        c.eventos.clear(); c.datos.clear();
        d.ops("doc", Map.of("ops", List.of("todo")), null);
        assertTrue(c.eventos.isEmpty());
    }

    @Test
    void sinLectoresNiPantallasNoSeCalculaLaCopia() {
        Difusor d = new Difusor(50, 0);
        d.quienSoloMira((espacio, usuario) -> false);
        d.suscribir(usuario("ana"), "t1", "esp", "doc", "dana-01", new ProteccionesTest.Canal());
        assertFalse(d.hayMirones("doc"));
    }
}
