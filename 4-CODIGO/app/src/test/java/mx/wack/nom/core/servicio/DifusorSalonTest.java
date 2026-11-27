package mx.wack.nom.core.servicio;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

import mx.wack.nom.core.modelo.Usuario;

/**
 * Un salón entero en un espacio: entran 50 personas a la vez; la 51 recibe un aviso que se entiende; quien ya está
 * puede abrir otra pestaña; y en cuanto alguien sale, hay lugar otra vez.
 */
class DifusorSalonTest {

    private static Usuario u(String id) { Usuario x = new Usuario(); x.id = id; x.nombre = id; x.color = "#123456"; return x; }

    @Test
    void cincuentaPersonasYLaSiguienteEsperaLugar() {
        Difusor d = new Difusor(Difusor.MAX_PERSONAS, 0);
        DifusorFilasTest.Canal2[] canales = new DifusorFilasTest.Canal2[50];
        for (int i = 0; i < 50; i++) d.suscribir(u("al" + i), "esp", "doc", "c" + i, canales[i] = new DifusorFilasTest.Canal2());
        ErrorNom e = assertThrows(ErrorNom.class, () -> d.suscribir(u("al50"), "esp", "doc", "c50", new DifusorFilasTest.Canal2()));
        assertEquals(ErrorNom.CONFLICTO, e.estado);
        assertTrue(e.getMessage().startsWith("El espacio está lleno: ya hay 50 personas conectadas"), e.getMessage());
        // Quien ya está conectado abre otra pestaña o su celular: no cuenta como otra persona.
        d.suscribir(u("al7"), "esp", "doc", "c7b", new DifusorFilasTest.Canal2());
        // Otro espacio no se ve afectado.
        d.suscribir(u("al50"), "otro", null, "c50", new DifusorFilasTest.Canal2());
        // Sale una persona (sus dos canales) y entra la que esperaba.
        canales[3].cerrar();
        d.suscribir(u("al50"), "esp", "doc", "c50b", new DifusorFilasTest.Canal2());
    }
}
