package mx.wack.nom.core.servicio;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

import org.junit.jupiter.api.Test;

import mx.wack.nom.core.modelo.Usuario;

/** El latido cierra los canales en vivo cuya sesión ya terminó y deja los demás (con la presencia al día). */
class DifusorSesionesTest {

    /** Un canal de prueba: guarda lo que recibe y si se cerró. */
    static final class CanalPrueba implements Canal {
        final List<String> eventos = new ArrayList<>();
        boolean cerrado;
        private Runnable alCerrar;
        @Override public void enviar(String evento, Object datos) { if (!cerrado) eventos.add(evento); }
        @Override public void latido() { }
        @Override public void cerrar() { cerrado = true; if (alCerrar != null) alCerrar.run(); }
        @Override public void alCerrar(Runnable r) { alCerrar = r; }
    }

    private static Usuario usuario(String id) {
        Usuario u = new Usuario();
        u.id = id; u.nombre = id; u.color = "#123456";
        return u;
    }

    @Test
    void cierraSoloLosCanalesDeSesionesTerminadas() {
        Difusor d = new Difusor(50, 0);
        Set<String> abiertas = ConcurrentHashMap.newKeySet();
        abiertas.add("sesion-ana");
        d.revisarSesionesCon(abiertas::contains);
        CanalPrueba ana = new CanalPrueba(), beto = new CanalPrueba(), proyector = new CanalPrueba(), sinSesion = new CanalPrueba();
        d.suscribir(usuario("ana"), "sesion-ana", "esp", "doc", "c1", ana);
        d.suscribir(usuario("beto"), "sesion-beto", "esp", "doc", "c2", beto);
        d.suscribir(usuario("carla"), "esp", null, "c3", sinSesion);   // (una API que no pasa la sesión: no se revisa)
        d.suscribirProyector("esp", "doc", "ABC123", "ana", proyector);
        assertEquals(Set.of("ana", "beto", "carla"), d.usuariosEn("esp"));
        ana.eventos.clear();

        d.latido();

        assertFalse(ana.cerrado, "la sesión de Ana sigue abierta");
        assertTrue(beto.cerrado, "la de Beto terminó: su canal se cierra");
        assertFalse(sinSesion.cerrado);
        assertFalse(proyector.cerrado, "las pantallas de proyección no tienen sesión");
        assertEquals(Set.of("ana", "carla"), d.usuariosEn("esp"), "Beto sale de la presencia");
        assertTrue(ana.eventos.contains("presencia"), "a Ana (en el documento) le llega quién sigue ahí");
    }

    @Test
    void siNoSePuedeRevisarElCanalSeQueda() {
        Difusor d = new Difusor(50, 0);
        d.revisarSesionesCon(t -> { throw new IllegalStateException("la base no contesta"); });
        CanalPrueba c = new CanalPrueba();
        d.suscribir(usuario("ana"), "sesion-ana", "esp", null, "c1", c);
        d.latido();
        assertFalse(c.cerrado);
    }

    @Test
    void conMuchasSesionesSeRevisaUnaParteEnCadaLatido() {
        Difusor d = new Difusor(50, 0);
        List<String> revisadas = new ArrayList<>();
        d.revisarSesionesCon(t -> { revisadas.add(t); return true; });
        for (int i = 0; i < Difusor.MAX_REVISIONES + 50; i++)
            d.suscribir(usuario("u" + (i % 40)), "s" + i, "esp" + i, null, "c" + i, new CanalPrueba());
        d.latido();
        assertTrue(revisadas.size() <= Difusor.MAX_REVISIONES, "como mucho " + Difusor.MAX_REVISIONES + " por latido");
    }
}
