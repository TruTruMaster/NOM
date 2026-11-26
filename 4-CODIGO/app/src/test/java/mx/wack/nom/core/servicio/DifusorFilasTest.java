package mx.wack.nom.core.servicio;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import java.util.Map;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import mx.wack.nom.core.modelo.Usuario;

/**
 * Cada canal en vivo tiene su propia fila: quien deja de leer no detiene a quien escribe ni a los demás, se le corta si
 * se le junta demasiado o si una escritura lleva mucho atorada, y lo que se le manda llega en orden.
 */
class DifusorFilasTest {

    /** Un canal de prueba que puede quedarse «atorado» al escribir (como un socket que nadie lee). */
    static final class Canal2 implements Canal {
        final List<String> eventos = new CopyOnWriteArrayList<>();
        final CountDownLatch soltar = new CountDownLatch(1);
        volatile boolean atorado, cerrado;
        private volatile Runnable alCerrar;
        @Override public void enviar(String evento, Object datos) throws java.io.IOException {
            if (atorado) {
                try { soltar.await(); } catch (InterruptedException e) { Thread.currentThread().interrupt(); }
                if (cerrado) throw new java.io.IOException("cerrado");
            }
            eventos.add(evento + (datos instanceof Map<?, ?> m && m.get("n") != null ? ":" + m.get("n") : ""));
        }
        @Override public void latido() throws java.io.IOException { enviar("latido", null); }
        @Override public void cerrar() { cerrado = true; soltar.countDown(); Runnable r = alCerrar; if (r != null) r.run(); }
        @Override public void alCerrar(Runnable r) { alCerrar = r; }
    }

    private static Usuario u(String id) { Usuario x = new Usuario(); x.id = id; x.nombre = id; x.color = "#123456"; return x; }

    private static void esperar(java.util.function.BooleanSupplier c) throws InterruptedException {
        long fin = System.currentTimeMillis() + 3000;
        while (!c.getAsBoolean() && System.currentTimeMillis() < fin) Thread.sleep(10);
    }

    @AfterEach
    void restaurar() { Difusor.MAX_ESCRITURA_MS = 30_000L; }

    @Test
    void quienNoLeeNoDetieneAQuienEscribeNiALosDemas() throws Exception {
        Difusor d = new Difusor(50, 4);
        Canal2 lento = new Canal2(), bien = new Canal2();
        d.suscribir(u("lento"), "t1", "esp", "doc", "c1", lento);
        d.suscribir(u("bien"), "t2", "esp", "doc", "c2", bien);
        esperar(() -> lento.eventos.contains("hola") && bien.eventos.contains("hola"));
        lento.atorado = true;
        long t0 = System.nanoTime();
        for (int i = 1; i <= 3; i++) d.ops("doc", Map.of("n", i));
        long ms = (System.nanoTime() - t0) / 1_000_000;
        assertTrue(ms < 500, "quien escribe no espera (" + ms + " ms)");
        esperar(() -> bien.eventos.contains("ops:3"));
        List<String> ops = bien.eventos.stream().filter(e -> e.startsWith("ops")).toList();
        assertEquals(List.of("ops:1", "ops:2", "ops:3"), ops, "los demás reciben todo y en orden");
        lento.cerrar();
    }

    @Test
    void siSeLeJuntaDemasiadoSeCorta() throws Exception {
        Difusor d = new Difusor(50, 2);
        Canal2 lento = new Canal2();
        d.suscribir(u("lento"), "t1", "esp", "doc", "c1", lento);
        esperar(() -> lento.eventos.contains("hola"));
        lento.atorado = true;
        for (int i = 0; i < Difusor.MAX_FILA + 20; i++) d.ops("doc", Map.of("n", i));
        esperar(() -> lento.cerrado);
        assertTrue(lento.cerrado, "el canal que no lee se cierra");
        assertFalse(d.usuariosEn("esp").contains("lento"), "y sale de la presencia");
    }

    @Test
    void unaEscrituraAtoradaSeCortaEnElLatido() throws Exception {
        Difusor d = new Difusor(50, 2);
        Canal2 lento = new Canal2();
        d.suscribir(u("lento"), "t1", "esp", "doc", "c1", lento);
        esperar(() -> lento.eventos.contains("hola"));
        lento.atorado = true;
        d.ops("doc", Map.of("n", 1));
        Thread.sleep(150);
        Difusor.MAX_ESCRITURA_MS = 50;
        d.latido();
        esperar(() -> lento.cerrado);
        assertTrue(lento.cerrado);
        assertFalse(d.usuariosEn("esp").contains("lento"));
    }

    @Test
    void elAvisoDeExpulsionLlegaAntesDeCerrar() throws Exception {
        Difusor d = new Difusor(50, 2);
        Canal2 c = new Canal2();
        d.suscribir(u("beto"), "t1", "esp", null, "c1", c);
        d.expulsar("esp", "beto");
        esperar(() -> c.cerrado);
        assertTrue(c.cerrado);
        assertTrue(c.eventos.contains("expulsado"), "le llegó por qué se cerró: " + c.eventos);
        assertTrue(c.eventos.indexOf("hola") < c.eventos.indexOf("expulsado"));
        assertTrue(TimeUnit.SECONDS.toMillis(1) > 0);
    }
}
