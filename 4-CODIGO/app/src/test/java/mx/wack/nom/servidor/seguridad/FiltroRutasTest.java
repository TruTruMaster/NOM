package mx.wack.nom.servidor.seguridad;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

/** Solo se atienden direcciones que Tomcat, Spring y los filtros leen igual. */
class FiltroRutasTest {

    @Test
    void rechazaLoQueSeLeeDistinto() {
        for (String ok : new String[] { "/", "/index.html", "/api/red/info", "/api/espacios/Ab3_x-9Zq1Lm/documentos", "/ayuda/cuentas/entrar",
                "/js/nucleo.js", "/api/espacios/abc/binarios/0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef", "/api/red/%69nfo" })
            assertTrue(FiltroRutas.valida(ok), ok);
        for (String malo : new String[] { "/api;x/red/info", "/api/red/info;jsessionid=1", "//api/red/info", "/api/./red/info", "/api/proyeccion/../espacios",
                "/api/documentos/x/..", "/api/%2e%2e/espacios", "/api/a%2Fb", "/api/a%5cb", "/api\\red", "/api/a%252e", "/api/a%00", "/api/a%3bb", "api", "" })
            assertFalse(FiltroRutas.valida(malo), malo);
    }

    @Test
    void soloNombresDeNom() {
        java.util.Set<String> c = FiltroRutas.conocidos("https://nomcolab.dpdns.org", "https://www.nom.qd.je", null);
        for (String ok : new String[] { "www.nom.qd.je", "nom.qd.je", "nomcolab.dpdns.org", "www.nomcolab.dpdns.org", "WWW.NOM.QD.JE", "127.0.0.1:8080",
                "192.168.1.20:8080", "[::1]:8080", "[fe80::1%25wlan0]:8080", "localhost:8096", "laptop-de-ana:8080", "laptop.local:8080", "pc.lan", "nom.qd.je." })
            assertTrue(FiltroRutas.hostPermitido(ok, c), ok);
        for (String malo : new String[] { "ataque.example.com", "127.0.0.1.nip.io", "nom.qd.je.ataque.com", "evil.box", "a b", "x\"y" })
            assertFalse(FiltroRutas.hostPermitido(malo, c), malo);
    }
}
