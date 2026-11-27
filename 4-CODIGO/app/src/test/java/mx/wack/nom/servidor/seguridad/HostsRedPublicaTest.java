package mx.wack.nom.servidor.seguridad;

import static org.junit.jupiter.api.Assertions.assertEquals;

import mx.wack.nom.core.servicio.Hosts;
import org.junit.jupiter.api.Test;

/** El teléfono que comparte agrupa igual que la computadora: mismos casos que OrigenTest. */
class HostsRedPublicaTest {

    @Test
    void igualQueOrigenAgrupar() {
        String[] casos = { "2001:db8:1:2:aaaa:bbbb:cccc:dddd", "2001:db8:1:2::1", "203.0.113.5", "::ffff:203.0.113.5",
                "fd00::5", "fe80::1", "fe80::1%wlan0", "no-es-ip", "192.168.1.20", null };
        for (String c : casos) assertEquals(Origen.agrupar(c == null ? null : c.replaceFirst("%.*$", "")),
                c != null && c.contains("%") ? Hosts.agruparRedPublica(c).replaceFirst("%.*$", "") : Hosts.agruparRedPublica(c), "caso " + c);
    }

    @Test
    void dosEquiposDeLaMismaRedIpv6CuentanJuntos() {
        assertEquals(Hosts.agruparRedPublica("2806:2a0:1232:83bd::10"), Hosts.agruparRedPublica("2806:2a0:1232:83bd:1:2:3:4"));
    }

    @Test
    void dosEquiposDelMismoWifiNoSeJuntan() {
        assertEquals("fe80::1", Hosts.agruparRedPublica("fe80::1"));
        assertEquals("fe80::2", Hosts.agruparRedPublica("fe80::2"));
    }
}
