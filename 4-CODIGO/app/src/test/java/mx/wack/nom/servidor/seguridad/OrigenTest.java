package mx.wack.nom.servidor.seguridad;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

/** Para los topes, una IPv6 de Internet cuenta por su red /64; lo demás se queda como llega. */
class OrigenTest {

    @Test
    void ipv6PublicaPorSuRed() {
        assertEquals("2001:db8:1:2::/64", Origen.agrupar("2001:db8:1:2:aaaa:bbbb:cccc:dddd"));
        assertEquals("2001:db8:1:2::/64", Origen.agrupar("2001:db8:1:2::1"));
        assertEquals("203.0.113.5", Origen.agrupar("203.0.113.5"));
        assertEquals("203.0.113.5", Origen.agrupar("::ffff:203.0.113.5"));
        assertEquals("fd00::5", Origen.agrupar("fd00::5"));
        assertEquals("fe80::1", Origen.agrupar("fe80::1"));
        assertEquals("no-es-ip", Origen.agrupar("no-es-ip"));
        assertEquals(null, Origen.agrupar(null));
    }
}
