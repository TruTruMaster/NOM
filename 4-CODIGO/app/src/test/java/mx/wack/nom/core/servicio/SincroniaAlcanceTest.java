package mx.wack.nom.core.servicio;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.net.InetAddress;

import org.junit.jupiter.api.Test;

/**
 * La copia en otro NOM llega hasta donde llega quien la configura. Solo IP escritas
 * (InetAddress.getByName con una IP no consulta el DNS), así la prueba no depende de la red.
 */
class SincroniaAlcanceTest {

    private static InetAddress ip(String s) throws Exception { return InetAddress.getByName(s); }

    @Test
    void midiendoQueTanLejosQuedaCadaIp() throws Exception {
        // 0: este mismo equipo y direcciones especiales
        for (String s : new String[] { "127.0.0.1", "127.8.9.10", "0.0.0.0", "::1", "::", "169.254.169.254", "fe80::1",
                "224.0.0.1", "255.255.255.255", "::ffff:127.0.0.1", "::127.0.0.1", "64:ff9b::7f00:1", "64:ff9b::a9fe:a9fe" })
            assertEquals(0, ServicioSincronia.distancia(ip(s)), s);
        // 1: redes privadas (la casa, el salón, Tailscale/CGNAT, IPv6 ULA)
        for (String s : new String[] { "10.0.0.5", "172.16.0.1", "172.31.255.255", "192.168.0.1", "100.64.0.1",
                "100.120.5.51", "fd7a:115c:a1e0::1", "::ffff:192.168.1.1", "64:ff9b::c0a8:101", "192.0.2.1" })
            assertEquals(1, ServicioSincronia.distancia(ip(s)), s);
        // 2: Internet
        for (String s : new String[] { "8.8.8.8", "104.21.87.200", "172.32.0.1", "100.128.0.1", "2606:4700:3037::ac43:91ef",
                "::ffff:8.8.8.8", "64:ff9b::808:808" })
            assertEquals(2, ServicioSincronia.distancia(ip(s)), s);
    }

    @Test
    void elAlcanceDependeDeDondeEstaQuienPide() {
        assertEquals(ServicioSincronia.EQUIPO, ServicioSincronia.alcanceDe(null));      // la petición salió de aquí
        assertEquals(ServicioSincronia.EQUIPO, ServicioSincronia.alcanceDe(" "));
        assertEquals(ServicioSincronia.RED, ServicioSincronia.alcanceDe("192.168.0.50"));
        assertEquals(ServicioSincronia.RED, ServicioSincronia.alcanceDe("100.77.150.52"));
        assertEquals(ServicioSincronia.INTERNET, ServicioSincronia.alcanceDe("8.8.8.8"));
        assertEquals(ServicioSincronia.INTERNET, ServicioSincronia.alcanceDe("2a02:908::1"));
        // Una IP de este mismo equipo solo llega escrita en una cabecera de reenvío: no se le cree.
        assertEquals(ServicioSincronia.INTERNET, ServicioSincronia.alcanceDe("127.0.0.1"));
        assertEquals(ServicioSincronia.INTERNET, ServicioSincronia.alcanceDe("::1"));
        assertEquals(ServicioSincronia.INTERNET, ServicioSincronia.alcanceDe("169.254.1.1"));
        // Un nombre (no una IP) nunca se busca en el DNS: lo más estricto.
        assertEquals(ServicioSincronia.INTERNET, ServicioSincronia.alcanceDe("localhost"));
        assertEquals(ServicioSincronia.INTERNET, ServicioSincronia.alcanceDe("no es una ip"));
    }

    @Test
    void cadaAlcanceSoloLlegaHastaDondeLlegaQuienLoConfiguro() throws Exception {
        String i = ServicioSincronia.INTERNET, r = ServicioSincronia.RED, e = ServicioSincronia.EQUIPO;
        assertTrue(ServicioSincronia.permitida(ip("8.8.8.8"), i));
        assertFalse(ServicioSincronia.permitida(ip("192.168.0.1"), i));
        assertFalse(ServicioSincronia.permitida(ip("100.100.100.100"), i));
        assertFalse(ServicioSincronia.permitida(ip("127.0.0.1"), i));
        assertFalse(ServicioSincronia.permitida(ip("169.254.169.254"), i));

        assertTrue(ServicioSincronia.permitida(ip("8.8.8.8"), r));
        assertTrue(ServicioSincronia.permitida(ip("192.168.0.1"), r));
        assertFalse(ServicioSincronia.permitida(ip("127.0.0.1"), r));
        assertFalse(ServicioSincronia.permitida(ip("::ffff:127.0.0.1"), r));
        assertFalse(ServicioSincronia.permitida(ip("169.254.169.254"), r));

        assertTrue(ServicioSincronia.permitida(ip("127.0.0.1"), e));     // las pruebas: dos servidores en el mismo equipo
        assertTrue(ServicioSincronia.permitida(ip("192.168.0.1"), e));
    }
}
