package mx.wack.nom.core.servicio;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.net.Inet6Address;
import java.net.InetAddress;
import java.net.InterfaceAddress;
import java.net.NetworkInterface;
import java.util.Collections;

import org.junit.jupiter.api.Test;

/** Las direcciones globales de la red donde está el servidor (y las suyas) no cuentan como Internet. */
class FuentesRedPropiaTest {

    @Test
    void laRedPropiaNoEsInternet() throws Exception {
        assertTrue(ServicioFuentes.esPublica(InetAddress.getByName("2606:4700:3037::ac43:91ef")));
        for (NetworkInterface ni : Collections.list(NetworkInterface.getNetworkInterfaces())) {
            for (InterfaceAddress ia : ni.getInterfaceAddresses()) {
                InetAddress a = ia.getAddress();
                if (!(a instanceof Inet6Address) || a.isLinkLocalAddress() || a.isLoopbackAddress()) continue;
                byte[] b = a.getAddress();
                assertFalse(ServicioFuentes.esPublica(a), a.toString());
                b[15] ^= 0x5a;   // otro equipo de la misma red /64
                b[9] ^= 0x33;
                assertFalse(ServicioFuentes.esPublica(InetAddress.getByAddress(b)), "vecino de " + a);
                assertTrue(ServicioSincronia.distancia(InetAddress.getByAddress(b)) == 1 || (b[0] & 0xfe) == 0xfc);
            }
        }
    }
}
