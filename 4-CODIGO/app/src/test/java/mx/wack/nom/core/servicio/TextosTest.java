package mx.wack.nom.core.servicio;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;

import org.junit.jupiter.api.Test;

/** Lo que escribe la gente se guarda sin caracteres de control ni marcas que cambian la dirección del texto. */
class TextosTest {

    @Test
    void unMensajeDelChatConservaSusRenglonesYPierdeLasMarcasDeDireccion() {
        assertEquals("hola\nequipo\tlisto", Textos.limpioConRenglones("hola\nequipo\tlisto\u0007"));
        assertEquals("archivo exe.pdf", Textos.limpioConRenglones("archivo ‮exe.pdf"));     // forzar derecha a izquierda
        assertEquals("a b", Textos.limpioConRenglones("a⁦ b⁩"));                       // aislar
        String igual = "Nos vemos mañana 😀";
        assertSame(igual, Textos.limpioConRenglones(igual));
        assertEquals(null, Textos.limpioConRenglones(null));
    }
}
