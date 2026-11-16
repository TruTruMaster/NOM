package mx.wack.nom.core.servicio;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

/** Las paginas de ayuda y legal se muestran sin index.html y sin diagonal final. */
class PaginasLimpiasTest {

    @Test
    void reconoceLasDireccionesLimpiasYLasViejas() {
        assertTrue(PaginasLimpias.esLimpia("legal/terminos"));
        assertTrue(PaginasLimpias.esLimpia("ayuda"));
        assertTrue(PaginasLimpias.esLimpia("ayuda/empezar/que-es"));
        assertFalse(PaginasLimpias.esLimpia("legal/terminos/"));
        assertFalse(PaginasLimpias.esLimpia("ayuda/ayuda.css"));
        assertFalse(PaginasLimpias.esLimpia("ayuda/a/b/c"));
        assertFalse(PaginasLimpias.esLimpia("legal/../api"));
        assertEquals("legal/terminos", PaginasLimpias.limpiaDe("legal/terminos/index.html"));
        assertEquals("legal/terminos", PaginasLimpias.limpiaDe("legal/terminos/"));
        assertEquals("ayuda", PaginasLimpias.limpiaDe("ayuda/index.html"));
        assertNull(PaginasLimpias.limpiaDe("legal/terminos"));
        assertNull(PaginasLimpias.limpiaDe("ayuda/ayuda.js"));
    }

    @Test
    void cambiaLosEnlacesRelativosPorDireccionesCompletas() {
        assertEquals("/legal/privacidad", PaginasLimpias.absoluta("../privacidad/index.html", "legal/terminos"));
        assertEquals("/ayuda/espacios/ramas#fusionar", PaginasLimpias.absoluta("../../espacios/ramas/index.html#fusionar", "ayuda/editores/texto"));
        assertEquals("/", PaginasLimpias.absoluta("../../index.html", "legal/terminos"));
        assertEquals("/ayuda/sin-internet", PaginasLimpias.absoluta("index.html", "ayuda/sin-internet"));
        assertEquals("/ayuda/ayuda.css", PaginasLimpias.absoluta("../../ayuda/ayuda.css", "legal/terminos"));
        assertEquals("/legal/licencias/GPL-3.0.txt", PaginasLimpias.absoluta("GPL-3.0.txt", "legal/licencias"));
    }

    @Test
    void dejaIgualLoQueNoEsRelativo() {
        for (String e : new String[] { "#contenido", "https://github.com/x", "mailto:a@b.mx", "/ayuda", "data:image/png;base64,AA", "" })
            assertEquals(e, PaginasLimpias.absoluta(e, "legal/terminos"));
    }

    @Test
    void limpiaUnaPaginaCompleta() {
        String html = "<link rel=\"stylesheet\" href=\"../../ayuda/ayuda.css\"><script src=\"../../ayuda/ayuda.js\" defer></script>"
                + "<form action=\"../../ayuda/buscar/index.html\"><a href=\"#uno\">1</a><a href=\"../index.html\">Legal</a></form>";
        assertEquals("<link rel=\"stylesheet\" href=\"/ayuda/ayuda.css\"><script src=\"/ayuda/ayuda.js\" defer></script>"
                + "<form action=\"/ayuda/buscar\"><a href=\"#uno\">1</a><a href=\"/legal\">Legal</a></form>",
                PaginasLimpias.limpiar(html, "legal/terminos"));
    }
}
