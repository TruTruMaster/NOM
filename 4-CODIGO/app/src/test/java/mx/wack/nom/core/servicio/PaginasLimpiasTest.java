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

    @org.junit.jupiter.api.Test
    void laPortadaPorHttpVaSinLaPoliticaEnMeta() throws Exception {
        String html = new String(getClass().getResourceAsStream("/static/index.html").readAllBytes(), java.nio.charset.StandardCharsets.UTF_8);
        org.junit.jupiter.api.Assertions.assertTrue(html.contains("http-equiv=\"Content-Security-Policy\""));   // (la app la usa)
        String sin = PaginasLimpias.sinPoliticaEnMeta(html);
        org.junit.jupiter.api.Assertions.assertFalse(sin.contains("Content-Security-Policy"));
        org.junit.jupiter.api.Assertions.assertTrue(sin.contains("<script src=\"js/nucleo.js\"></script>") && sin.contains("<meta charset"));
        org.junit.jupiter.api.Assertions.assertEquals("<p>x</p>", PaginasLimpias.sinPoliticaEnMeta("<p>x</p>"));
    }

    @Test
    void laPortadaPideLosArchivosConSuVersionParaQueNingunaCacheDeUnoViejo() {
        String html = "<link rel=\"stylesheet\" href=\"css/nom.css\" data-nom>\n<link rel=\"icon\" href=\"img/icono.svg\">\n"
                + "<script src=\"js/nucleo.js\"></script>\n<script src=\"js/vendor/qrcode.js\"></script>\n<script src=\"https://x.y/z.js\"></script>";
        String con = PaginasLimpias.conVersion(html, "0.23.10");
        assertTrue(con.contains("href=\"css/nom.css?v=0.23.10\""), con);
        assertTrue(con.contains("src=\"js/nucleo.js?v=0.23.10\""), con);
        assertTrue(con.contains("src=\"js/vendor/qrcode.js?v=0.23.10\""), con);
        assertTrue(con.contains("href=\"img/icono.svg\""), con);          // los iconos no cambian con la versión
        assertTrue(con.contains("src=\"https://x.y/z.js\""), con);        // lo de fuera, tal cual
        assertEquals(con, PaginasLimpias.conVersion(con, "0.23.10"));      // dos veces no agrega otro ?v=
        assertEquals(html, PaginasLimpias.conVersion(html, "1.0\"><script>"));   // una versión rara no se escribe
    }
}
