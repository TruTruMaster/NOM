package mx.wack.nom.core.servicio;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

import tools.jackson.databind.json.JsonMapper;

/** Los datos para citar una página: los mismos de siempre y, con cualquier página, en poco tiempo. */
class FuentesPaginaTest {

    private static final JsonMapper JSON = JsonMapper.builder().build();
    private static final ServicioFuentes F = new ServicioFuentes(null, null, new CodecJson() {
        @Override public byte[] escribir(Object valor) { return JSON.writeValueAsBytes(valor); }
        @Override public Object leer(String texto) { return JSON.readValue(texto, Object.class); }
    }, null);

    private static String repetir(String s, int largo) {
        StringBuilder b = new StringBuilder(largo + s.length());
        while (b.length() < largo) b.append(s);
        return b.substring(0, largo);
    }

    @Test
    void unaPaginaNormalDaSusDatos() {
        String html = "<!doctype html><html><head><META charset=\"utf-8\"><Title>Título &amp; más</title>"
                + "<meta name=\"citation_title\" content=\"El artículo\"><meta property='og:site_name' content='Revista X'>"
                + "<meta name=author content=\"García Márquez, Gabriel\"><meta name=\"citation_publication_date\" content=\"2021/03/04\">"
                + "<metadata>no es una etiqueta meta</metadata>"
                + "<script type=\"application/ld+json\">{\"@type\":\"NewsArticle\",\"headline\":\"Otro\",\"author\":[{\"name\":\"Ana Pérez\"}]}</script>"
                + "</head><body>…</body></html>";
        Map<String, Object> f = F.datosDePagina(html, "https://www.ejemplo.mx/nota");
        assertEquals("El artículo", f.get("titulo"));
        assertEquals("Revista X", f.get("sitio"));
        assertEquals("2021-03-04", f.get("fecha"));
        assertEquals("2021", f.get("anio"));
        assertEquals("web", f.get("tipo"));
        @SuppressWarnings("unchecked") List<Map<String, Object>> as = (List<Map<String, Object>>) f.get("autores");
        assertEquals("Ana Pérez", as.get(0).get("nombres") + " " + as.get(0).get("familia"));
        // sin las etiquetas de cita, el título sale de <title>
        assertEquals("Título & más", F.datosDePagina("<html><head><title>Título &amp; más</title></head></html>", "https://a.mx").get("titulo"));
        assertEquals("Hola", F.datosDePagina("<title x='1'>Hola</TITLE>", "https://a.mx").get("titulo"));
    }

    @Test
    void cualquierPaginaSeAnalizaEnPocoTiempo() {
        int n = 400_000;
        String[] paginas = {
            "<meta " + repetir("a", n - 7) + ">",
            repetir("<title>", n),
            repetir("<meta x", n),
            repetir("<script ", n),
            repetir("<script type='application/ld+json'>", n),
            "<meta content=\"" + repetir("a b ", n) + "\">",
            "<meta " + repetir("a=\"x ", n) + ">",
        };
        for (String p : paginas) {
            long t0 = System.nanoTime();
            F.datosDePagina(p, "https://a.mx");
            long ms = (System.nanoTime() - t0) / 1_000_000;
            assertTrue(ms < ServicioFuentes.MAX_MS_ANALISIS + 1_000, "tardó " + ms + " ms con " + p.substring(0, 20));
        }
    }
}
