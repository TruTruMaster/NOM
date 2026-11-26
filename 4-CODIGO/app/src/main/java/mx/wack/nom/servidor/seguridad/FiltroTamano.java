package mx.wack.nom.servidor.seguridad;

import java.io.IOException;

import org.springframework.core.Ordered;
import org.springframework.core.annotation.Order;
import org.springframework.stereotype.Component;
import org.springframework.web.filter.OncePerRequestFilter;

import jakarta.servlet.FilterChain;
import jakarta.servlet.ReadListener;
import jakarta.servlet.ServletException;
import jakarta.servlet.ServletInputStream;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletRequestWrapper;
import jakarta.servlet.http.HttpServletResponse;

/**
 * Tope de tamaño de lo que llega a la API: lo más grande que manda la página es un documento importado (el documento
 * tiene un tope de 25 MB). Un JSON ocupa en memoria varias veces lo que mide, así que los envíos grandes se atienden
 * por turnos, según la memoria del servidor (con 512 MB, uno de 25 MB a la vez; con 1.5 GB, unos tres). El almacén de
 * binarios revisa su propio tope (BinariosApi). Lo que llega sin decir su tamaño se cuenta mientras se lee.
 */
@Component
@Order(Ordered.HIGHEST_PRECEDENCE + 3)
public class FiltroTamano extends OncePerRequestFilter {

    /** Lo que se puede estar recibiendo a la vez: una vigésima parte de la memoria de Java. */
    static final long PRESUPUESTO = Math.max(8L << 20, Runtime.getRuntime().maxMemory() / 20);
    /** El tope de un envío (en MB con NOM_MAX_CUERPO_MB; nunca más que el presupuesto). */
    static final long MAX = Math.min(PRESUPUESTO, Math.max(1, Long.parseLong(System.getenv().getOrDefault("NOM_MAX_CUERPO_MB", "32"))) << 20);
    /** Tope para comprobar un recibo (sin sesión). */
    static final long MAX_RECIBO = Math.min(MAX, 2L << 20);
    /** Solo la subida de un binario lleva su propio tope y turno en BinariosApi. */
    private static final java.util.regex.Pattern SUBIR_BINARIO = java.util.regex.Pattern.compile("^/api/espacios/[^/]+/binarios/[0-9a-f]{64}$");
    /** Lo chico no espera turno. */
    private static final long CHICO = 256 * 1024;
    private static final java.util.concurrent.Semaphore TURNOS = new java.util.concurrent.Semaphore((int) (PRESUPUESTO >> 10), true);

    @Override
    protected void doFilterInternal(HttpServletRequest req, HttpServletResponse res, FilterChain cadena) throws ServletException, IOException {
        String ruta = req.getRequestURI();
        if (ruta == null || !ruta.startsWith("/api/") || ("PUT".equals(req.getMethod()) && SUBIR_BINARIO.matcher(ruta).matches())) {
            cadena.doFilter(req, res);
            return;
        }
        long limite = "/api/recibos/verificar".equals(ruta) ? MAX_RECIBO : MAX;
        long largo = req.getContentLengthLong();
        if (largo > limite) { responder(res, 413, "Lo que se mandó pasa de " + (limite >> 20) + " MB"); return; }
        boolean conCuerpo = largo > CHICO || (largo < 0 && (req.getHeader("Transfer-Encoding") != null));
        if (!conCuerpo) { cadena.doFilter(largo >= 0 ? req : new Contado(req, limite), res); return; }
        int kb = (int) (((largo < 0 ? limite : largo) + 1023) >> 10);
        boolean turno;
        try { turno = TURNOS.tryAcquire(kb, 60, java.util.concurrent.TimeUnit.SECONDS); }
        catch (InterruptedException e) { Thread.currentThread().interrupt(); turno = false; }
        if (!turno) { responder(res, 503, "El servidor está recibiendo mucho a la vez: intenta de nuevo en un momento"); return; }
        try { cadena.doFilter(largo >= 0 ? req : new Contado(req, limite), res); }
        finally { TURNOS.release(kb); }
    }

    private static void responder(HttpServletResponse res, int codigo, String mensaje) throws IOException {
        res.setStatus(codigo);
        res.setContentType("application/json;charset=UTF-8");
        res.getOutputStream().write(("{\"error\":\"" + mensaje + "\"}").getBytes(java.nio.charset.StandardCharsets.UTF_8));
    }

    /** La petición sin tamaño declarado: se corta al pasar del tope. */
    private static final class Contado extends HttpServletRequestWrapper {
        private final long limite;
        private ServletInputStream entrada;

        Contado(HttpServletRequest req, long limite) { super(req); this.limite = limite; }

        @Override
        public ServletInputStream getInputStream() throws IOException {
            if (entrada == null) {
                ServletInputStream base = super.getInputStream();
                entrada = new ServletInputStream() {
                    private long leidos;

                    private void contar(long n) throws IOException {
                        if (n > 0 && (leidos += n) > limite) throw new IOException("Lo que se mandó es demasiado grande");
                    }

                    @Override public int read() throws IOException { int c = base.read(); if (c >= 0) contar(1); return c; }
                    @Override public int read(byte[] b, int o, int l) throws IOException { int n = base.read(b, o, l); contar(n); return n; }
                    @Override public boolean isFinished() { return base.isFinished(); }
                    @Override public boolean isReady() { return base.isReady(); }
                    @Override public void setReadListener(ReadListener l) { base.setReadListener(l); }
                };
            }
            return entrada;
        }
    }
}
