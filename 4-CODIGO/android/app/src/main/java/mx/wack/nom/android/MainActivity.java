package mx.wack.nom.android;

import android.app.Activity;
import android.content.ContentValues;
import android.content.Context;
import android.content.Intent;
import android.graphics.Color;
import android.graphics.Rect;
import android.net.Uri;
import android.net.wifi.WifiManager;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.print.PrintAttributes;
import android.print.PrintManager;
import android.provider.MediaStore;
import android.util.Base64;
import android.view.ActionMode;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.view.WindowInsets;
import android.webkit.JavascriptInterface;
import android.webkit.PermissionRequest;
import android.webkit.ValueCallback;
import android.webkit.WebBackForwardList;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceRequest;
import android.webkit.WebResourceResponse;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.FrameLayout;
import android.widget.Toast;

import mx.wack.nom.android.anfitrion.Respaldo;
import mx.wack.nom.android.anfitrion.ServicioAnfitrion;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.net.DatagramPacket;
import java.net.DatagramSocket;
import java.net.InetSocketAddress;
import java.net.SocketTimeoutException;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;

/**
 * Cliente Android de NOM (ClienteAndroid del diagrama de clases).
 *
 * La interfaz es la misma que sirve el servidor, empaquetada en assets/: así el
 * editor abre aunque no haya red (CU-26) y las operaciones se encolan en el
 * dispositivo hasta volver a ver el servidor de la LAN. Lo único nativo es lo que
 * un navegador no puede hacer: escuchar el anuncio UDP de la red local, guardar
 * archivos en Descargas y abrir el selector de archivos.
 *
 * Desde la 0.19 la interfaz tiene su propio origen (ORIGEN): la app contesta esas direcciones con sus archivos de
 * assets/ (nada sale a la red) y el WebView ya no abre archivos con file://. Lo guardado por las versiones de
 * antes (en el origen file://) se pasa al nuevo la primera vez (migrar.html y js/migrar-origen.js).
 */
public class MainActivity extends Activity {

    private static final int ELEGIR_ARCHIVO = 41;
    private static final int PERMISO_CAMARA = 42;
    private static final int PUERTO_DESCUBRIMIENTO = 47474;

    /** Origen de la interfaz desde la 0.19 (dominio reservado para esto en Android: nunca sale a Internet). */
    static final String ORIGEN = "https://appassets.androidplatform.net";
    private static final String HOST = "appassets.androidplatform.net";
    static final String INICIO = ORIGEN + "/index.html";
    private static final String PREFS = "nom", MIGRADO = "origen-propio", INTENTOS = "origen-propio-intentos";
    /** Paso al origen propio: 0 = no se está haciendo, 1 = leyendo en file://, 2 = escribiendo en ORIGEN. */
    private volatile int fase;
    private BufferedWriter escritura;
    private BufferedReader lectura;
    /** Al llegar a la interfaz después de pasar los datos, Atrás no debe regresar a la página de paso. */
    private volatile boolean limpiarHistoria;

    private WebView web;
    /** El WebView ya se destruyó (su página se cerró sola y se va a abrir otro): no se destruye dos veces. */
    private boolean webDestruida;
    /** Cuándo se cerraron solas las últimas páginas (para no abrirla otra vez sin fin si se cierra al cargar). */
    private static final java.util.ArrayDeque<Long> paginasCaidas = new java.util.ArrayDeque<>();
    private ValueCallback<Uri[]> alElegir;
    /** La cámara que pidió la página (escanear un QR) mientras Android pregunta el permiso. */
    private PermissionRequest pedidoCamara;

    @Override
    protected void onCreate(Bundle guardado) {
        super.onCreate(guardado);
        web = new Hoja(this);
        web.setBackgroundColor(Color.parseColor("#F4F5F7"));
        // El WebView no respeta su propio padding: el margen de las barras va en un contenedor.
        FrameLayout raiz = new FrameLayout(this);
        raiz.setBackgroundColor(Color.parseColor("#ECEDF1"));
        raiz.addView(web, new FrameLayout.LayoutParams(FrameLayout.LayoutParams.MATCH_PARENT, FrameLayout.LayoutParams.MATCH_PARENT));
        setContentView(raiz);
        respetarBarrasDelSistema(raiz);

        WebSettings s = web.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);            // cola sin conexión y copia local de documentos
        // La red del salón (otro teléfono o una computadora que comparte) va por http y la interfaz es https.
        s.setMixedContentMode(WebSettings.MIXED_CONTENT_ALWAYS_ALLOW);
        s.setMediaPlaybackRequiresUserGesture(true);
        s.setTextZoom(100);                      // el tamaño de letra del sistema no deforma la interfaz
        // La interfaz tiene sus propios temas claro y oscuro: el sistema (modo oscuro forzado de
        // Android o de la marca del teléfono) no la debe repintar.
        if (Build.VERSION.SDK_INT >= 33) s.setAlgorithmicDarkeningAllowed(false);
        else if (Build.VERSION.SDK_INT >= 29) s.setForceDark(WebSettings.FORCE_DARK_OFF);
        if (Build.VERSION.SDK_INT >= 29) web.setForceDarkAllowed(false);

        web.setWebViewClient(new WebViewClient() {
            @Override
            public WebResourceResponse shouldInterceptRequest(WebView v, WebResourceRequest pedido) {
                return deLaApp(pedido.getUrl());
            }

            @Override
            public boolean shouldOverrideUrlLoading(WebView v, WebResourceRequest pedido) {
                return afuera(pedido.getUrl().toString());
            }

            @Override
            public boolean shouldOverrideUrlLoading(WebView v, String url) {
                return afuera(url);
            }

            @Override
            public void onPageFinished(WebView v, String url) {
                if (limpiarHistoria && url != null && url.startsWith(INICIO)) {
                    limpiarHistoria = false;
                    v.clearHistory();
                }
            }

            /**
             * La página se cerró sola (se quedó sin memoria, p. ej. con un archivo enorme, o Android la cerró para
             * liberar memoria). Sin esto Android cierra toda la app. Se quita el WebView y la actividad se abre de nuevo;
             * lo que faltaba enviar sigue guardado en el equipo. Si se cae tres veces en un minuto, se cierra la ventana.
             */
            @Override
            public boolean onRenderProcessGone(WebView v, android.webkit.RenderProcessGoneDetail detalle) {
                long ahora = android.os.SystemClock.elapsedRealtime();
                synchronized (paginasCaidas) {
                    while (!paginasCaidas.isEmpty() && ahora - paginasCaidas.peekFirst() > 60_000) paginasCaidas.pollFirst();
                    paginasCaidas.addLast(ahora);
                }
                if (v.getParent() instanceof android.view.ViewGroup padre) padre.removeView(v);
                if (v == web && !webDestruida) { webDestruida = true; v.destroy(); }
                else if (v != web) v.destroy();
                if (paginasCaidas.size() >= 3) finish(); else recreate();
                return true;
            }
        });
        web.setWebChromeClient(new WebChromeClient() {
            /** 2.1 Escanear un código QR: la página pide la cámara; primero se pide el permiso a Android. */
            @Override
            public void onPermissionRequest(PermissionRequest pedido) {
                runOnUiThread(() -> {
                    boolean soloCamara = true;
                    for (String r : pedido.getResources()) if (!PermissionRequest.RESOURCE_VIDEO_CAPTURE.equals(r)) soloCamara = false;
                    if (!soloCamara) { pedido.deny(); return; }
                    if (Build.VERSION.SDK_INT < 23 || checkSelfPermission(android.Manifest.permission.CAMERA) == android.content.pm.PackageManager.PERMISSION_GRANTED) {
                        pedido.grant(pedido.getResources());
                        return;
                    }
                    if (pedidoCamara != null) pedidoCamara.deny();
                    pedidoCamara = pedido;
                    requestPermissions(new String[] { android.Manifest.permission.CAMERA }, PERMISO_CAMARA);
                });
            }

            @Override
            public boolean onShowFileChooser(WebView v, ValueCallback<Uri[]> cb, FileChooserParams p) {
                if (alElegir != null) alElegir.onReceiveValue(null);
                alElegir = cb;
                try {
                    startActivityForResult(p.createIntent(), ELEGIR_ARCHIVO);
                } catch (Exception e) {
                    alElegir = null;
                    return false;
                }
                return true;
            }
        });
        web.addJavascriptInterface(new Puente(), "NomAndroid");

        // Desde Android 16 (targetSdk 36+) el sistema ya no llama a onBackPressed(): sin esto, Atrás
        // sacaba de la app aunque hubiera una ventana abierta.
        if (Build.VERSION.SDK_INT >= 33) {
            getOnBackInvokedDispatcher().registerOnBackInvokedCallback(
                android.window.OnBackInvokedDispatcher.PRIORITY_DEFAULT, () -> atras(true));
        }

        boolean migrado = getSharedPreferences(PREFS, MODE_PRIVATE).getBoolean(MIGRADO, false);
        if (guardado == null) tomarEnlace(getIntent(), false);
        if (!migrado) { pasarAlOrigenPropio(); return; }
        ajustesDelOrigenPropio();
        WebBackForwardList h = guardado == null ? null : web.restoreState(guardado);
        String actual = h == null || h.getCurrentItem() == null ? null : h.getCurrentItem().getUrl();
        if (actual == null || !actual.startsWith(ORIGEN + "/")) web.loadUrl(INICIO);
    }

    // ------------------------------------------------------------ origen propio (0.19)

    /** Lo de siempre: sin acceso a archivos con file:// (la interfaz se sirve desde ORIGEN). */
    private void ajustesDelOrigenPropio() {
        WebSettings s = web.getSettings();
        s.setAllowFileAccess(false);
        s.setAllowFileAccessFromFileURLs(false);
        s.setAllowUniversalAccessFromFileURLs(false);
    }

    /** Solo para leer lo que guardaron las versiones de antes: la página de paso corre en file:// como ellas. */
    private void pasarAlOrigenPropio() {
        WebSettings s = web.getSettings();
        s.setAllowFileAccess(true);
        s.setAllowFileAccessFromFileURLs(true);
        s.setAllowUniversalAccessFromFileURLs(true);
        fase = 1;
        web.loadUrl("file:///android_asset/migrar.html");
    }

    /** Termina el paso (bien o mal) y abre la interfaz. Si falló, lo de antes sigue en su lugar y se reintenta
     *  al abrir la app, hasta 3 veces; después se sigue sin ello (sigue guardado en el origen de antes). */
    private void terminarPaso(boolean ok, String info) {
        fase = 0;
        cerrarPaso();
        android.content.SharedPreferences p = getSharedPreferences(PREFS, MODE_PRIVATE);
        int intentos = p.getInt(INTENTOS, 0) + 1;
        boolean listo = ok || intentos >= 3;
        p.edit().putInt(INTENTOS, intentos).putBoolean(MIGRADO, listo).putString(MIGRADO + "-info", (ok ? "bien: " : "mal: ") + info).apply();
        if (listo) archivoPaso().delete();
        android.util.Log.i("NOM", "Paso al origen propio (" + intentos + "): " + (ok ? "bien, " : "mal, ") + info);
        runOnUiThread(() -> {
            ajustesDelOrigenPropio();
            limpiarHistoria = true;
            web.loadUrl(INICIO);
        });
    }

    private File archivoPaso() { return new File(getFilesDir(), "paso-origen.jsonl"); }

    private void cerrarPaso() {
        try { if (escritura != null) escritura.close(); } catch (IOException ignorado) { }
        try { if (lectura != null) lectura.close(); } catch (IOException ignorado) { }
        escritura = null;
        lectura = null;
    }

    /** Enlaces: lo de la interfaz se queda en la app; lo demás, al navegador. */
    private boolean afuera(String url) {
        if (url.startsWith(ORIGEN + "/") || url.equals(ORIGEN) || (fase != 0 && url.startsWith("file:///android_asset/"))) return false;
        try { startActivity(new Intent(Intent.ACTION_VIEW, Uri.parse(url))); } catch (Exception ignorado) { }
        return true;
    }

    private static final Map<String, String> TIPOS = new HashMap<>();
    static {
        String[][] t = {
            { "html", "text/html" }, { "js", "text/javascript" }, { "mjs", "text/javascript" }, { "css", "text/css" },
            { "json", "application/json" }, { "webmanifest", "application/manifest+json" }, { "txt", "text/plain" },
            { "xml", "application/xml" }, { "svg", "image/svg+xml" }, { "png", "image/png" }, { "jpg", "image/jpeg" },
            { "jpeg", "image/jpeg" }, { "gif", "image/gif" }, { "webp", "image/webp" }, { "avif", "image/avif" },
            { "ico", "image/x-icon" }, { "woff2", "font/woff2" }, { "woff", "font/woff" }, { "ttf", "font/ttf" },
            { "otf", "font/otf" }, { "wasm", "application/wasm" }, { "pdf", "application/pdf" }, { "mp3", "audio/mpeg" },
            { "ogg", "audio/ogg" }, { "opus", "audio/ogg" }, { "wav", "audio/wav" }, { "mp4", "video/mp4" },
            { "webm", "video/webm" }, { "zip", "application/zip" },
            { "docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
            { "xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
            { "pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
        };
        for (String[] x : t) TIPOS.put(x[0], x[1]);
    }
    private static final java.util.Set<String> TEXTO = new java.util.HashSet<>(java.util.Arrays.asList(
            "html", "js", "mjs", "css", "json", "webmanifest", "txt", "xml", "svg"));

    /** Las direcciones de ORIGEN se contestan con los archivos de assets/ (una carpeta, con su index.html). */
    private WebResourceResponse deLaApp(Uri u) {
        if (u == null || !"https".equals(u.getScheme()) || !HOST.equals(u.getHost()) || u.getPort() != -1) return null;
        Map<String, String> cab = new HashMap<>();
        cab.put("Cache-Control", "no-cache");
        cab.put("X-Content-Type-Options", "nosniff");
        String ruta = u.getPath() == null ? "" : u.getPath();
        while (ruta.startsWith("/")) ruta = ruta.substring(1);
        if (ruta.isEmpty() || ruta.endsWith("/")) ruta += "index.html";
        if (ruta.contains("..") || ruta.contains("\\") || ruta.indexOf('\0') >= 0) return noEsta(cab);
        int punto = ruta.lastIndexOf('.'), barra = ruta.lastIndexOf('/');
        String ext = punto > barra ? ruta.substring(punto + 1).toLowerCase(Locale.ROOT) : "";
        InputStream in;
        try {
            in = getAssets().open(ruta);
        } catch (IOException e) {
            if (!ext.isEmpty()) return noEsta(cab);
            try { in = getAssets().open(ruta + "/index.html"); ext = "html"; } catch (IOException e2) { return noEsta(cab); }
        }
        String tipo = TIPOS.containsKey(ext) ? TIPOS.get(ext) : "application/octet-stream";
        return new WebResourceResponse(tipo, TEXTO.contains(ext) ? "utf-8" : null, 200, "OK", cab, in);
    }

    private static WebResourceResponse noEsta(Map<String, String> cab) {
        return new WebResourceResponse("text/plain", "utf-8", 404, "Not Found", cab, new ByteArrayInputStream(new byte[0]));
    }

    /** DEF-49: el enlace de invitación con el que se abrió la app (la página lo pide al arrancar). */
    private volatile String enlacePendiente;

    /** La app ya estaba abierta y se tocó un enlace de invitación: se le pasa a la página. */
    @Override
    protected void onNewIntent(Intent intento) {
        super.onNewIntent(intento);
        setIntent(intento);
        tomarEnlace(intento, true);
    }

    private void tomarEnlace(Intent intento, boolean abierta) {
        Uri u = intento != null && Intent.ACTION_VIEW.equals(intento.getAction()) ? intento.getData() : null;
        if (u == null) return;
        String s = u.toString();
        if (!s.startsWith("https://") || s.length() > 2000) return;
        // Solo los enlaces que declara el manifiesto.
        String host = u.getHost() == null ? "" : u.getHost().toLowerCase(java.util.Locale.ROOT), camino = u.getPath() == null ? "" : u.getPath();
        if (!(host.equals("www.nom.qd.je") || host.equals("nom.qd.je") || host.equals("nomcolab.dpdns.org")) || !(camino.equals("/unirse") || camino.equals("/proyectar"))) return;
        if (abierta) web.evaluateJavascript("window.NOM && NOM.union && NOM.union.usar(" + org.json.JSONObject.quote(s) + ")", null);
        else enlacePendiente = s;
    }

    /** Desde Android 15 la app dibuja bajo las barras; se deja el margen que ocupan. */
    private void respetarBarrasDelSistema(View raiz) {
        raiz.setOnApplyWindowInsetsListener((v, insets) -> {
            int arriba, abajo, izq, der;
            if (Build.VERSION.SDK_INT >= 30) {
                // (también la muesca o la cámara en la pantalla: con el teléfono acostado quedan a un lado)
                android.graphics.Insets i = insets.getInsets(WindowInsets.Type.systemBars() | WindowInsets.Type.ime() | WindowInsets.Type.displayCutout());
                arriba = i.top; abajo = i.bottom; izq = i.left; der = i.right;
            } else {
                arriba = insets.getSystemWindowInsetTop(); abajo = insets.getSystemWindowInsetBottom();
                izq = insets.getSystemWindowInsetLeft(); der = insets.getSystemWindowInsetRight();
                if (Build.VERSION.SDK_INT >= 28 && insets.getDisplayCutout() != null) {
                    android.view.DisplayCutout m = insets.getDisplayCutout();
                    arriba = Math.max(arriba, m.getSafeInsetTop()); abajo = Math.max(abajo, m.getSafeInsetBottom());
                    izq = Math.max(izq, m.getSafeInsetLeft()); der = Math.max(der, m.getSafeInsetRight());
                }
            }
            v.setPadding(izq, arriba, der, abajo);
            return insets;
        });
        if (Build.VERSION.SDK_INT >= 23) {
            getWindow().getDecorView().setSystemUiVisibility(View.SYSTEM_UI_FLAG_LIGHT_STATUS_BAR);
        }
    }

    @Override
    protected void onSaveInstanceState(Bundle estado) {
        super.onSaveInstanceState(estado);
        web.saveState(estado);
    }

    @Override
    public void onBackPressed() {
        atras(false);
    }

    /** Botón o gesto Atrás: primero se cierra la ventana o el menú que esté abierto; si no hay, se
     *  regresa de pantalla; y solo en la primera pantalla se sale (la app queda en segundo plano). */
    private void atras(boolean nuevo) {
        if (fase != 0) return;   // mientras se pasan los datos al origen propio (unos segundos), Atrás no hace nada
        // Primero se cierra lo que esté encima: un menú, la ruedita o una ventana (antes, con un menú
        // abierto, Atrás lo cerraba y además salía del documento).
        web.evaluateJavascript("(window.NOM && NOM.atras) ? (NOM.atras() ? 'cerrada' : 'no') : ((window.NOM && NOM.hayHoja && NOM.hayHoja()) ? (NOM.cerrarHoja(), 'cerrada') : 'no')", r -> {
            if (r != null && r.contains("cerrada")) return;
            if (web.canGoBack()) web.goBack();
            else if (nuevo) moveTaskToBack(true);
            else MainActivity.super.onBackPressed();
        });
    }

    @Override
    protected void onActivityResult(int pedido, int resultado, Intent datos) {
        super.onActivityResult(pedido, resultado, datos);
        if (pedido != ELEGIR_ARCHIVO || alElegir == null) return;
        alElegir.onReceiveValue(WebChromeClient.FileChooserParams.parseResult(resultado, datos));
        alElegir = null;
    }

    @Override
    public void onRequestPermissionsResult(int pedido, String[] permisos, int[] resultados) {
        super.onRequestPermissionsResult(pedido, permisos, resultados);
        if (pedido != PERMISO_CAMARA || pedidoCamara == null) return;
        boolean ok = resultados.length > 0 && resultados[0] == android.content.pm.PackageManager.PERMISSION_GRANTED;
        if (ok) pedidoCamara.grant(pedidoCamara.getResources()); else pedidoCamara.deny();
        pedidoCamara = null;
    }

    @Override
    protected void onDestroy() {
        if (!webDestruida) { webDestruida = true; web.destroy(); }
        super.onDestroy();
    }

    /**
     * El nombre con que se guarda un archivo: viene de la página (el nombre de un documento, que pone cualquiera que lo
     * edite), así que no puede llevar carpetas («../»), caracteres de control ni empezar con punto.
     */
    static String nombreSeguro(String nombre) {
        String n = nombre == null ? "" : nombre.replaceAll("[\\p{Cntrl}/\\\\:*?\"<>|]", "_").replaceAll("^[.\\s]+", "").trim();
        if (n.length() > 150) n = n.substring(n.length() - 150);
        return n.isEmpty() ? "NOM" : n;
    }

    /** Guarda un archivo en Descargas (Android 10+ con MediaStore; antes, en la carpeta de la app). */
    private void aDescargas(String nombre, String mime, InputStream datos) throws Exception {
        nombre = nombreSeguro(nombre);
        if (mime == null || !mime.matches("[a-z]+/[A-Za-z0-9.+_-]{1,100}")) mime = "application/octet-stream";
        try (InputStream in = datos) {
            OutputStream salida;
            if (Build.VERSION.SDK_INT >= 29) {
                ContentValues v = new ContentValues();
                v.put(MediaStore.Downloads.DISPLAY_NAME, nombre);
                v.put(MediaStore.Downloads.MIME_TYPE, mime);
                Uri uri = getContentResolver().insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, v);
                salida = getContentResolver().openOutputStream(uri);
            } else {
                File dir = getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS);
                salida = new FileOutputStream(new File(dir, nombre));
            }
            try (OutputStream o = salida) {
                byte[] buf = new byte[64 * 1024];
                int n;
                while ((n = in.read(buf)) > 0) o.write(buf, 0, n);
            }
        }
    }

    /** Lo que la interfaz web puede pedirle al teléfono (window.NomAndroid). */
    private class Puente {

        // ---- paso al origen propio (0.19): solo responden mientras MainActivity lo está haciendo

        /** 1. (en file://) empieza a recibir lo guardado. */
        @JavascriptInterface
        public boolean migracionInicio() {
            if (fase != 1) return false;
            synchronized (MainActivity.this) {
                cerrarPaso();
                try {
                    escritura = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(archivoPaso(), false), StandardCharsets.UTF_8), 1 << 16);
                    return true;
                } catch (IOException e) {
                    terminarPaso(false, "no se pudo escribir: " + e.getMessage());
                    return false;
                }
            }
        }

        /** Un pedazo de un mensaje; con fin, el mensaje termina (un renglón del archivo). */
        @JavascriptInterface
        public boolean migracionParte(String texto, boolean fin) {
            if (fase != 1) return false;
            synchronized (MainActivity.this) {
                if (escritura == null || texto == null || texto.indexOf('\n') >= 0) return false;
                try {
                    escritura.write(texto);
                    if (fin) escritura.write('\n');
                    return true;
                } catch (IOException e) {
                    return false;
                }
            }
        }

        /** 1. terminó de leer: si salió bien, se abre la misma página en el origen nuevo para escribir. */
        @JavascriptInterface
        public void migracionFin(boolean ok, String info) {
            if (fase != 1) return;
            synchronized (MainActivity.this) {
                boolean bien = ok;
                try { if (escritura != null) escritura.close(); } catch (IOException e) { bien = false; }
                escritura = null;
                if (!bien) { terminarPaso(false, "al leer: " + info); return; }
                fase = 2;
            }
            runOnUiThread(() -> {
                ajustesDelOrigenPropio();
                web.loadUrl(ORIGEN + "/migrar.html");
            });
        }

        /** 2. (en el origen nuevo) el siguiente pedazo: "1…" termina un mensaje, "0…" sigue; "" = ya no hay más. */
        @JavascriptInterface
        public String migracionSiguiente() {
            if (fase != 2) return "";
            synchronized (MainActivity.this) {
                try {
                    if (lectura == null) lectura = new BufferedReader(new InputStreamReader(new FileInputStream(archivoPaso()), StandardCharsets.UTF_8), 1 << 16);
                    StringBuilder b = new StringBuilder();
                    int c;
                    while (b.length() < (1 << 20) && (c = lectura.read()) != -1) {
                        if (c == '\n') return "1" + b;
                        b.append((char) c);
                    }
                    if (b.length() == 0) return "";
                    return "0" + b;
                } catch (IOException e) {
                    return "";
                }
            }
        }

        /** 2. terminó de escribir (bien o mal): se abre la interfaz. */
        @JavascriptInterface
        public void migracionHecha(boolean ok, String info) {
            if (fase != 2) return;
            terminarPaso(ok, (ok ? "" : "al escribir: ") + info);
        }

        /** DEF-49: el enlace de invitación con el que se abrió la app ("" si no hubo). Se entrega una sola vez. */
        @JavascriptInterface
        public String enlacePendiente() {
            String s = enlacePendiente;
            enlacePendiente = null;
            return s == null ? "" : s;
        }

        /**
         * Identificador de este teléfono para el tope de cuentas nuevas por hora del servidor. La MAC
         * no se puede leer (Android 6+ devuelve una falsa y la cambia al azar en cada red); se usa el
         * ANDROID_ID, que es propio de este teléfono y de la llave del APK y sobrevive a reinstalar.
         * Nunca sale crudo: va resumido con SHA-256.
         */
        @JavascriptInterface
        public String equipo() {
            try {
                String a = android.provider.Settings.Secure.getString(getContentResolver(), android.provider.Settings.Secure.ANDROID_ID);
                if (a == null || a.isEmpty()) return "";
                // "nom-colab-equipo" (el nombre de antes) se queda así: con otro texto cada teléfono cambiaría de identificador.
                byte[] h = java.security.MessageDigest.getInstance("SHA-256").digest(("nom-colab-equipo:" + a).getBytes(java.nio.charset.StandardCharsets.UTF_8));
                return Base64.encodeToString(h, Base64.URL_SAFE | Base64.NO_PADDING | Base64.NO_WRAP).substring(0, 32);
            } catch (Exception e) {
                return "";
            }
        }

        /** Letra más grande en la app (Ajustes → Tamaño de la letra). 100 = normal, hasta 140. */
        @JavascriptInterface
        public void tamanoLetra(int porciento) {
            final int p = Math.max(100, Math.min(140, porciento));
            runOnUiThread(() -> web.getSettings().setTextZoom(p));
        }

        /**
         * CU-15 en Android: escucha 3.5 s los anuncios UDP de los equipos que tienen
         * la "Conexión del proyecto" activa y responde a window.__nomServidores(json).
         */
        @JavascriptInterface
        public void buscarServidores() {
            new Thread(() -> {
                Map<String, JSONObject> vistos = new LinkedHashMap<>();
                WifiManager wifi = (WifiManager) getApplicationContext().getSystemService(Context.WIFI_SERVICE);
                WifiManager.MulticastLock candado = wifi == null ? null : wifi.createMulticastLock("nom-descubrimiento");
                DatagramSocket sock = null;
                try {
                    if (candado != null) { candado.setReferenceCounted(false); candado.acquire(); }
                    sock = new DatagramSocket(null);
                    sock.setReuseAddress(true);
                    sock.setBroadcast(true);
                    sock.bind(new InetSocketAddress(PUERTO_DESCUBRIMIENTO));
                    sock.setSoTimeout(500);
                    byte[] buf = new byte[16 * 1024];
                    long fin = System.currentTimeMillis() + 3500;
                    while (System.currentTimeMillis() < fin) {
                        try {
                            DatagramPacket p = new DatagramPacket(buf, buf.length);
                            sock.receive(p);
                            JSONObject a = new JSONObject(new String(p.getData(), 0, p.getLength(), StandardCharsets.UTF_8));
                            if (a.optInt("nom") != 1) continue;
                            JSONObject s = new JSONObject();
                            s.put("equipo", a.optString("equipo"));
                            s.put("host", p.getAddress().getHostAddress());
                            s.put("puerto", a.optInt("puerto", 8080));
                            s.put("espacios", a.optJSONArray("espacios") == null ? 0 : a.optJSONArray("espacios").length());
                            vistos.put(a.optString("servidor"), s);
                        } catch (SocketTimeoutException ignorado) {
                            // se sigue escuchando hasta el plazo
                        }
                    }
                } catch (Exception e) {
                    // sin Wi-Fi o puerto ocupado: se devuelve lo que se haya visto
                } finally {
                    if (sock != null) sock.close();
                    if (candado != null && candado.isHeld()) candado.release();
                }
                String json = new JSONArray(vistos.values()).toString();
                runOnUiThread(() -> web.evaluateJavascript(
                        "window.__nomServidores && window.__nomServidores(" + JSONObject.quote(json) + ")", null));
            }, "nom-buscar").start();
        }

        // ------------------------------------------------------------ el teléfono como servidor

        /** Enciende el servidor del teléfono y espera a que abra su puerto (hasta 6 s). */
        @JavascriptInterface
        public String iniciarServidor() {
            if (Build.VERSION.SDK_INT < 26)
                return "{\"activo\":false,\"error\":\"Para que el teléfono sea el servidor se necesita Android 8.0 o más reciente.\"}";
            if (Build.VERSION.SDK_INT >= 33 && checkSelfPermission(android.Manifest.permission.POST_NOTIFICATIONS) != android.content.pm.PackageManager.PERMISSION_GRANTED)
                runOnUiThread(() -> requestPermissions(new String[] { android.Manifest.permission.POST_NOTIFICATIONS }, 3));
            startForegroundService(new Intent(MainActivity.this, ServicioAnfitrion.class));
            long fin = System.currentTimeMillis() + 6000;
            String estado = ServicioAnfitrion.estadoJson();
            while (System.currentTimeMillis() < fin && !estado.contains("\"activo\":true") && !estado.contains("\"error\"")) {
                try { Thread.sleep(60); } catch (InterruptedException e) { break; }
                estado = ServicioAnfitrion.estadoJson();
            }
            return estado;
        }

        @JavascriptInterface
        public String detenerServidor() {
            startService(new Intent(MainActivity.this, ServicioAnfitrion.class).setAction(ServicioAnfitrion.DETENER));
            return "{\"activo\":false}";
        }

        @JavascriptInterface
        public String estadoServidor() { return ServicioAnfitrion.estadoJson(); }

        /**
         * Fase 13: la llave con la que el dueño de este teléfono entra a su propio servidor (su alias).
         * Vive en la carpeta privada de la app que no se respalda (no va en "Exportar todo"): otra app
         * del teléfono que hable con 127.0.0.1 no la conoce.
         */
        @JavascriptInterface
        public String llaveDueno() {
            File f = new File(Build.VERSION.SDK_INT >= 21 ? getNoBackupFilesDir() : getFilesDir(), "llave-dueno");
            try {
                if (f.isFile() && f.length() >= 32) {
                    try (FileInputStream in = new FileInputStream(f)) {
                        byte[] b = new byte[(int) f.length()];
                        int n = in.read(b);
                        return new String(b, 0, Math.max(0, n), StandardCharsets.US_ASCII).trim();
                    }
                }
                byte[] r = new byte[32];
                new java.security.SecureRandom().nextBytes(r);
                String llave = Base64.encodeToString(r, Base64.URL_SAFE | Base64.NO_PADDING | Base64.NO_WRAP);
                try (FileOutputStream o = new FileOutputStream(f)) { o.write(llave.getBytes(StandardCharsets.US_ASCII)); }
                return llave;
            } catch (Exception e) {
                return "";
            }
        }

        /** Hoja de compartir de Android (WhatsApp, correo, Bluetooth…) con un texto, p. ej. la dirección del servidor. */
        @JavascriptInterface
        public void compartirTexto(String texto) {
            runOnUiThread(() -> {
                Intent i = new Intent(Intent.ACTION_SEND).setType("text/plain").putExtra(Intent.EXTRA_TEXT, texto);
                startActivity(Intent.createChooser(i, "Compartir"));
            });
        }

        /** CU-29 en Android: los Blob de descarga no funcionan en WebView; se guarda en Descargas. */
        @JavascriptInterface
        public void guardarArchivo(String nombre, String mime, String base64) {
            try {
                aDescargas(nombre, mime, new java.io.ByteArrayInputStream(Base64.decode(base64, Base64.DEFAULT)));
                runOnUiThread(() -> Toast.makeText(MainActivity.this, "Guardado en Descargas: " + nombreSeguro(nombre), Toast.LENGTH_LONG).show());
            } catch (Exception e) {
                runOnUiThread(() -> Toast.makeText(MainActivity.this, "No se pudo guardar el archivo", Toast.LENGTH_LONG).show());
            }
        }

        /**
         * En la app no existe "imprimir" del navegador: se imprime la página con el servicio de
         * impresión de Android (que también guarda como PDF). La página decide qué se ve al imprimir.
         */
        @JavascriptInterface
        public void imprimir(String titulo) {
            runOnUiThread(() -> {
                try {
                    PrintManager pm = (PrintManager) getSystemService(Context.PRINT_SERVICE);
                    String t = titulo == null || titulo.isEmpty() ? "NOM" : titulo;
                    pm.print(t, web.createPrintDocumentAdapter(t), new PrintAttributes.Builder().setMediaSize(PrintAttributes.MediaSize.NA_LETTER).build());
                } catch (Exception e) {
                    Toast.makeText(MainActivity.this, "No se pudo abrir la impresión", Toast.LENGTH_LONG).show();
                }
            });
        }

        /** 6.4 Los respaldos diarios del servidor de este teléfono: [{nombre, bytes, fecha}]. */
        @JavascriptInterface
        public String respaldos() { return Respaldo.lista(getFilesDir()); }

        /** 6.4 "Exportar todo": un respaldo de este momento, a Descargas. */
        @JavascriptInterface
        public String exportarRespaldo() {
            try {
                File f = Respaldo.ahora(getFilesDir());
                aDescargas(f.getName(), "application/zip", new FileInputStream(f));
                runOnUiThread(() -> Toast.makeText(MainActivity.this, "Guardado en Descargas: " + f.getName(), Toast.LENGTH_LONG).show());
                return f.getName();
            } catch (Exception e) {
                runOnUiThread(() -> Toast.makeText(MainActivity.this, "No se pudo exportar: " + e.getMessage(), Toast.LENGTH_LONG).show());
                return "";
            }
        }
    }

    /**
     * WebView sin la barra nativa de "Cortar · Copiar · Seleccionar todo" al seleccionar texto: esas
     * acciones (y más) están en la ruedita, que sale al mantener presionado. Las manijas para ajustar
     * la selección se conservan.
     */
    static class Hoja extends WebView {
        Hoja(Context c) { super(c); }

        @Override
        public ActionMode startActionMode(ActionMode.Callback cb) { return super.startActionMode(new SinMenu(cb)); }

        @Override
        public ActionMode startActionMode(ActionMode.Callback cb, int tipo) { return super.startActionMode(new SinMenu(cb), tipo); }
    }

    /** Envuelve el callback del sistema y deja su menú vacío (así la barra flotante no aparece). */
    static class SinMenu extends ActionMode.Callback2 {
        private final ActionMode.Callback original;
        SinMenu(ActionMode.Callback original) { this.original = original; }

        @Override
        public boolean onCreateActionMode(ActionMode modo, Menu menu) {
            original.onCreateActionMode(modo, menu);
            menu.clear();
            return true;
        }

        @Override
        public boolean onPrepareActionMode(ActionMode modo, Menu menu) {
            original.onPrepareActionMode(modo, menu);
            menu.clear();
            return true;
        }

        @Override
        public boolean onActionItemClicked(ActionMode modo, MenuItem item) { return original.onActionItemClicked(modo, item); }

        @Override
        public void onDestroyActionMode(ActionMode modo) { original.onDestroyActionMode(modo); }

        @Override
        public void onGetContentRect(ActionMode modo, View vista, Rect rect) {
            if (original instanceof ActionMode.Callback2) ((ActionMode.Callback2) original).onGetContentRect(modo, vista, rect);
            else super.onGetContentRect(modo, vista, rect);
        }
    }
}
