// Evalúa JavaScript en una página por el protocolo de depuración (CDP), sin Playwright: sirve para el
// WebView de la app Android (en el A55 o en el emulador), al que Playwright no se conecta.
//
//   adb forward tcp:9223 localabstract:webview_devtools_remote_$(adb shell pidof mx.wack.nom)
//   node pruebas/reales/cdp.mjs http://localhost:9223 "location.href"
//   node pruebas/reales/cdp.mjs http://localhost:9223 "await NOM.api.get('/api/red/info')"
//
// Imprime el resultado como JSON (las promesas se esperan). Código 1 si la página lanzó un error.
const [, , base, expr, msArg] = process.argv;
if (!base || !expr) { console.error("uso: node cdp.mjs http://localhost:9223 '<expresión>' [ms]"); process.exit(2); }
const paginas = await (await fetch(base.replace(/\/$/, "") + "/json")).json();
const p = paginas.find((x) => x.type === "page" && x.webSocketDebuggerUrl) || paginas[0];
if (!p) { console.error("no hay páginas"); process.exit(2); }
const ws = new WebSocket(p.webSocketDebuggerUrl);
await new Promise((ok, mal) => { ws.onopen = ok; ws.onerror = mal; });
let n = 0;
const pendientes = new Map();
ws.onmessage = (m) => { const d = JSON.parse(m.data); if (d.id && pendientes.has(d.id)) { pendientes.get(d.id)(d); pendientes.delete(d.id); } };
const enviar = (method, params) => new Promise((ok) => { const id = ++n; pendientes.set(id, ok); ws.send(JSON.stringify({ id, method, params })); });
const t = setTimeout(() => { console.error("tiempo agotado"); process.exit(3); }, +msArg || 60000);
const r = await enviar("Runtime.evaluate", { expression: `(async () => { return (${expr}); })()`, awaitPromise: true, returnByValue: true, userGesture: true });
clearTimeout(t);
ws.close();
if (r.result && r.result.exceptionDetails) {
  console.log(JSON.stringify({ error: (r.result.exceptionDetails.exception && r.result.exceptionDetails.exception.description) || r.result.exceptionDetails.text }));
  process.exit(1);
}
console.log(JSON.stringify(r.result && r.result.result ? r.result.result.value : r));
process.exit(0);
