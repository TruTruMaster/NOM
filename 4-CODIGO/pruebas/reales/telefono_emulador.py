"""El teléfono que comparte, de punta a punta, con la app en el emulador (o en el A55) y una computadora.

  python pruebas/reales/telefono_emulador.py http://localhost:8098 CODIGO DOC [http://localhost:9223]

- La app ya encendió su servidor y entró como dueña (alias); en el teléfono hay un espacio con el
  documento DOC y un enlace de invitación CODIGO (ver docs/RELEVO-PARIDAD-Y-PRUEBAS-REALES.md, B-3).
- La computadora SIN acceso a la nube abre /unirse?c=CODIGO en el teléfono, entra con un alias y abre el
  documento; escribe y lo que escribe le llega a la app; la app escribe (por CDP, en su WebView) y le
  llega a la computadora; luego se corta la red de la computadora un momento y al volver se junta todo.
"""
import json, subprocess, sys, time, os
from pathlib import Path
from playwright.sync_api import sync_playwright
TEL, COD, DOC = sys.argv[1].rstrip("/"), sys.argv[2], sys.argv[3]
CDP = sys.argv[4] if len(sys.argv) > 4 else "http://localhost:9223"
AQUI = Path(__file__).resolve().parent
res = []
def anota(n, ok, det=""):
    res.append((n, ok)); print(("BIEN " if ok else "MAL  ") + n + (f" · {det}" if det else ""), flush=True)
def app(expr, ms=60000):
    r = subprocess.run(["node", str(AQUI / "cdp.mjs"), CDP, expr, str(ms)], capture_output=True, text=True, timeout=ms / 1000 + 10)
    try: return json.loads(r.stdout.strip() or "null")
    except Exception: return {"error": r.stdout + r.stderr}
TEXTO_APP = "(() => { const c = document.querySelector('.wd-cuerpo'); return c ? c.innerText : null; })()"
with sync_playwright() as pw:
    # Sin Internet: la nube no resuelve (como en un salón sin red).
    b = pw.chromium.launch(args=["--host-resolver-rules=MAP nomcolab.dpdns.org 127.0.0.1:1"])
    ctx = b.new_context(viewport={"width": 1300, "height": 850}); p = ctx.new_page(); err = []
    p.on("pageerror", lambda e: err.append(str(e)[:200]))
    p.goto(f"{TEL}/unirse?c={COD}"); p.wait_for_timeout(4000)
    botones = p.evaluate("() => [...document.querySelectorAll('button')].filter(x => x.offsetParent).map(x => x.innerText.trim()).filter(Boolean)")
    anota("la computadora sin Internet abre el enlace del teléfono", any("alias" in x.lower() for x in botones), str(botones[:6]))
    for t in ("Entrar con un alias", "Entrar sin cuenta", "Entrar"):
        try: p.get_by_role("button", name=t).first.click(timeout=2500); break
        except Exception: pass
    p.wait_for_timeout(4000)
    yo = p.evaluate("() => { try { return NOM.api.usuario(); } catch (e) { return null; } }") or {}
    anota("entra con un alias al azar (sin cuenta)", bool(yo.get("alias")), yo.get("nombre", ""))
    p.goto(f"{TEL}/#/doc/{DOC}"); p.wait_for_selector(".wd-cuerpo", timeout=30000); p.wait_for_timeout(1500)
    anota("abre el documento del teléfono", True)
    # La computadora escribe
    p.click(".wd-cuerpo"); p.keyboard.press("Control+End"); p.keyboard.type("Escrito en la computadora. ")
    t0 = time.time(); llego = False
    while time.time() - t0 < 20:
        tx = app(TEXTO_APP) or ""
        if isinstance(tx, str) and "Escrito en la computadora" in tx: llego = True; break
        time.sleep(0.5)
    anota("lo de la computadora le llega a la app", llego, f"{time.time() - t0:.1f} s")
    # La app escribe (texto real en su editor: insertText por el foco de la página)
    r = app("""(async () => { const c = document.querySelector('.wd-cuerpo'); c.focus(); const s = getSelection(); const r = document.createRange();
      r.selectNodeContents(c); r.collapse(false); s.removeAllRanges(); s.addRange(r); document.execCommand('insertText', false, 'Escrito en la app. ');
      await new Promise((ok) => setTimeout(ok, 1500)); return c.innerText; })()""")
    t0 = time.time()
    try:
        p.wait_for_function("() => document.querySelector('.wd-cuerpo').innerText.includes('Escrito en la app')", timeout=20000)
        anota("lo de la app le llega a la computadora", True, f"{time.time() - t0:.1f} s")
    except Exception:
        anota("lo de la app le llega a la computadora", False, str(r)[:200])
    # Se corta la red de la computadora, escribe, y vuelve
    ctx.set_offline(True); p.wait_for_timeout(500)
    p.click(".wd-cuerpo"); p.keyboard.press("Control+End"); p.keyboard.type("Escrito sin red. ")
    p.wait_for_timeout(3000)
    ctx.set_offline(False)
    t0 = time.time(); llego = False
    while time.time() - t0 < 30:
        tx = app(TEXTO_APP) or ""
        if isinstance(tx, str) and "Escrito sin red" in tx: llego = True; break
        time.sleep(1)
    anota("lo escrito sin red le llega a la app al volver la red", llego, f"{time.time() - t0:.1f} s")
    final_pc = p.evaluate("() => document.querySelector('.wd-cuerpo').innerText")
    final_app = app(TEXTO_APP) or ""
    anota("los dos quedan con el mismo texto", final_pc.strip() == str(final_app).strip(), repr(final_pc[:120]))
    anota("sin errores de JavaScript en la computadora", not err, str(err[:2]))
    p.screenshot(path=str(Path.home() / ".cache/nom-pruebas/registros/reales/telefono-computadora.png"))
    b.close()
print(f"\nTeléfono que comparte: {sum(1 for r in res if r[1])} de {len(res)} bien")
sys.exit(0 if all(r[1] for r in res) else 1)
