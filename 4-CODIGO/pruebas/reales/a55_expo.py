"""La app en el A55 (o cualquier teléfono por adb) con los documentos de la expo, con toques reales:

  NOM_A55=SERIE python3 pruebas/reales/a55_expo.py "Expo NOM" "Presentación del proyecto (NOM)" "test_features" …

Abre el espacio, abre cada documento por su nombre, espera a que se dibuje, deja una captura y revisa que no haya en
pantalla avisos de error («Actualiza NOM», «No se pudo», «Algo falló», «Sin conexión»). Al final vuelve al espacio.
(El APK publicable no deja depurar su WebView: todo se hace como una persona, con el árbol de accesibilidad.)
"""
import sys, time, re
import a55

MALOS = re.compile(r"Actualiza NOM|No se pudo|Algo falló|Sin conexión|no responde|Error", re.I)


def abrir_espacio(nombre):
    for _ in range(3):
        n = a55.buscar(nombre, exacto=True)
        if n: a55.tocar(n["cx"], n["cy"]); time.sleep(2.5); return True
        a55.tecla(4); time.sleep(1.5)
    return False


def main():
    espacio, docs = sys.argv[1], sys.argv[2:]
    a55.sh("input keyevent KEYCODE_WAKEUP"); time.sleep(0.5)
    bien = mal = 0
    if not abrir_espacio(espacio): print("MAL  no encontré el espacio", espacio); sys.exit(1)
    for d in docs:
        n = None
        for _ in range(6):
            n = a55.buscar(d)
            if n and n["cy"] < 2150: break
            a55.deslizar(540, 1700, 540, 900); time.sleep(0.8); n = None
        if not n: print("MAL ", d, "(no aparece en el espacio)"); mal += 1; continue
        a55.tocar(n["cx"], n["cy"]); time.sleep(7)
        textos = " | ".join((x["t"] or x["d"]) for x in a55.nodos() if (x["t"] or x["d"]))
        f = a55.cap("expo-" + re.sub(r"\W+", "_", d)[:40], 360)
        m = MALOS.search(textos)
        if m: print("MAL ", d, "→", m.group(0), f); mal += 1
        else: print("ok  ", d, f); bien += 1
        a55.tecla(4); time.sleep(2.5)
        if not a55.buscar(espacio): abrir_espacio(espacio)
    print(f"\nA55: {bien} bien, {mal} mal")


if __name__ == "__main__":
    main()
