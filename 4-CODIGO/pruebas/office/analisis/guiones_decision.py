#!/usr/bin/env python3
"""Cuándo parte Word una palabra al final del renglón («Guiones automáticos»): compara los renglones de Word
(w15-guiones-*.docx, hechos con windows/hacer_guiones2.py y medidos con verdad.py) con varias reglas, usando los anchos
de las fuentes de NOM (las medidas de Office) y los puntos donde Word puede partir cada palabra (guiones_word.py).

  guiones_decision.py PUNTOS.json [w15-guiones-z18 w15-guiones-z7 …]

Cada renglón se simula desde donde Word lo empezó (así cada decisión se mide sola) y se cuenta en cuántos la regla
deja exactamente lo que dejó Word.
"""
import json, os, re, sys
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from acomodo_word import avances, NOM_FUENTE, CACHE  # noqa: E402

LETRA = re.compile(r"[a-záéíóúñüA-ZÁÉÍÓÚÑÜ]+")


def renglones(verdad):
    out = []
    for p in verdad["paginas"]:
        for z in p["zonas"]:
            if z.get("historia") != "cuerpo": continue
            out += [l["texto"] for l in z["lineas"]]
    pars, cur = [], []
    for t in out:
        cur.append(t)
        if t.endswith("\r"): pars.append(cur); cur = []
    return pars


def main():
    puntos = {}
    for f in sys.argv[1].split(","):
        for idi, ws in json.load(open(f, encoding="utf-8")).items(): puntos.setdefault(idi, {}).update(ws)
    nombres = sys.argv[2:] or ["w15-guiones-z18", "w15-guiones-z7", "w15-guiones-z36", "w15-guiones-lim2"]
    docs = []
    letras = set(" -")
    for n in nombres:
        meta = json.load(open(CACHE / "office-real/archivos" / (n + ".json"), encoding="utf-8"))
        verdad = json.load(open(CACHE / "office-real/verdad" / (n + ".verdad.json"), encoding="utf-8"))
        pars = renglones(verdad)
        assert len(pars) == len(meta["parrafos"]), (n, len(pars), len(meta["parrafos"]))
        docs.append((n, meta, pars))
        for p in meta["parrafos"]: letras |= set(p["texto"])
    cache_adv = CACHE / "office-real" / "w15-guiones.avances.json"
    adv = json.load(open(cache_adv)) if cache_adv.exists() else None
    if not adv or not set(letras) <= set(next(iter(adv.values()))["m"]):
        adv = avances(sorted(set(NOM_FUENTE[f] for f in ("Calibri", "Times New Roman", "Arial", "Cambria"))), sorted(letras))
        json.dump(adv, open(cache_adv, "w"))
    m = {f: adv[NOM_FUENTE[f]]["m"] for f in ("Calibri", "Times New Roman", "Arial", "Cambria")}

    # casos: cada renglón que no es el último, con lo que Word dejó en él
    casos = []
    faltan = Counter()
    for n, meta, pars in docs:
        zona, limite = meta["zona"], meta["limite"]
        for p, ls in zip(meta["parrafos"], pars):
            texto = p["texto"]
            assert "".join(ls).rstrip("\r") == texto, (n, texto[:50], "".join(ls)[:50])
            pos = 0
            seguidos = 0
            for i, l in enumerate(ls[:-1]):
                ini, fin = pos, pos + len(l)
                pos = fin
                partida = fin < len(texto) and texto[fin - 1] != " " and texto[fin] != " "
                casos.append(dict(doc=n, f=p["fuente"], tam=p["tam"], alin=p["alin"], A=meta["ancho"] - p["sangria"], idioma=p["idioma"],
                                  zona=zona, limite=limite, texto=texto, ini=ini, fin=fin, partida=partida, seguidos=seguidos))
                seguidos = seguidos + 1 if partida else 0
    print(f"{len(casos)} renglones con decisión; {sum(c['partida'] for c in casos)} terminan en una palabra partida")

    def ancho(t, f, tam):
        a = m[f]
        return sum(a[ch] for ch in t) * tam

    def pts_de(pal, idioma):
        w = pal.lower()
        r = puntos.get(idioma, {}).get(w)
        if r is None: faltan[(idioma, w)] += 1; return []
        return r

    def simular(c, regla):
        """Dónde termina el renglón que empieza en c['ini'] (posición en el texto) según la regla."""
        t, f, tam, A = c["texto"], c["f"], c["tam"], c["A"]
        esp = m[f][" "] * tam
        guion = m[f]["-"] * tam
        # palabras desde ini (separadas por un espacio)
        i = c["ini"]
        fin = i
        ancho_dentro = 0.0
        n_esp = 0
        primero = True
        while True:
            j = t.find(" ", i)
            if j < 0: j = len(t)
            pal = t[i:j]
            w_pal = ancho(pal, f, tam)
            nuevo = ancho_dentro + (0 if primero else esp) + w_pal
            cabe = nuevo <= A + 1e-6
            if not cabe and c["alin"] == "just" and not primero and regla.get("encoge", True):
                # justificado: entra si se pasa a lo más la mitad del hueco y ≤ 25 % por espacio
                hueco = A - ancho_dentro
                sobra = nuevo - A
                if sobra <= hueco / 2 and sobra <= (n_esp + 1) * esp * 0.25: cabe = True
            if cabe:
                ancho_dentro = nuevo
                if not primero: n_esp += 1
                primero = False
                fin = j
                if j >= len(t): return len(t)
                i = j + 1
                continue
            if primero:
                return j   # (una palabra sola que no cabe: no pasa en estos párrafos)
            # la palabra no cabe: ¿se parte?
            hueco = A - ancho_dentro
            lim_ok = c["limite"] == 0 or c["seguidos"] < c["limite"]
            zona = c["zona"] * regla.get("kzona", 1.0)
            cond = hueco >= zona if regla.get("ge", True) else hueco > zona
            if regla.get("hueco_con_espacio"): cond = (hueco - esp) >= zona if regla.get("ge", True) else (hueco - esp) > zona
            if lim_ok and cond:
                # letras de la palabra: lo de antes y después (signos) se quedan con su parte
                mm = LETRA.search(pal)
                if mm and mm.start() == 0 or (mm and regla.get("signos_antes", True)):
                    base = mm.group(0)
                    pts = pts_de(base, c["idioma"])
                    for q in sorted(pts, reverse=True):
                        pre = pal[:mm.start() + q]
                        w_pre = ancho(pre, f, tam) + (guion if regla.get("con_guion", True) else 0)
                        nuevo = ancho_dentro + esp + w_pre
                        ok = nuevo <= A + 1e-6
                        if not ok and c["alin"] == "just" and regla.get("encoge_partida", False):
                            sobra = nuevo - A
                            if sobra <= hueco / 2 and sobra <= (n_esp + 1) * esp * 0.25: ok = True
                        if ok:
                            return i + mm.start() + q
            return fin

    reglas = {
        "zona ≥, con guion": dict(),
        "zona >, con guion": dict(ge=False),
        "zona ≥, sin guion": dict(con_guion=False),
        "zona ≥, partida encoge": dict(encoge_partida=True),
        "hueco sin el espacio ≥ zona": dict(hueco_con_espacio=True),
        "zona × 0.5": dict(kzona=0.5),
        "sin zona": dict(kzona=0.0),
    }
    for nombre, regla in reglas.items():
        por = defaultdict(lambda: [0, 0])
        malos = []
        for c in casos:
            fin = simular(c, regla)
            # Word deja el espacio del final en el renglón: se compara sin él
            real = c["fin"] - (1 if c["texto"][c["fin"] - 1:c["fin"]] == " " else 0)
            ok = fin == real
            k = (c["doc"], c["alin"])
            por[k][0] += 1
            if not ok:
                por[k][1] += 1
                malos.append((c, fin, real))
        tot = sum(v[0] for v in por.values()); mal = sum(v[1] for v in por.values())
        print(f"\n== {nombre}: {tot - mal}/{tot} ({100 * (tot - mal) / tot:.2f} %)")
        for k, v in sorted(por.items()): print(f"   {k[0]:18s} {k[1]:5s} {v[0] - v[1]:4d}/{v[0]:4d}")
        if nombre == "zona ≥, con guion":
            for c, fin, real in malos[:25]:
                t = c["texto"]
                print(f"   [{c['doc']} {c['f']} {c['tam']} {c['alin']} A={c['A']:.1f} {c['idioma']}] Word: …{t[max(c['ini'], real - 30):real]!r}|{t[real:real + 15]!r}  regla: …{t[max(c['ini'], fin - 30):fin]!r}|")
    if faltan: print("\npalabras sin puntos de Word:", len(faltan), list(faltan)[:10])


if __name__ == "__main__":
    main()
