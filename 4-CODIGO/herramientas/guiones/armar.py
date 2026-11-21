#!/usr/bin/env python3
"""Arma los patrones de división de palabras de NOM a partir de lo que hizo Word 2024 real.

  armar.py [COSECHA.json …]

Sin argumentos usa las listas guardadas aquí (palabras-es-MX.txt.gz y palabras-en-US.txt.gz: cada palabra con un
guion donde Word la puede partir). Con argumentos, primero junta esas listas con las cosechas nuevas de
pruebas/office/analisis/guiones_word.py ({"es-MX": {"palabra": [puntos]}, …}) y las guarda.

Después, por idioma: patgen (TeX) aprende patrones de Liang que dan los mismos puntos que Word (8 niveles; ver
PARAMETROS), se revisa cada palabra con el mismo algoritmo de js/guiones.js y lo que no salga igual va como excepción.
Escribe app/src/main/resources/static/js/guiones-es.js y guiones-en-us.js. Necesita patgen (paquete texlive-binaries).
"""
import gzip, json, os, subprocess, sys, tempfile
from pathlib import Path

AQUI = Path(__file__).resolve().parent
JS = AQUI.parent.parent / "app/src/main/resources/static/js"
IDIOMAS = {"es-MX": ("es", "guiones-es.js", "español (México)"), "en-US": ("en-us", "guiones-en-us.js", "inglés (Estados Unidos)")}
# niveles de patgen: (hyph_start, hyph_finish) y por nivel (pat_start, pat_finish) (good, bad, threshold)
PARAMETROS = "1 8\n2 5\n1 2 20\n2 5\n2 1 8\n2 6\n1 4 7\n2 7\n3 2 1\n2 8\n1 4 1\n2 8\n1 1 1\n2 9\n1 4 1\n2 9\n1 1 1\nn\n"
MIN = (2, 2)   # Word deja al menos 2 letras en cada lado (en el texto corrido pide 3 al principio del renglón siguiente)
LARGO_PATRONES = 30   # las palabras más largas no entran a patgen: van completas como excepciones
LARGO_MAX = 120


def leer_lista(ruta):
    out = {}
    if not ruta.exists(): return out
    for e in gzip.open(ruta, "rt", encoding="utf-8").read().split():
        w, pts, k = "", [], 0
        for ch in e:
            if ch == "-": pts.append(k)
            else: w += ch; k += 1
        out[w] = pts
    return out


def guardar_lista(ruta, d):
    with gzip.open(ruta, "wt", encoding="utf-8", compresslevel=9) as o:
        for w in sorted(d):
            p = set(d[w])
            o.write("".join(w[i] + ("-" if (i + 1) in p else "") for i in range(len(w))) + "\n")


def liang(t, largo, w):
    s = "." + w + "."
    n = len(s)
    v = [0] * (n + 1)
    for i in range(n):
        for k in range(i + 1, min(n, i + largo) + 1):
            p = t.get(s[i:k])
            if p:
                for q, x in enumerate(p):
                    if x > v[i + q]: v[i + q] = x
    return [m for m in range(MIN[0], len(w) - MIN[1] + 1) if v[m + 1] % 2 == 1]


def main():
    for idioma in IDIOMAS:
        ruta = AQUI / f"palabras-{idioma}.txt.gz"
        d = leer_lista(ruta)
        antes = len(d)
        for f in sys.argv[1:]:
            for w, p in json.load(open(f, encoding="utf-8")).get(idioma, {}).items():
                if 2 <= len(w) <= LARGO_MAX and all("a" <= c <= "z" or ("à" <= c <= "ÿ" and c != "÷") for c in w):
                    d[w] = [x for x in p if MIN[0] <= x <= len(w) - MIN[1]]
        if sys.argv[1:]: guardar_lista(ruta, d)
        print(f"{idioma}: {len(d)} palabras ({len(d) - antes} nuevas)", flush=True)
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            dic, tra, vac, pat = tmp / "d.dic", tmp / "d.tra", tmp / "vacio.pat", tmp / "d.pat"
            with open(dic, "w", encoding="latin-1") as o:
                for w in sorted(x for x in d if len(x) <= LARGO_PATRONES):
                    p = set(d[w]); o.write("".join(w[i] + ("-" if (i + 1) in p else "") for i in range(len(w))) + "\n")
            letras = sorted(set("".join(d)))
            with open(tra, "w", encoding="latin-1") as o:
                o.write(f"{MIN[0]:2d}{MIN[1]:2d}\n")
                for c in letras:
                    u = c.upper()
                    o.write(f" {c} {u} \n" if u != c and len(u) == 1 and ord(u) <= 255 else f" {c} \n")
            vac.write_text("")
            r = subprocess.run(["patgen", str(dic), str(vac), str(pat), str(tra)], input=PARAMETROS.encode(), capture_output=True, cwd=tmp)
            linea = [l for l in r.stdout.decode("latin-1").splitlines() if "%" in l][-1:]
            patrones = pat.read_text(encoding="latin-1").split()
        t, largo = {}, 0
        for p in patrones:
            letras = "".join(c for c in p if not c.isdigit())
            v = [0] * (len(letras) + 1); i = 0
            for c in p:
                if c.isdigit(): v[i] = int(c)
                else: i += 1
            t[letras] = v
            largo = max(largo, len(letras))
        exc = []
        for w in sorted(d):
            if len(w) < MIN[0] + MIN[1]: continue
            if len(w) > LARGO_PATRONES or liang(t, largo, w) != sorted(d[w]):
                p = set(d[w]); exc.append("".join(w[i] + ("-" if (i + 1) in p else "") for i in range(len(w))))
        juego, archivo, nombre = IDIOMAS[idioma]
        cab = (f"/* NOM · Dónde parte Word 2024 las palabras con «Guiones automáticos» en {nombre}.\n"
               f"   Patrones de Liang aprendidos con patgen de lo que hizo Word real con {len(d)} palabras (herramientas/guiones:\n"
               f"   LEEME.md y armar.py); las que los patrones no dan igual van como excepciones ({len(exc)}). Generado: no editar. */\n")
        cuerpo = f'(window.NOM_GUIONES = window.NOM_GUIONES || {{}})["{juego}"] = {{ min: [{MIN[0]}, {MIN[1]}], patrones: {json.dumps(" ".join(patrones), ensure_ascii=True)}, excepciones: {json.dumps(" ".join(exc), ensure_ascii=True)} }};\n'
        (JS / archivo).write_text(cab + cuerpo, encoding="utf-8")
        print(f"   patgen: {linea[0].strip() if linea else '?'} · {len(patrones)} patrones, {len(exc)} excepciones → js/{archivo} ({(JS / archivo).stat().st_size} bytes)", flush=True)


if __name__ == "__main__":
    main()
