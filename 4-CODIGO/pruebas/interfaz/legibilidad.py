#!/usr/bin/env python3
"""Mide qué tan fácil se leen los textos que ve la persona en NOM (sin navegador: lee el código).

Saca los textos visibles de las plantillas de `static/js/*.js` (cadenas con palabras en español: avisos,
títulos, botones, `aria-label`, `placeholder`, `data-micro`…), de `static/index.html` y de la ayuda
(`4-CODIGO/ayuda/contenido/**/*.md`), y calcula:

  · el índice de Flesch-Szigriszt con la escala INFLESZ (Barrio-Cantalejo et al., 2008,
    An. Sist. Sanit. Navar. 31(2):135-152, doi:10.4321/S1137-66272008000300004):
      IFSZ = 206.835 − 62.3 · (sílabas / palabras) − (palabras / frases)
      < 40 muy difícil · 40-55 algo difícil · 55-65 normal · 65-80 bastante fácil · > 80 muy fácil
  · las palabras del glosario que NO deben verse (GUIA-FACIL-DE-USAR.md §3).

El índice solo es estable con textos de varias palabras: se califica cada texto de 12 palabras o más y,
además, todo junto por archivo. La meta es «bastante fácil» (65) o mejor; los textos de la interfaz
de 4 a 11 palabras se listan solo si tienen una palabra prohibida.

  python3 pruebas/interfaz/legibilidad.py [--todo] [--json salida.json]

Sale con código 1 si hay palabras prohibidas en la interfaz o algún archivo queda por debajo de la meta.
"""
import argparse, glob, json, os, re, sys, unicodedata

RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
STATIC = os.path.join(RAIZ, "app", "src", "main", "resources", "static")
AYUDA = os.path.join(RAIZ, "ayuda", "contenido")
META = 65
# Un archivo se califica en conjunto solo con 10 textos o más (con menos, el índice no es estable), y no se califican
# los catálogos de nombres (estilos, transiciones e íconos de Office, contenido de ejemplo de los diagramas): son
# nombres sueltos, no frases, y la fórmula los trata como una sola frase larga.
MIN_TEXTOS = 10
CATALOGO = re.compile(r"(-datos|iconos|transiciones-office|diagramas)\.js$")

# Glosario (§3): palabra técnica → lo que se dice. Se buscan como palabra completa, sin importar mayúsculas.
PROHIBIDAS = {
    r"servidor(es)?": "el equipo que comparte · NOM en Internet · otro NOM",
    r"red local|LAN": "el mismo Wi-Fi",
    r"sincroniz\w*": "guardar una copia en otro NOM · enviar tus cambios · Guardando…",
    r"token": "(no se muestra)",
    r"solo lectura|sólo lectura": "Solo ver / Solo puede ver",
    r"revoca\w*|caduca\w*|caducad\w*": "Anular, vencer",
    r"exporta\w*": "Descargar",
    r"caracteres|carácter": "letras o números",
    r"bienvenid[oa]": "Entrar con tu cuenta · Ya estás en NOM",
    # «cifrado de extremo a extremo» se acepta: es la frase que ya conoce quien usa WhatsApp (nombre de la función).
    r"cifrad[oa]s?(?! de extremo)|cifrar|AES": "protegido con tu contraseña",
}
PROH_RE = re.compile(r"\b(" + "|".join(PROHIBIDAS) + r")\b", re.I)

VOCALES = "aeiouáéíóúü"
FUERTES = "aeoáéíóú"   # las débiles con acento (í, ú) rompen el diptongo: cuentan como fuertes


def silabas(p):
    """Sílabas de una palabra en español (aproximado: núcleos vocálicos con diptongos e hiatos)."""
    p = p.lower()
    p = re.sub(r"(?<=[qg])u(?=[eiéí])", "", p)   # «que», «gui»: la u no suena
    if p.endswith("y") and len(p) > 1 and p[-2] not in VOCALES: p = p[:-1] + "i"
    n, prev = 0, ""
    for ch in p:
        if ch in VOCALES:
            if not prev: n += 1
            elif prev in FUERTES and ch in FUERTES: n += 1   # hiato
            prev = ch
        else:
            prev = ""
    return max(1, n)


def palabras(t):
    return re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+", t)


def frases(t):
    return max(1, len([f for f in re.split(r"[.!?¿¡:;…\n]+", t) if palabras(f)]))


def inflesz(t):
    ps = palabras(t)
    if not ps: return None
    s = sum(silabas(p) for p in ps)
    return 206.835 - 62.3 * s / len(ps) - len(ps) / frases(t)


def grado(v):
    return "muy difícil" if v < 40 else "algo difícil" if v < 55 else "normal" if v < 65 else "bastante fácil" if v < 80 else "muy fácil"


def limpio(s):
    s = re.sub(r"\$\{[^{}]*(\{[^{}]*\}[^{}]*)*\}", " ", s)   # interpolaciones
    s = re.sub(r"<[^>]+>", " ", s)                          # etiquetas
    s = re.sub(r"&[a-z]+;|\\[nt]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


CODIGO = re.compile(r"[{};=]|=>|\(\)|^\.|^#|^[a-z-]+:|\bvar\(|\bpx\b|^\w+$|\.js\b|/api/|https?:")
CADENA = re.compile(r'"((?:[^"\\\n]|\\.)*)"|\'((?:[^\'\\\n]|\\.)*)\'|`((?:[^`\\]|\\.)*)`', re.S)
ATRIB = re.compile(r'(?:aria-label|title|placeholder|data-micro|alt)="([^"$]*)"')


def es_texto(s):
    if len(palabras(s)) < 2: return False
    if CODIGO.search(s): return False
    if not re.search(r"[áéíóúñ¿¡]|\b(el|la|los|las|de|en|un|una|tu|para|con|que|se|no|lo|al|del|y|o)\b", s): return False
    letras = sum(c.isalpha() for c in s)
    return letras >= 0.6 * len(s.replace(" ", ""))


def textos_js(ruta):
    fuente = open(ruta, encoding="utf-8").read()
    # Sin comentarios (los de bloque y los de línea que empiezan la línea o van después de código).
    sin = re.sub(r"/\*.*?\*/", lambda m: "\n" * m.group(0).count("\n"), fuente, flags=re.S)
    sin = re.sub(r"(^|[^:\\'\"`])//[^\n]*", r"\1", sin)
    sin = re.sub(r"console\.\w+\([^;\n]*", "", sin)   # lo de la consola no lo ve la persona
    out = []
    for m in CADENA.finditer(sin):
        crudo = m.group(1) or m.group(2) or m.group(3) or ""
        linea = sin.count("\n", 0, m.start()) + 1
        piezas = ATRIB.findall(crudo) if "<" in crudo else []
        # Texto entre etiquetas y en atributos visibles
        for t in piezas + [limpio(crudo)]:
            for parte in t.split("|"):      # data-micro="Título|explicación"
                parte = parte.strip()
                if es_texto(parte): out.append((linea, parte))
    return out


def textos_html(ruta):
    h = open(ruta, encoding="utf-8").read()
    h = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", h, flags=re.S)
    out = [(0, t) for t in ATRIB.findall(h) if es_texto(t)]
    for t in re.split(r"<[^>]+>", h):
        t = re.sub(r"\s+", " ", t).strip()
        if es_texto(t): out.append((0, t))
    return out


def textos_md(ruta):
    md = open(ruta, encoding="utf-8").read()
    md = re.sub(r"^---.*?---", "", md, flags=re.S)
    md = re.sub(r"```.*?```", " ", md, flags=re.S)
    md = re.sub(r"`[^`]*`|\]\([^)]*\)|[#*_>|\[]", " ", md)
    return [(i + 1, l.strip()) for i, l in enumerate(md.split("\n")) if es_texto(l.strip())]


def principal():
    ap = argparse.ArgumentParser(); ap.add_argument("--todo", action="store_true"); ap.add_argument("--json")
    a = ap.parse_args()
    fuentes = [("interfaz", f, textos_js(f)) for f in sorted(glob.glob(os.path.join(STATIC, "js", "*.js")))]
    fuentes.append(("interfaz", os.path.join(STATIC, "index.html"), textos_html(os.path.join(STATIC, "index.html"))))
    fuentes += [("ayuda", f, textos_md(f)) for f in sorted(glob.glob(os.path.join(AYUDA, "**", "*.md"), recursive=True))]
    malos, prohibidas, resumen = [], [], []
    for zona, f, ts in fuentes:
        if not ts: continue
        rel = os.path.relpath(f, RAIZ)
        todo = ". ".join(t for _, t in ts)
        v = inflesz(todo)
        resumen.append((zona, rel, len(ts), v))
        for ln, t in ts:
            for m in PROH_RE.finditer(t):
                prohibidas.append((zona, rel, ln, m.group(0), t[:140]))
            if len(palabras(t)) >= 12:
                vi = inflesz(t)
                if vi < META: malos.append((zona, rel, ln, round(vi), grado(vi), t[:160]))
    print("== Índice INFLESZ por archivo (todo el texto visible junto); meta ≥", META)
    for zona, rel, n, v in sorted(resumen, key=lambda x: x[3]):
        if a.todo or v < META:
            print(f"  {v:6.1f} {grado(v):15} {n:4} textos  {rel}")
    total = ". ".join(t for _, _, ts in fuentes for _, t in ts)
    vi_int = inflesz(". ".join(t for z, _, ts in fuentes if z == "interfaz" for _, t in ts))
    vi_ay = inflesz(". ".join(t for z, _, ts in fuentes if z == "ayuda" for _, t in ts))
    print(f"\n  INTERFAZ en conjunto: {vi_int:.1f} ({grado(vi_int)})   AYUDA en conjunto: {vi_ay:.1f} ({grado(vi_ay)})")
    print(f"\n== Textos de 12 palabras o más por debajo de {META}: {len(malos)}")
    for zona, rel, ln, v, g, t in sorted(malos, key=lambda x: x[3])[: (None if a.todo else 60)]:
        print(f"  {v:4} {zona:8} {rel}:{ln}  {t}")
    pi = [x for x in prohibidas if x[0] == "interfaz"]
    print(f"\n== Palabras del glosario que se ven: interfaz {len(pi)}, ayuda {len(prohibidas) - len(pi)}")
    for zona, rel, ln, w, t in prohibidas if a.todo else pi:
        print(f"  {zona:8} {rel}:{ln}  «{w}»  {t}")
    if a.json:
        json.dump({"resumen": resumen, "malos": malos, "prohibidas": prohibidas, "interfaz": vi_int, "ayuda": vi_ay}, open(a.json, "w"), ensure_ascii=False, indent=1)
    archivos_bajos = [r for r in resumen if r[0] == "interfaz" and r[3] < META and r[2] >= MIN_TEXTOS and not CATALOGO.search(r[1])]
    sys.exit(1 if pi or archivos_bajos else 0)


if __name__ == "__main__":
    principal()
