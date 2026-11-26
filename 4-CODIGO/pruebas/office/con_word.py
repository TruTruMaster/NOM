"""Entregables con Word real: guarda un .docx otra vez con Word (propiedades de Word, sin rastros del programa que lo
generó) y saca su PDF con las fuentes reales. Usa la PC de Windows con Office (remoto.py).

  from con_word import guardar_y_pdf
  guardar_y_pdf("ENTREGA/x.docx", autor="WACK", titulo="…")   # reemplaza x.docx y escribe x.pdf
"""
import json, os, tempfile, time
from pathlib import Path
import remoto


def disponible():
    return remoto.disponible()


def guardar_y_pdf(docx, autor, titulo=None, carpeta="nomoffice/entregables"):
    docx = Path(docx).resolve()
    ident = "ent" + str(int(time.time() * 1000))[-9:]
    base = carpeta.split("/", 1)[1] if carpeta.startswith("nomoffice/") else carpeta
    with tempfile.TemporaryDirectory() as t:
        t = Path(t)
        entrada = t / f"{ident}-entrada.docx"; entrada.write_bytes(docx.read_bytes())
        datos = t / f"{ident}.json"
        datos.write_text(json.dumps({"entrada": f"{base}/{entrada.name}", "docx": f"{base}/{ident}-salida.docx",
                                     "pdf": f"{base}/{ident}-salida.pdf", "salida": f"{base}/{ident}-resultado.json",
                                     "autor": autor, "titulo": titulo}, ensure_ascii=False), encoding="utf-8")
        remoto.subir([entrada, datos], carpeta)
        r = remoto.correr_office(ident, "word_guardar_pdf.py", f"{base}/{datos.name}", timeout=900)
        res = t / "resultado.json"
        if not remoto.bajar(f"{carpeta}/{ident}-resultado.json", res):
            raise RuntimeError("Word no dejó resultado: " + (r.stdout or "")[-400:])
        info = json.loads(res.read_text(encoding="utf-8"))
        if not (info.get("docx") and info.get("pdf")):
            raise RuntimeError(f"Word no pudo con {docx.name}: {info}")
        nuevo, pdf = t / "salida.docx", t / "salida.pdf"
        remoto.bajar(f"{carpeta}/{ident}-salida.docx", nuevo)
        remoto.bajar(f"{carpeta}/{ident}-salida.pdf", pdf)
        docx.write_bytes(nuevo.read_bytes())
        docx.with_suffix(".pdf").write_bytes(pdf.read_bytes())
        remoto.ejecutar(f'del /q "{carpeta}\\{ident}-*" "{carpeta}\\{ident}.json"'.replace("/", "\\"))
        return info
