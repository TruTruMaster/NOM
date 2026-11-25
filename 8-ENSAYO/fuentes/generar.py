"""Arma el ensayo en Word (.docx) y lo pasa a PDF, en dos versiones:

    ENTREGA/Ensayo_Rodriguez_Esparza.docx y .pdf                 4 páginas contando la portada
    ENTREGA/Ensayo_Rodriguez_Esparza_portada_aparte.docx y .pdf  portada + 4 páginas

    ~/.cache/nom-pruebas/venv/bin/python generar.py

El diagrama sale de componentes.puml (plantuml -tpng componentes.puml).
"""
import os, re, subprocess, tempfile
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

AQUI = os.path.dirname(os.path.abspath(__file__))
ENTREGA = os.path.join(AQUI, "..", "..", "ENTREGA")
DIAGRAMA = os.path.join(AQUI, "componentes.png")

TITULO = "NOM como sistema distribuido"
AUTOR = "Rodríguez Esparza Diego Alejandro"

# (tipo, texto, versión): tipo = t (título), p (párrafo), fig (figura), ref (referencia);
# versión = "" en las dos, "L" solo en la larga, "C" solo en la corta. *texto* = cursiva.
TEXTO = [
    ("t", "Introducción", ""),
    ("p", "En la última revisión del proyecto NOM se comentó que el sistema no podía considerarse distribuido, porque el servidor y la base de datos están en el mismo equipo. La sugerencia fue separarlos en direcciones IP distintas, aunque fuera dentro de la misma máquina. Me pareció una observación importante, porque obliga a pensar qué es lo que de verdad hace distribuido a un sistema. En este ensayo defiendo que NOM sí es un sistema distribuido y que eso no depende de dónde esté la base de datos. Para argumentarlo reviso primero cómo definen algunos autores un sistema distribuido, después lo comparo con lo que hace el sistema y al final respondo a la observación.", "L"),
    ("p", "En la última revisión del proyecto NOM se comentó que el sistema no podía considerarse distribuido, porque el servidor y la base de datos están en el mismo equipo. La sugerencia fue separarlos en direcciones IP distintas, aunque fuera dentro de la misma máquina. En este ensayo defiendo que NOM sí es un sistema distribuido y que eso no depende de dónde esté la base de datos.", "C"),

    ("t", "¿Qué es un sistema distribuido?", ""),
    ("p", "Liu (2004) define un sistema distribuido como un conjunto de computadoras independientes, conectadas por una red, que pueden colaborar en una tarea, y aclara que se consideran independientes cuando no comparten memoria ni espacio de ejecución. Coulouris et al. (2012) dan una definición parecida: los componentes están en computadoras conectadas en red y se comunican y coordinan solamente por medio de mensajes. De ahí salen tres características que, según ellos, tiene todo sistema distribuido: la concurrencia, la falta de un reloj global y que cada componente puede fallar por su cuenta. Tanenbaum y van Steen (2008) agregan que, para el usuario, todo el conjunto debe verse como un solo sistema.", "L"),
    ("p", "Liu (2004) define un sistema distribuido como un conjunto de computadoras independientes, conectadas por una red, que pueden colaborar en una tarea; son independientes cuando no comparten memoria ni espacio de ejecución. Coulouris et al. (2012) agregan que esos componentes se comunican y coordinan solo por medio de mensajes, y que de ahí salen tres características: la concurrencia, la falta de un reloj global y las fallas independientes. Tanenbaum y van Steen (2008) añaden que el conjunto debe verse como un solo sistema para el usuario.", "C"),
    ("p", "En ninguna de estas definiciones se menciona la base de datos. Separar la base de datos del servidor es lo que Tanenbaum y van Steen (2008) llaman distribución vertical, que es una manera de repartir las capas de una aplicación en varias máquinas, pero no es lo que define a un sistema distribuido.", ""),

    ("t", "Cómo funciona NOM", ""),
    ("p", "NOM es una aplicación para editar textos, hojas de cálculo, presentaciones y diagramas entre varias personas al mismo tiempo. Puede funcionar en Internet, con un servidor en una Raspberry Pi, o sin Internet, cuando una computadora o un teléfono Android hace de servidor y los demás se conectan a la misma red Wi-Fi. En la figura 1 se muestran sus componentes principales.", ""),
    ("p", "En el caso sin Internet hay un teléfono que funciona como servidor y varios teléfonos o computadoras que se conectan a él. Cada uno tiene su propia dirección IP, su memoria y su procesador, así que son las computadoras independientes de las que habla Liu. Estos equipos no comparten memoria y todo lo que saben unos de otros les llega por mensajes: peticiones HTTP para guardar y consultar, un canal de eventos (SSE) por el que el servidor manda los cambios de los demás, y anuncios por UDP para que la app encuentre el servidor en la red sin tener que escribir la IP. Es el modelo cliente-servidor con sockets que explica Liu (2004), aplicado a un sistema completo.", "L"),
    ("p", "En el caso sin Internet, un teléfono funciona como servidor y varios teléfonos o computadoras se conectan a él. Cada uno tiene su propia dirección IP, su memoria y su procesador, así que son las computadoras independientes de las que habla Liu. No comparten memoria y se comunican solo por mensajes: peticiones HTTP, un canal de eventos (SSE) por el que llegan los cambios de los demás y anuncios por UDP para encontrar el servidor en la red. Es el modelo cliente-servidor con sockets que explica Liu (2004).", "C"),
    ("fig", "Diagrama de componentes de NOM", "L"),
    ("p", "Lo que más diferencia a NOM de una página web normal es que cada dispositivo tiene su propia copia del documento. Cuando alguien escribe, el cambio se aplica primero en su equipo y después se manda a los demás, y hasta 50 personas pueden editar el mismo documento a la vez. Liu (2004) menciona las aplicaciones colaborativas (*groupware*) como uno de los paradigmas de la computación distribuida, y NOM entra en esa categoría.", "L"),
        ("p", "Además, cada dispositivo tiene su propia copia del documento y hasta 50 personas pueden editar a la vez, lo que corresponde a las aplicaciones colaborativas (*groupware*) que Liu (2004) incluye entre los paradigmas de la computación distribuida.", "C"),
    ("fig", "Diagrama de componentes de NOM", "C"),
    ("p", "El sistema también tuvo que resolver problemas que solo aparecen cuando hay varias computadoras. Los relojes de los teléfonos casi nunca coinciden, así que no se pueden ordenar los cambios por la hora. Por eso cada cambio lleva un reloj lógico de Lamport (1978) y el identificador del equipo que lo hizo, y con eso todos los dispositivos deciden lo mismo cuando dos personas modifican el mismo dato. El texto se guarda letra por letra con una estructura de tipo CRDT (Shapiro et al., 2011), en la que el resultado final no depende del orden en que lleguen los cambios. También se tomaron en cuenta las fallas: si alguien pierde la conexión puede seguir editando, sus cambios se guardan en una cola en su equipo y se mandan cuando vuelve la red. Si mientras tanto otra persona modificó el mismo documento, lo que se editó sin conexión queda en una rama aparte para compararlo y unirlo después. Por último, un espacio que se trabajó sin Internet en el teléfono se puede sincronizar con el servidor de Internet, así que dos servidores distintos terminan con la misma información.", "L"),
    ("p", "El sistema también resuelve problemas que solo aparecen cuando hay varias computadoras. Como los relojes de los teléfonos no coinciden, cada cambio lleva un reloj lógico de Lamport (1978) y el identificador del equipo, y así todos deciden lo mismo cuando dos personas modifican el mismo dato; el texto usa una estructura de tipo CRDT (Shapiro et al., 2011) en la que el resultado no depende del orden de llegada. Si alguien pierde la conexión sigue editando, sus cambios esperan en una cola y se mandan al volver la red, y si otra persona cambió el documento mientras tanto, lo suyo queda en una rama para compararlo después.", "C"),

    ("t", "La observación sobre la base de datos", ""),
    ("p", "Creo que la observación mezcla dos ideas diferentes. Tener la base de datos en otra máquina es una decisión sobre cómo repartir las capas de la aplicación, pero un sistema cliente-servidor de dos capas también es distribuido; de hecho, Coulouris et al. (2012) ponen la Web como uno de sus ejemplos principales.", "L"),
    ("p", "Creo que la observación mezcla dos ideas. Tener la base de datos en otra máquina es una forma de repartir las capas de la aplicación, pero un sistema cliente-servidor de dos capas también es distribuido; Coulouris et al. (2012) ponen la Web como ejemplo.", "C"),
    ("p", "Además, la base de datos H2 del servidor solo guarda las cuentas, las sesiones y las preferencias. Los documentos, que son lo que realmente comparten los usuarios, están al mismo tiempo en el servidor, en cada dispositivo conectado y, si se configuró, en otro servidor. Es decir, los datos importantes ya están en direcciones IP distintas.", "L"),
    ("p", "Además, la base H2 solo guarda cuentas y sesiones; los documentos, que es lo que se comparte, ya están en el servidor y en cada dispositivo conectado, o sea, en direcciones IP distintas.", "C"),
    ("p", "Sobre la idea de asignar dos direcciones IP dentro del mismo equipo, considero que cumpliría la condición solo en apariencia. Los dos procesos seguirían usando el mismo hardware y la misma fuente de energía, y si la máquina se apaga los dos dejan de funcionar al mismo tiempo, así que no habría fallas independientes. Además, abrir la base de datos a la red agregaría un punto más que proteger, cuando una de las decisiones de seguridad del proyecto fue precisamente que la base no se pueda acceder desde fuera.", "L"),
    ("p", "Asignar dos direcciones IP dentro del mismo equipo cumpliría la condición solo en apariencia: los dos procesos seguirían en el mismo hardware y, si la máquina se apaga, fallarían juntos, así que no habría fallas independientes. Además, abrir la base de datos a la red agregaría un punto más que proteger, cuando una decisión de seguridad del proyecto fue que la base no se pueda acceder desde fuera.", "C"),
    ("p", "También hay que reconocer las limitaciones. En cada instalación hay un servidor que coordina el espacio de trabajo, por lo que la arquitectura es cliente-servidor y no de igual a igual; si ese servidor se cae, los usuarios pueden seguir editando, pero no ven los cambios de los demás hasta que vuelve. Tampoco es una base de datos distribuida en el sentido de Özsu y Valduriez (2020), porque la base de cuentas no está fragmentada ni replicada en varios sitios. Si la observación se refería a eso, tiene razón, pero es un concepto diferente al de sistema distribuido.", "L"),
        ("p", "Eso sí, NOM no es una base de datos distribuida en el sentido de Özsu y Valduriez (2020), porque la base de cuentas no está repartida en varios sitios; si la observación se refería a eso, tiene razón, pero es otro concepto.", "C"),

    ("t", "Conclusión", ""),
    ("p", "Con base en las definiciones revisadas, considero que NOM sí es un sistema distribuido. Funciona con varias computadoras independientes que se comunican por mensajes, trabajan sobre el mismo documento y para el usuario se ven como un solo sistema. Además resuelve problemas propios de este tipo de sistemas, como la falta de un reloj común y las desconexiones. La ubicación de la base de datos es una decisión de arquitectura, no algo que defina si un sistema es distribuido, y separarla de manera artificial en el mismo equipo no cambiaría eso.", "L"),
    ("p", "Con base en las definiciones revisadas, considero que NOM sí es un sistema distribuido: varias computadoras independientes se comunican por mensajes, trabajan sobre el mismo documento y se ven como un solo sistema, y además resuelve problemas propios de estos sistemas, como la falta de un reloj común y las desconexiones. La ubicación de la base de datos es una decisión de arquitectura, no algo que defina si un sistema es distribuido.", "C"),

    ("t", "Referencias", ""),
    ("ref", "Coulouris, G., Dollimore, J., Kindberg, T. y Blair, G. (2012). *Distributed systems: Concepts and design* (5.ª ed.). Addison-Wesley.", ""),
    ("ref", "Lamport, L. (1978). Time, clocks, and the ordering of events in a distributed system. *Communications of the ACM, 21*(7), 558-565. https://doi.org/10.1145/359545.359563", ""),
    ("ref", "Liu, M. L. (2004). *Computación distribuida: Fundamentos y aplicaciones*. Pearson Educación.", ""),
    ("ref", "Özsu, M. T. y Valduriez, P. (2020). *Principles of distributed database systems* (4.ª ed.). Springer.", ""),
    ("ref", "Shapiro, M., Preguiça, N., Baquero, C. y Zawirski, M. (2011). Conflict-free replicated data types. En *Stabilization, Safety, and Security of Distributed Systems* (pp. 386-400). Springer.", ""),
    ("ref", "Tanenbaum, A. S. y van Steen, M. (2008). *Sistemas distribuidos: Principios y paradigmas* (2.ª ed.). Pearson Educación.", ""),
]


def formato(p, alin=WD_ALIGN_PARAGRAPH.JUSTIFY, sangria=True, antes=0, despues=0):
    f = p.paragraph_format
    f.alignment = alin
    f.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    f.space_before, f.space_after = Pt(antes), Pt(despues)
    if sangria: f.first_line_indent = Cm(1.25)


def texto(p, t, negrita=False):
    for i, trozo in enumerate(re.split(r"\*", t)):
        if trozo:
            r = p.add_run(trozo); r.italic = bool(i % 2); r.bold = negrita


def numero_de_pagina(seccion):
    p = seccion.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run()
    for tipo, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if tipo:
            e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), tipo)
        else:
            e = OxmlElement("w:instrText"); e.set(qn("xml:space"), "preserve"); e.text = txt
        r._r.append(e)


def armar(version, ruta):
    d = Document()
    est = d.styles["Normal"]
    est.font.name = "Times New Roman"; est.font.size = Pt(12)
    est.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    d.core_properties.author = AUTOR
    d.core_properties.title = TITULO
    s = d.sections[0]
    s.page_width, s.page_height = Cm(21.59), Cm(27.94)
    s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Cm(2.5)
    s.different_first_page_header_footer = True
    numero_de_pagina(s)

    # Portada
    for _ in range(7): formato(d.add_paragraph(), sangria=False)
    p = d.add_paragraph(); formato(p, WD_ALIGN_PARAGRAPH.CENTER, False); texto(p, "Ensayo")
    p = d.add_paragraph(); formato(p, WD_ALIGN_PARAGRAPH.CENTER, False, 6, 0)
    r = p.add_run(TITULO); r.bold = True; r.font.size = Pt(16)
    for _ in range(9): formato(d.add_paragraph(), sangria=False)
    for linea in (AUTOR, "Grupo 5IM7", "", "Ciudad de México, noviembre de 2026"):
        p = d.add_paragraph(); formato(p, WD_ALIGN_PARAGRAPH.CENTER, False); texto(p, linea)
    p.add_run().add_break(WD_BREAK.PAGE)

    for tipo, t, v in TEXTO:
        if v and v != version: continue
        if tipo == "t":
            p = d.add_paragraph(); formato(p, WD_ALIGN_PARAGRAPH.LEFT, False, 6, 0)
            p.paragraph_format.keep_with_next = True
            texto(p, t, negrita=True)
            if t == "Referencias": p.paragraph_format.page_break_before = False
        elif tipo == "p":
            p = d.add_paragraph(); formato(p); texto(p, t)
        elif tipo == "fig":
            p = d.add_paragraph(); formato(p, WD_ALIGN_PARAGRAPH.LEFT, False, 6, 0)
            p.paragraph_format.keep_with_next = True
            r = p.add_run("Figura 1"); r.bold = True
            p = d.add_paragraph(); formato(p, WD_ALIGN_PARAGRAPH.LEFT, False)
            p.paragraph_format.keep_with_next = True
            r = p.add_run(t); r.italic = True
            p = d.add_paragraph(); formato(p, WD_ALIGN_PARAGRAPH.CENTER, False, 0, 6)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p.add_run().add_picture(DIAGRAMA, width=Cm(15.5))
        elif tipo == "ref":
            p = d.add_paragraph(); formato(p, WD_ALIGN_PARAGRAPH.LEFT, False)
            p.paragraph_format.left_indent = Cm(1.25); p.paragraph_format.first_line_indent = Cm(-1.25)
            texto(p, t)
    d.save(ruta)


def paginas(pdf):
    return int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout).group(1))


for version, nombre, maximo in (("C", "Ensayo_Rodriguez_Esparza", 4), ("L", "Ensayo_Rodriguez_Esparza_portada_aparte", 5)):
    docx = os.path.abspath(os.path.join(ENTREGA, nombre + ".docx"))
    armar(version, docx)
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(["soffice", f"-env:UserInstallation=file://{tmp}/perfil", "--headless", "--convert-to", "pdf",
                        "--outdir", os.path.dirname(docx), docx], check=True, capture_output=True)
    n = paginas(docx[:-5] + ".pdf")
    print(f"{nombre}: {n} páginas (máximo {maximo})" + ("" if n <= maximo else "  ¡SE PASA!"))
