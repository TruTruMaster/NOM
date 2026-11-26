"""Documentos técnicos de NOM (WACK), en Word:
    ENTREGA/5IM7_WACK_Arquitectura.docx y .pdf           Modelo-Vista-Controlador, DAO, Spring Boot y tecnologías
    ENTREGA/5IM7_WACK_Sistemas_Distribuidos.docx y .pdf  NOM como sistema distribuido (documento técnico)
Los diagramas salen de los .puml de esta carpeta (plantuml -tpng -Sdpi=170 *.puml). El .docx final y el PDF los hace
Word en la PC de Windows (4-CODIGO/pruebas/office/con_word.py); sin esa PC, el PDF sale de LibreOffice y se avisa.
Uso: ~/.cache/nom-pruebas/venv/bin/python 9-DOCUMENTOS-TECNICOS/fuentes/generar.py
"""
import datetime, os, subprocess, sys, tempfile
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor

AQUI = os.path.dirname(os.path.abspath(__file__))
ENTREGA = os.path.join(AQUI, "..", "..", "ENTREGA")
TINTA = RGBColor(0x2B, 0x29, 0x52)
EQUIPO = [("Martínez Ortega Uriel Alejandro", "Director general"), ("Soria Flores Josep Aarón", "Project manager"),
          ("González Azuara Arturo Leonardo", "Subdirector de operaciones"), ("Rodríguez Esparza Diego Alejandro", "Desarrollador y QA")]


# ---------------------------------------------------------------- formato
def fondo(celda, color):
    tc = celda._tc.get_or_add_tcPr(); s = OxmlElement("w:shd")
    s.set(qn("w:val"), "clear"); s.set(qn("w:color"), "auto"); s.set(qn("w:fill"), color); tc.append(s)


def pie_numerado(seccion, texto):
    p = seccion.footer.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run(texto + "   Página "); r.font.size = Pt(9)
    for tipo, t in (("begin", None), (None, "PAGE"), ("end", None)):
        rr = p.add_run(); rr.font.size = Pt(9)
        if tipo:
            f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), tipo); rr._r.append(f)
        else:
            i = OxmlElement("w:instrText"); i.set(qn("xml:space"), "preserve"); i.text = t; rr._r.append(i)


def texto(p, t):
    """Texto con **negritas**, *cursivas* y `código` (un * suelto, como en *Api, se queda tal cual)."""
    partes, i = [], 0
    while i < len(t):
        for marca, tipo in (("**", "b"), ("`", "c"), ("*", "i")):
            if t.startswith(marca, i) and marca in t[i + len(marca):]:
                j = t.index(marca, i + len(marca))
                if tipo == "i" and (j == i + 1 or t[i + 1] == " "): continue
                partes.append((t[i + len(marca):j], tipo)); i = j + len(marca); break
        else:
            k = i + 1
            while k < len(t) and t[k] not in "*`": k += 1
            partes.append((t[i:k], "n")); i = k
    for s, tipo in partes:
        r = p.add_run(s)
        if tipo == "i": r.italic = True
        if tipo == "b": r.bold = True
        if tipo == "c": r.font.name = "Consolas"; r.font.size = Pt(10)


class Doc:
    def __init__(self, titulo, clave, subtitulo):
        self.d = Document(); self.titulo, self.clave = titulo, clave; self.fig = self.tab = 0
        st = self.d.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(11)
        st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
        st.paragraph_format.space_after = Pt(6); st.paragraph_format.line_spacing = 1.15
        for n, tam in (("Heading 1", 16), ("Heading 2", 13)):
            h = self.d.styles[n]; h.font.name = "Calibri"; h.font.size = Pt(tam); h.font.bold = True; h.font.color.rgb = TINTA
            h.element.rPr.rFonts.set(qn("w:asciiTheme"), "minorHAnsi")
            h.paragraph_format.space_before = Pt(14 if n == "Heading 1" else 10); h.paragraph_format.space_after = Pt(6)
        cp = self.d.core_properties
        cp.author = cp.last_modified_by = "WACK"; cp.title = titulo; cp.comments = ""; cp.subject = subtitulo; cp.keywords = ""
        cp.created = cp.modified = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0, tzinfo=None)
        s = self.d.sections[0]
        s.page_width, s.page_height = Cm(21.59), Cm(27.94)
        s.left_margin = s.right_margin = Cm(2.5); s.top_margin = s.bottom_margin = Cm(2.3)
        s.different_first_page_header_footer = True
        pie_numerado(s, f"NOM · {titulo} · {clave} · WACK")
        self.portada(subtitulo)

    def portada(self, subtitulo):
        d = self.d
        p = d.add_paragraph(); r = p.add_run("WACK"); r.bold = True; r.font.size = Pt(26); r.font.color.rgb = TINTA
        p = d.add_paragraph(); r = p.add_run("Desarrollo de software para colaborar en tiempo real"); r.font.size = Pt(12)
        for _ in range(6): d.add_paragraph()
        p = d.add_paragraph(); r = p.add_run(f"Documento {self.clave} · Revisión 1"); r.font.size = Pt(11); r.font.color.rgb = TINTA
        p = d.add_paragraph(); r = p.add_run(self.titulo); r.bold = True; r.font.size = Pt(24)
        p = d.add_paragraph(); r = p.add_run(subtitulo); r.font.size = Pt(13)
        p = d.add_paragraph(); r = p.add_run("Proyecto: NOM (página web, app de Android y servidor), versión 0.23.3"); r.font.size = Pt(11)
        for _ in range(7): d.add_paragraph()
        p = d.add_paragraph(); r = p.add_run("Equipo del proyecto"); r.bold = True
        for n, puesto in EQUIPO:
            p = d.add_paragraph(); p.paragraph_format.space_after = Pt(0); p.add_run(n + ", " + puesto.lower())
        d.add_paragraph()
        p = d.add_paragraph("Ciudad de México, noviembre de 2026")
        p.runs[-1].add_break(WD_BREAK.PAGE)

    def h1(self, t): self.d.add_heading(t, level=1)
    def h2(self, t): self.d.add_heading(t, level=2)

    def p(self, t):
        par = self.d.add_paragraph(); par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; texto(par, t); return par

    def lista(self, items):
        for t in items:
            par = self.d.add_paragraph(style="List Bullet"); par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; texto(par, t)

    def tabla(self, titulo, cab, filas, anchos=None):
        self.tab += 1
        t = self.d.add_table(rows=1, cols=len(cab)); t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, c in enumerate(cab):
            cel = t.rows[0].cells[i]; cel.text = ""; r = cel.paragraphs[0].add_run(c); r.bold = True
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.size = Pt(10); fondo(cel, "2B2952")
        for f in filas:
            cs = t.add_row().cells
            for i, v in enumerate(f):
                cs[i].text = ""; par = cs[i].paragraphs[0]; texto(par, v)
                for r in par.runs: r.font.size = Pt(9.5)
        if anchos:
            for fila in t.rows:
                for i, w in enumerate(anchos): fila.cells[i].width = Cm(w)
        par = self.d.add_paragraph(); par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = par.add_run(f"Tabla {self.tab}. {titulo}"); r.italic = True; r.font.size = Pt(9.5)

    def figura(self, archivo, titulo, ancho=15.5):
        self.fig += 1
        par = self.d.add_paragraph(); par.alignment = WD_ALIGN_PARAGRAPH.CENTER; par.paragraph_format.keep_with_next = True
        img = par.add_run().add_picture(os.path.join(AQUI, archivo), width=Cm(ancho))
        img._inline.docPr.set("descr", titulo); img._inline.docPr.set("name", f"Figura {self.fig}")
        img._inline.graphic.graphicData.pic.nvPicPr.cNvPr.set("name", f"Figura {self.fig}")
        par = self.d.add_paragraph(); par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = par.add_run(f"Figura {self.fig}. {titulo}"); r.italic = True; r.font.size = Pt(9.5)

    def referencias(self, refs):
        self.h1("Referencias")
        for t in refs:
            par = self.d.add_paragraph(); par.paragraph_format.left_indent = Cm(1.25); par.paragraph_format.first_line_indent = Cm(-1.25)
            texto(par, t)

    def guardar(self, nombre):
        ruta = os.path.abspath(os.path.join(ENTREGA, nombre + ".docx")); self.d.save(ruta); return ruta


# ================================================================ documento 1: arquitectura
def arquitectura():
    D = Doc("Arquitectura del software", "ARQ-NOM-01", "Modelo-Vista-Controlador, objetos de acceso a datos (DAO), Spring Boot y tecnologías")
    D.h1("1. Introducción")
    D.p("Este documento describe cómo está construido NOM: en qué partes se divide, qué hace cada una y con qué "
        "tecnologías corre. Se explica con dos patrones de diseño que el proyecto usa de principio a fin, el patrón "
        "Modelo-Vista-Controlador (MVC) y el patrón de objetos de acceso a datos (DAO), y con el marco de trabajo del "
        "servidor, Spring Boot. Todo lo que aquí se menciona corresponde al código de la versión 0.23.3 y puede "
        "revisarse en el repositorio del proyecto (carpeta `4-CODIGO`).")
    D.p("NOM tiene tres piezas. La **interfaz** es una página web escrita en JavaScript sin marcos de trabajo (67 "
        "archivos), que se usa igual en el navegador y dentro de la app de Android. El **núcleo** es una biblioteca de "
        "Java sin dependencias (modelo, reglas de negocio y sincronización) que comparten el servidor y el teléfono. El "
        "**servidor** expone el núcleo por HTTP: en la nube y en la computadora del salón es una aplicación de Spring "
        "Boot; en el teléfono que comparte su red es un servidor HTTP propio, más ligero, que repite las mismas rutas.")
    D.tabla("Tamaño de cada pieza (versión 0.23.3)", ["Pieza", "Dónde está", "Tamaño"], [
        ["Interfaz (vista y controlador del cliente)", "`app/src/main/resources/static/js`", "67 archivos, 47 932 líneas"],
        ["Núcleo (modelo, servicios, CRDT, DAO)", "`nucleo/src/main/java`", "8 507 líneas de Java"],
        ["Servidor Spring Boot (controladores, filtros, DAO)", "`app/src/main/java`", "3 425 líneas, 120 rutas"],
        ["Servidor del teléfono", "`android/app/src/main/java`", "3 068 líneas, 90 rutas"]], [6, 6, 4])

    D.h1("2. El patrón Modelo-Vista-Controlador")
    D.h2("2.1 Qué es")
    D.p("El patrón MVC separa una aplicación en tres responsabilidades: el **modelo** guarda los datos y las reglas, la "
        "**vista** los presenta y el **controlador** recibe lo que hace el usuario y decide qué parte del modelo cambia y "
        "qué vista se muestra (Krasner y Pope, 1988). La separación permite cambiar la presentación sin tocar las reglas "
        "y probar cada parte por separado.")
    D.p("En una aplicación web moderna como NOM el patrón aparece en dos niveles, porque la vista ya no se arma en el "
        "servidor: el servidor responde datos en JSON y la página construye la pantalla. La figura 1 muestra los dos "
        "niveles y cómo se conectan.")
    D.figura("capas.png", "Capas de NOM: MVC en el cliente y en el servidor, y el DAO debajo del modelo", 12.5)
    D.h2("2.2 MVC en el servidor (Spring MVC)")
    D.p("El servidor usa *Spring MVC*, el módulo de Spring que implementa el patrón. Toda petición entra por un único "
        "punto, el `DispatcherServlet`, que la entrega al controlador que corresponde según su ruta y su método HTTP.")
    D.lista([
        "**Controladores**: las clases `*Api` del paquete `servidor.api`, anotadas con `@RestController`. Cada método "
        "atiende una ruta, por ejemplo `POST /api/documentos/{id}/ops`. No contienen reglas de negocio: reciben los "
        "datos, obtienen la cuenta que hizo la petición y llaman al servicio.",
        "**Modelo**: las clases del paquete `core.modelo` (`Usuario`, `Espacio`, `Documento`, `Version`, `Sesion`, "
        "`Permiso`) y los servicios de `core.servicio`, que aplican las reglas: quién puede leer o editar, topes, "
        "versiones, papelera y la fusión de cambios (`MotorCRDT`, `Secuencia`).",
        "**Vista**: la representación JSON de la respuesta, que Spring genera con Jackson a partir de lo que devuelve el "
        "controlador. Los errores también salen como JSON, desde `ManejoErrores` (`@RestControllerAdvice`).",
    ])
    D.tabla("Controladores del servidor y su servicio", ["Controlador", "Rutas", "Servicio que usa"], [
        ["`CuentasApi`", "`/api/cuentas/…`: registro, sesión, recuperación, llaves, datos", "`ServicioCuentas`, `ServicioPases`"],
        ["`EspaciosApi`", "`/api/espacios/…`: espacios, accesos, enlaces, chat, papelera", "`ServicioEspacios`, `ServicioChat`"],
        ["`DocumentosApi`", "`/api/documentos/…`: abrir, cambios, versiones, ramas, recibo", "`ServicioDocumentos`"],
        ["`BinariosApi`", "imágenes y archivos de Office de cada espacio", "`ServicioBinarios`"],
        ["`PublicoApi`", "rutas sin sesión: enlaces de invitación, ver, proyector", "`ServicioEspacios`, `ServicioDocumentos`"],
        ["`FuentesApi`", "investigar y citar (Wikipedia, Crossref, Open Library)", "`ServicioFuentes`"],
        ["`RedApi`, `PaginasApi`", "información del servidor, páginas de ayuda y legales", "`Nodo`, `Buscadores`"]], [3.4, 8, 4.6])
    D.h2("2.3 MVC en el cliente")
    D.p("La página tiene su propio MVC. El **modelo** del cliente es la copia local de cada documento (`crdt.js` para "
        "los bloques y `secuencia.js` para el texto letra por letra) más la bóveda cifrada donde se guarda lo que "
        "todavía no se envía (`boveda.js`, IndexedDB). Las **vistas** son los módulos que dibujan cada pantalla: "
        "`vista-cuentas.js`, `vista-espacios.js`, `vista-dispositivo.js` y los cuatro editores (`modo-texto.js`, "
        "`modo-hoja.js`, `lienzo.js` y `diapositivas.js`). El **controlador** es `app.js`, que lee la dirección "
        "(`#/espacio/…`, `#/doc/…`) y decide qué vista montar, junto con los manejadores de eventos de cada editor y "
        "`api.js`, que habla con el servidor.")
    D.h2("2.4 Recorrido de una acción")
    D.p("La figura 2 sigue una letra desde que se escribe hasta que la ven las demás personas. La vista avisa al "
        "controlador del cliente, que crea una operación en el modelo local y la pone en la cola. La cola la manda al "
        "servidor; ahí los filtros comprueban la sesión, el controlador la pasa al servicio, el servicio revisa el "
        "permiso, la fusiona y la guarda por medio del DAO, y el difusor la reparte a los demás por el canal en vivo. "
        "Solo cuando llega el acuse del servidor la operación sale de la cola.")
    D.figura("peticion.png", "Recorrido de un cambio a través de las capas", 15.5)
    D.h2("2.5 Lo que se gana")
    D.lista([
        "El mismo modelo y los mismos servicios corren en la nube, en la computadora del salón y en el teléfono: solo "
        "cambia la capa de controladores (Spring en la computadora, `ApiTelefono` en Android).",
        "Las reglas se prueban sin servidor ni navegador (pruebas unitarias del núcleo con JUnit).",
        "La interfaz puede cambiar por completo sin tocar las reglas de permisos ni de sincronización.",
    ])

    D.h1("3. El patrón de objetos de acceso a datos (DAO)")
    D.h2("3.1 Qué es")
    D.p("Un DAO (*Data Access Object*) es un objeto que esconde cómo y dónde se guardan los datos. Las reglas de negocio "
        "piden y guardan objetos por medio de una interfaz, sin saber si detrás hay una base de datos, archivos o un "
        "servicio remoto; así se puede cambiar el almacenamiento sin tocar las reglas (Alur, Crupi y Malks, 2003).")
    D.h2("3.2 Los DAO de NOM")
    D.p("Las interfaces están en el paquete `core.almacen` del núcleo y se llaman *repositorios*. Los servicios solo "
        "conocen estas interfaces; cada plataforma da su implementación.")
    D.tabla("Interfaces DAO y sus implementaciones", ["Interfaz (núcleo)", "Qué guarda", "En el servidor", "En el teléfono"], [
        ["`RepositorioUsuarios`", "cuentas, sesiones y preferencias", "`RepositorioUsuariosJdbc` (H2 por JDBC)", "`Repositorios.Usuarios` (JSON)"],
        ["`RepositorioEspacios`", "espacios, accesos, enlaces, chat, biblioteca", "`RepositorioEspaciosJson`", "`Repositorios.Espacios`"],
        ["`RepositorioDocumentos`", "estado de cada documento, versiones y bitácora", "`RepositorioDocumentosJson`", "`Repositorios.Documentos`"],
        ["`RepositorioBinarios`", "imágenes y originales de Office por espacio", "`RepositorioBinariosArchivos`", "`RepositorioBinariosTelefono`"]], [3.6, 4.4, 4.4, 3.6])
    D.h2("3.3 La base de datos")
    D.p("Las cuentas viven en **H2**, una base de datos relacional que corre dentro del mismo proceso de Java y guarda "
        "sus datos en un archivo (no hay que instalar nada). Tiene tres tablas: `usuario`, `preferencias` y `sesion`, "
        "creadas por `schema-h2.sql` al arrancar (cada sentencia es idempotente, así que actualizar la versión no "
        "borra nada). El acceso es por JDBC con `JdbcClient` de Spring y un grupo de conexiones HikariCP de hasta cinco "
        "conexiones. Todas las consultas usan parámetros (`WHERE id = ?`): ningún dato del usuario se pega dentro del "
        "texto de una consulta. El archivo de la base va cifrado con una contraseña derivada de la llave del servidor.")
    D.h2("3.4 Documentos y archivos")
    D.p("Los documentos, los espacios y la bitácora se guardan como archivos JSON por medio de `AlmacenJson`. Cada "
        "escritura es atómica (se escribe un archivo nuevo y se reemplaza el anterior), se conservan tres copias "
        "anteriores de cada archivo y todo va cifrado en reposo con AES-256-GCM. Ninguna ruta puede salir de la carpeta "
        "de datos. Los binarios se guardan una sola vez por espacio con su huella SHA-256 como nombre, lo que evita "
        "duplicados y permite comprobar que llegaron completos.")
    D.p("Se eligió este reparto porque el estado de un documento colaborativo es un conjunto de operaciones que se "
        "fusionan, no filas que se actualizan; guardarlo como JSON permite que el teléfono, que no tiene Spring ni "
        "base de datos, use exactamente el mismo formato y las mismas reglas.")

    D.h1("4. Spring Boot y las tecnologías del proyecto")
    D.h2("4.1 Tecnologías y versiones")
    D.tabla("Tecnologías sobre las que corre NOM", ["Parte", "Tecnología", "Versión"], [
        ["Lenguaje del servidor y del núcleo", "Java (se compila para 17; en la nube corre OpenJDK 21)", "17 / 21"],
        ["Marco de trabajo del servidor", "Spring Boot (Spring Framework 7)", "4.1.3"],
        ["Servidor web incluido", "Apache Tomcat (Jakarta Servlet 6.1)", "11.0.29"],
        ["JSON", "Jackson", "3.1.8"],
        ["Base de datos", "H2 (archivo, dentro del proceso) con HikariCP", "2.4.240 / 7.0.2"],
        ["Bitácora", "Logback (SLF4J)", "1.5.38"],
        ["Interfaz", "HTML, CSS y JavaScript sin marcos de trabajo", "67 módulos"],
        ["App", "Android (WebView), mínimo Android 7, objetivo Android 17", "API 24 / 37"],
        ["Bibliotecas de la página", "qrcode-generator, jsQR, Mediabunny", "1.4.4, 1.4.0, 1.65.1"],
        ["Construcción", "Maven (servidor) y Gradle (app)", "las del proyecto"],
        ["Nube", "Raspberry Pi 5 con Alpine Linux, servicios s6 y Cloudflare Tunnel", "Alpine 3.24"]], [5, 8, 3])
    D.h2("4.2 Qué aporta Spring Boot")
    D.p("Spring Boot arma una aplicación de Spring con una configuración por omisión razonable, de modo que el programa "
        "completo se entrega como un solo archivo `nom.jar` que incluye su propio servidor web (Tomcat). En NOM aporta:")
    D.lista([
        "**Inversión de control e inyección de dependencias**: la clase `Configuracion` declara con `@Bean` los "
        "servicios del núcleo y Spring entrega a cada controlador los que necesita en su constructor.",
        "**Spring MVC**: el `DispatcherServlet`, los controladores `@RestController` y la conversión a JSON.",
        "**Filtros e interceptores**: cabeceras de seguridad, CORS, tope de tamaño de lo que se recibe y la comprobación "
        "de la sesión antes de cada ruta protegida (tabla 5).",
        "**Tareas programadas** con `@Scheduled` (tabla 6).",
        "**Canal en vivo** con `SseEmitter` (eventos enviados por el servidor).",
        "**Configuración externa**: `application.properties` lee variables de entorno (`NOM_PUERTO`, `NOM_DATOS`, "
        "`NOM_CUOTA_MB`…), así el mismo `jar` sirve en la nube y en el salón.",
    ])
    D.tabla("Lo que atraviesa cada petición, en orden", ["Componente", "Tipo", "Qué hace"], [
        ["`FiltroCabeceras`", "filtro de servlet", "cabeceras de seguridad (política de contenido, HSTS, aislamiento)"],
        ["`FiltroCors`", "filtro de servlet", "qué orígenes pueden usar la API (la app y la red local)"],
        ["`FiltroTamano`", "filtro de servlet", "tope de lo que se recibe y turnos según la memoria"],
        ["`FiltroSesion`", "interceptor de Spring MVC", "token de la cabecera `Authorization`; sin sesión válida, 401"],
        ["Controlador `*Api`", "`@RestController`", "atiende la ruta y llama al servicio"],
        ["`ManejoErrores`", "`@RestControllerAdvice`", "convierte los errores en JSON con su código HTTP"]], [3.6, 3.8, 8.6])
    D.tabla("Tareas programadas del servidor", ["Tarea", "Cada cuánto", "Para qué"], [
        ["Guardar documentos", "0.8 s", "escribe a disco los documentos que cambiaron"],
        ["Bitácora", "10 s", "registra la actividad de cada documento"],
        ["Latido del canal en vivo", "20 s", "mantiene las conexiones y cierra las de sesiones terminadas o atoradas"],
        ["Autoguardado de versiones", "1 min", "crea versiones según la preferencia de cada cuenta"],
        ["Copias entre servidores", "5 min", "sincroniza los espacios conectados con otro servidor (en su propio hilo)"],
        ["Limpieza de binarios", "6 h", "borra los archivos que ningún documento usa desde hace 30 días"]], [4, 2.6, 9.4])
    D.h2("4.3 El servidor del teléfono")
    D.p("Spring Boot es demasiado grande para un teléfono, así que en Android el núcleo se expone con `ServidorHttp`, un "
        "servidor HTTP propio de unas 460 líneas, y `ApiTelefono`, que repite las rutas de los controladores de Spring "
        "llamando a los mismos servicios. Es la misma arquitectura con otra capa de controladores, y por eso cada ruta "
        "nueva del servidor se agrega también ahí.")
    D.h2("4.4 Dónde corre")
    D.p("En la nube, el `jar` corre en una Raspberry Pi 5 supervisado por s6 (si deja de contestar, un vigía lo "
        "reinicia), escucha solo en la propia máquina y se publica por un túnel de Cloudflare, que da el HTTPS. En el "
        "salón, la misma aplicación corre en cualquier computadora con Java. En la app de Android, la interfaz viene "
        "dentro del APK y se conecta a cualquiera de los servidores anteriores o al teléfono que comparte.")

    D.h1("5. Calidad")
    D.p("Cada capa tiene sus pruebas: 33 pruebas unitarias del núcleo (JUnit), 454 casos de la batería de seguridad y "
        "permisos contra el servidor, 204 de funciones de la interfaz, y pruebas en el sitio real de conexión, "
        "pestañas, pantallas y archivos. Los archivos de Office que genera NOM se comparan con Microsoft Office 2024.")

    D.h1("6. Conclusiones")
    D.p("NOM aplica MVC en el servidor, con Spring MVC, y en la página, con un enrutador propio, y separa el acceso a "
        "datos con interfaces DAO que tienen una implementación para la computadora y otra para el teléfono. Gracias a "
        "esa separación el mismo núcleo da servicio en tres lugares distintos y las reglas de permisos y sincronización "
        "se escriben y se prueban una sola vez.")
    D.referencias([
        "Alur, D., Crupi, J. y Malks, D. (2003). *Core J2EE patterns: Best practices and design strategies* (2.ª ed.). Prentice Hall.",
        "Gamma, E., Helm, R., Johnson, R. y Vlissides, J. (1994). *Design patterns: Elements of reusable object-oriented software*. Addison-Wesley.",
        "Krasner, G. E. y Pope, S. T. (1988). A cookbook for using the model-view controller user interface paradigm in Smalltalk-80. *Journal of Object-Oriented Programming, 1*(3), 26-49.",
        "VMware. (2026). *Spring Boot reference documentation* (versión 4.1). https://docs.spring.io/spring-boot/",
        "VMware. (2026). *Spring Framework reference documentation: Web on Servlet stack*. https://docs.spring.io/spring-framework/reference/web/webmvc.html",
    ])
    return D.guardar("5IM7_WACK_Arquitectura")


# ================================================================ documento 2: sistemas distribuidos
def distribuidos():
    D = Doc("NOM como sistema distribuido", "SD-NOM-01", "Arquitectura, comunicación, replicación, consistencia y tolerancia a fallas")
    D.h1("1. Introducción")
    D.p("Este documento explica las decisiones de NOM como sistema distribuido: qué computadoras participan, cómo se "
        "comunican, cómo se mantiene igual un documento que varias personas editan a la vez desde equipos distintos y "
        "qué pasa cuando la red falla. Se toma como base la definición de Coulouris, Dollimore, Kindberg y Blair (2012): "
        "componentes en computadoras conectadas en red que se comunican y coordinan solo por medio de mensajes, con "
        "tres consecuencias que todo sistema distribuido debe resolver: concurrencia, ausencia de un reloj global y "
        "fallas independientes.")

    D.h1("2. Nodos y arquitectura")
    D.p("NOM sigue el modelo cliente-servidor, con una diferencia importante: cada cliente guarda una réplica completa "
        "de los documentos que abre y puede seguir trabajando sin el servidor. Hay tres formas de conectarse, que usan "
        "los mismos nodos en distintas combinaciones (figura 1).")
    D.figura("despliegue.png", "Nodos de NOM y cómo se conectan", 15)
    D.tabla("Nodos y su papel", ["Nodo", "Papel", "Cómo se conecta"], [
        ["Navegador (computadora o celular)", "cliente con réplica local", "HTTPS a la nube o HTTP en la red local"],
        ["App de Android", "cliente con réplica local; también puede ser servidor", "HTTPS o HTTP local"],
        ["Raspberry Pi (nube)", "servidor con la copia de referencia de cada espacio", "detrás de un túnel de Cloudflare"],
        ["Computadora del salón", "servidor en una red sin Internet", "HTTP en el Wi-Fi; se anuncia por UDP"],
        ["Teléfono que comparte", "servidor sin Wi-Fi externo, en su punto de acceso", "HTTP en su propia red"]], [4.4, 5.6, 6])

    D.h1("3. Comunicación")
    D.lista([
        "**Peticiones HTTP con JSON** (120 rutas en el servidor de la computadora y 90 en el del teléfono). Cada cambio "
        "a un documento es una operación pequeña; se mandan en lotes de hasta 300 y el servidor contesta un acuse con "
        "cuántas aplicó.",
        "**Eventos enviados por el servidor** (SSE): una conexión abierta por cada pestaña por la que el servidor empuja "
        "los cambios de los demás, quién está conectado y los avisos. Un latido cada 20 s mantiene la conexión a través "
        "de los intermediarios. Como el navegador no puede mandar cabeceras en esta conexión, se identifica con un "
        "boleto de un solo uso que dura 30 s.",
        "**Anuncios UDP** en el puerto 47474: la computadora del salón anuncia sus espacios para que los demás equipos "
        "los encuentren sin escribir direcciones.",
        "**Códigos QR y enlaces** con la dirección del servidor y un código de invitación que caduca.",
    ])
    D.p("En la nube, Cloudflare termina el HTTPS y usa HTTP/2 o HTTP/3, que mezcla muchas peticiones en una sola "
        "conexión. En la red local se usa HTTP/1.1, donde el navegador abre como máximo seis conexiones por servidor "
        "entre todas sus pestañas; como cada canal en vivo ocupa una, una pestaña que lleva más de un minuto oculta "
        "pausa su canal y, al volver, se pone al día. Así varias pestañas abiertas no dejan sin conexión a las demás.")

    D.h1("4. Tiempo y orden de los eventos")
    D.p("Los relojes de los teléfonos y computadoras no coinciden, así que NOM no ordena los cambios por la hora. Cada "
        "operación lleva un reloj lógico de Lamport (1978) y el identificador del nodo que la generó; el par (reloj, "
        "nodo) da un orden total que todos los nodos calculan igual. La hora del servidor solo se usa para lo que es "
        "de él: el vencimiento de sesiones, de códigos de invitación y de la papelera.")

    D.h1("5. Replicación y consistencia")
    D.p("Cada documento está replicado en el servidor y en cada equipo que lo tiene abierto. Para que las réplicas "
        "lleguen al mismo resultado sin bloquear a nadie, NOM usa tipos de datos replicados libres de conflictos (CRDT; "
        "Shapiro, Preguiça, Baquero y Zawirski, 2011):")
    D.lista([
        "Cada **campo de cada bloque** (una celda, una figura, una diapositiva) es un registro en el que gana la "
        "escritura con el par (reloj, nodo) mayor. El resultado no depende del orden de llegada.",
        "El **texto** es una secuencia de tipo RGA: cada letra tiene un identificador único, así que dos personas pueden "
        "escribir en el mismo párrafo a la vez sin perder ninguna letra.",
        "Las operaciones son **idempotentes**: cada una tiene un identificador `nodo:contador` y el servidor descarta las "
        "repetidas, así que reenviar un lote después de un corte no duplica nada.",
    ])
    D.p("El modelo de consistencia es de consistencia eventual fuerte: dos réplicas que recibieron las mismas "
        "operaciones tienen el mismo contenido, aunque las hayan recibido en otro orden. Mientras tanto cada persona ve "
        "sus propios cambios al instante.")

    D.h1("6. Tolerancia a fallas")
    D.figura("sinred.png", "Edición sin conexión y vuelta a la red", 13.5)
    D.tabla("Fallas y cómo responde NOM", ["Falla", "Respuesta"], [
        ["Se corta la red de un equipo", "sigue editando; los cambios esperan en una cola cifrada en el equipo y se mandan en orden al volver"],
        ["Otros editaron mientras tanto (más de 2 min sin red)", "lo editado sin conexión va a una rama para compararla y fusionarla"],
        ["Se cierra el navegador antes de volver la red", "la cola está en IndexedDB; se manda la siguiente vez; se pide al navegador que no la borre"],
        ["El Wi-Fi contesta con su página de inicio de sesión", "solo un acuse de NOM saca los cambios de la cola; lo demás cuenta como sin red"],
        ["Un equipo deja de leer su canal en vivo", "cada canal tiene su propia fila; el servidor lo corta sin detener a los demás"],
        ["El servidor se reinicia", "los equipos se reconectan con esperas al azar y recargan el estado; el servidor guarda cada 0.8 s"],
        ["El servidor deja de contestar", "un vigía lo reinicia; una actualización que no contesta regresa sola a la versión anterior"],
        ["Se daña el almacenamiento", "tres copias por archivo y respaldo diario en la Pi (14 días) y en otra computadora (30)"]], [5.6, 10.4])

    D.h1("7. Copias entre servidores")
    D.p("Un espacio que se trabajó en el salón sin Internet puede conectarse con la nube: cada cinco minutos, si hay red, "
        "`ServicioSincronia` intercambia las operaciones que le faltan a cada lado (las mismas reglas del CRDT hacen "
        "que el resultado sea igual en los dos servidores). Para que una persona use su cuenta de la nube en el "
        "teléfono que comparte, la nube le da un **pase firmado** con ECDSA P-256 que el teléfono comprueba con la "
        "llave pública que trae la app, sin necesidad de Internet.")

    D.h1("8. Concurrencia y escalabilidad")
    D.p("Dentro del servidor cada documento tiene su propio candado, así que dos documentos se editan en paralelo; los "
        "envíos del canal en vivo salen de un grupo de hilos aparte, de modo que un equipo lento no detiene a quien "
        "escribe. Hay topes explícitos: 50 personas conectadas a la vez por espacio, 40 canales por cuenta y 5000 en "
        "todo el servidor; lo que se recibe a la vez se limita según la memoria. Medido en la Raspberry Pi: más de "
        "10 000 respuestas por segundo a una petición pequeña y unas 800 cargas de la página por segundo; con 100 "
        "personas escribiendo a la vez en un documento no se perdió ningún cambio.")

    D.h1("9. Seguridad en un sistema distribuido")
    D.lista([
        "Por Internet todo viaja por HTTPS (TLS 1.2 o 1.3) hasta Cloudflare y por un túnel cifrado hasta la Pi.",
        "La sesión viaja en una cabecera (`Authorization`), no en cookies, y cada ruta revisa el permiso de la cuenta "
        "sobre el espacio que pide.",
        "Un espacio puede cifrarse de extremo a extremo: el contenido se cifra en cada equipo con AES-GCM y la llave del "
        "espacio va envuelta para cada miembro, así que ni el servidor puede leerlo.",
        "Los pases entre servidores van firmados y vencen a los 10 minutos.",
    ])

    D.h1("10. Teorema CAP y transparencias")
    D.p("Según el teorema CAP (Brewer, 2000), ante una partición de red un sistema debe elegir entre consistencia "
        "inmediata y disponibilidad. NOM elige **disponibilidad y tolerancia a particiones**: cada persona puede seguir "
        "trabajando sin red, y la consistencia se recupera después gracias al CRDT. La tabla 3 resume las "
        "transparencias que describen Tanenbaum y van Steen (2008).")
    D.tabla("Transparencias de NOM", ["Transparencia", "Cómo se logra"], [
        ["De acceso", "las mismas rutas y formatos en la nube, la computadora del salón y el teléfono"],
        ["De ubicación", "se entra con un nombre (www.nom.qd.je), un código QR o el anuncio de la red, sin saber dónde está el servidor"],
        ["De replicación", "cada equipo tiene su copia y la persona ve un solo documento"],
        ["De concurrencia", "varias personas escriben a la vez y el CRDT junta los cambios sin bloqueos"],
        ["De fallas", "los cortes de red se esconden con la cola, los reintentos y las ramas"]], [4, 12])

    D.h1("11. Limitaciones")
    D.lista([
        "La base de cuentas está en un solo nodo por servidor: no es una base de datos distribuida.",
        "En la red local la comunicación va sin cifrar (HTTP); la app lo avisa y los espacios cifrados de extremo a "
        "extremo siguen protegidos.",
        "Algunas redes de escuelas aíslan a los equipos entre sí; ahí se usa el teléfono que comparte su propio punto de acceso.",
    ])

    D.h1("12. Conclusiones")
    D.p("NOM resuelve los tres problemas propios de un sistema distribuido: la concurrencia con un CRDT que fusiona sin "
        "bloquear, la falta de un reloj común con relojes de Lamport, y las fallas independientes con réplicas en cada "
        "equipo, colas que sobreviven a los cortes y servidores que se recuperan solos. El resultado es un sistema que "
        "sigue disponible aunque la red se caiga y que deja a todas las personas con el mismo documento cuando vuelve.")
    D.referencias([
        "Brewer, E. A. (2000). Towards robust distributed systems. En *Proceedings of the 19th Annual ACM Symposium on Principles of Distributed Computing* (p. 7). ACM. https://doi.org/10.1145/343477.343502",
        "Coulouris, G., Dollimore, J., Kindberg, T. y Blair, G. (2012). *Distributed systems: Concepts and design* (5.ª ed.). Addison-Wesley.",
        "Lamport, L. (1978). Time, clocks, and the ordering of events in a distributed system. *Communications of the ACM, 21*(7), 558-565. https://doi.org/10.1145/359545.359563",
        "Shapiro, M., Preguiça, N., Baquero, C. y Zawirski, M. (2011). Conflict-free replicated data types. En X. Défago, F. Petit y V. Villain (Eds.), *Stabilization, safety, and security of distributed systems* (pp. 386-400). Springer. https://doi.org/10.1007/978-3-642-24550-3_29",
        "Tanenbaum, A. S. y van Steen, M. (2008). *Sistemas distribuidos: Principios y paradigmas* (2.ª ed.). Pearson Educación.",
    ])
    return D.guardar("5IM7_WACK_Sistemas_Distribuidos")


if __name__ == "__main__":
    subprocess.run(["plantuml", "-tpng", "-Sdpi=170"] + [os.path.join(AQUI, f) for f in os.listdir(AQUI) if f.endswith(".puml")], check=True)
    sys.path.insert(0, os.path.join(AQUI, "..", "..", "4-CODIGO", "pruebas", "office"))
    try:
        import con_word
        con = con_word.disponible()
    except Exception:
        con = False
    for hacer, titulo in ((arquitectura, "NOM: arquitectura del software"), (distribuidos, "NOM como sistema distribuido")):
        ruta = hacer()
        if con:
            info = con_word.guardar_y_pdf(ruta, autor="WACK", titulo=titulo)
            print(os.path.basename(ruta), "→ PDF con Word,", info.get("paginas"), "páginas")
        else:
            with tempfile.TemporaryDirectory() as t:
                subprocess.run(["soffice", f"-env:UserInstallation=file://{t}/perfil", "--headless", "--convert-to", "pdf",
                                "--outdir", os.path.dirname(ruta), ruta], check=True, capture_output=True)
            print("AVISO: sin la PC de Windows; PDF de LibreOffice para", os.path.basename(ruta))
