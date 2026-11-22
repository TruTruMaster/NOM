package mx.wack.nom.core;

/**
 * Version de NOM que corre en este servidor (la computadora, la Raspberry o un telefono que
 * hace de servidor). La informa /api/red/info.
 *
 * La pagina siempre trae la interfaz de su servidor, asi que nunca se desfasa. La app Android trae
 * la suya empaquetada: si se conecta a un servidor de otra version lo avisa (app.js), porque lo que
 * una importa de Office la otra puede verlo distinto. Al cambiar el formato de los documentos o la
 * interfaz hay que subir INTERFAZ aqui y en js/nucleo.js (NOM.INTERFAZ), y publicar servidor y app
 * juntos.
 */
public final class Version {
    /** La misma que versionName del APK. */
    public static final String APP = "0.22.1";
    /** Nivel de la interfaz y del formato de documentos. 14 = espacios cifrados de extremo a extremo (letras con marca cx);
        15 = paginación nueva de Word (sin hojas en blanco de más) y Office con rutas irregulares (0.14.2);
        21 = campos nuevos de la 0.16.0 (transiciones y animaciones de PowerPoint: transOffice y sus árboles, sonidos,
        acciones al hacer clic; texto oculto y elevado de Word);
        22 = almacén de binarios de la 0.17.0 (imágenes y el Office original como «nom:bin/huella»: una página vieja no
        los sabe dibujar ni descargar);
        23 = herramientas de Excel de la 0.18.0 (nombres definidos, esquema, cálculo manual, tablas de datos y mostrar
        fórmulas que guarda el importador de Excel v10: una página vieja no los conoce);
        24 = guiones y números de línea de Word de la 0.19.0 (el idioma de cada párrafo, «numLineas» del documento y
        «suprimir números de línea» del párrafo: una página vieja no los conoce);
        25 = campos nuevos de la 0.20.0 (escenarios de Excel —importador v11—, cambios marcados de Word como sugerencias y
        «controlCambios», secciones, diapositivas ocultas y presentaciones personalizadas de PowerPoint);
        26 = campos nuevos de la 0.21.0 (estilos de párrafo del documento «estilosP», minigráficos de Excel «mg» —importador
        v12—, tabulaciones predeterminadas «tabDef» puestas desde NOM).
        27 = campos nuevos de la 0.22.0 (tablas con formato «lo» y tablas dinámicas «pv» de las hojas, el tema del libro en
        «xl.tema» —importador 13—, tabla de ilustraciones e índice alfabético «modo» del bloque índice). */
    public static final int INTERFAZ = 27;
    /** Fase 14 (F.5): version vigente de los Terminos y del Aviso de privacidad. Al cambiarla, la app
        pide a cada cuenta leerlos y aceptarlos otra vez. */
    public static final String LEGAL = "2026-11-13";

    private Version() { }
}
