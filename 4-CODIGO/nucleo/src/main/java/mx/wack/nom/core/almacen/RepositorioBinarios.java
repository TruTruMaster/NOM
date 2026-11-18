package mx.wack.nom.core.almacen;

import java.util.List;

/**
 * Almacén de binarios de cada espacio: imágenes, el archivo de Office original de un documento importado, y
 * (después) video y audio. Cada binario se guarda una sola vez por espacio con el nombre de su huella (SHA-256 de
 * los bytes que llegan; en un espacio cifrado de extremo a extremo, de los bytes YA cifrados en el equipo), así que
 * un binario nunca cambia: lo que cambia es qué documentos lo usan.
 */
public interface RepositorioBinarios {

    boolean existe(String espacioId, String huella);

    /** Guarda los bytes (si ya existían, no hace nada). */
    void guardar(String espacioId, String huella, byte[] datos);

    /** Los bytes, o null si no está. */
    byte[] leer(String espacioId, String huella);

    /** Lo que ocupa (en disco) todo lo del espacio. */
    long bytesDe(String espacioId);

    /** Las huellas que tiene el espacio. */
    List<String> huellas(String espacioId);

    void borrar(String espacioId, String huella);

    /** Milisegundos desde que se guardó (para no borrar lo recién subido que todavía no usa ningún documento). */
    long edadMs(String espacioId, String huella);
}
