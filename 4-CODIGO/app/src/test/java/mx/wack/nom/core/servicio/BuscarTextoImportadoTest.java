package mx.wack.nom.core.servicio;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.LinkedHashMap;
import java.util.Map;

import org.junit.jupiter.api.Test;

import mx.wack.nom.core.crdt.EstadoDocumento;
import mx.wack.nom.core.crdt.Secuencia;

/** «Buscar en los documentos» encuentra lo que dice un Word importado (su texto vive en «cuerpo» sin campos propios). */
class BuscarTextoImportadoTest {

    private static Secuencia texto(String t) {
        Secuencia s = new Secuencia();
        Map<String, Object> i = new LinkedHashMap<>();
        i.put("k", "i"); i.put("iz", ""); i.put("x", t);
        s.aplicar(i, 1, "ana");
        return s;
    }

    @Test
    void textoSinBloqueSeBusca() {
        EstadoDocumento est = new EstadoDocumento();
        est.seq.put("cuerpo", texto("Medimos la temperatura del agua"));
        assertTrue(ServicioDocumentos.textoBuscable(est).contains("temperatura"));
    }

    @Test
    void textoDeUnBloqueBorradoNo() {
        EstadoDocumento est = new EstadoDocumento();
        est.seq.put("caja1", texto("borrador SECRETO"));
        Map<String, mx.wack.nom.core.crdt.Registro> campos = new LinkedHashMap<>();
        campos.put("eliminado", new mx.wack.nom.core.crdt.Registro(Boolean.TRUE, 5, "ana"));
        est.bloques.put("caja1", campos);
        assertFalse(ServicioDocumentos.textoBuscable(est).contains("SECRETO"));
    }
}
