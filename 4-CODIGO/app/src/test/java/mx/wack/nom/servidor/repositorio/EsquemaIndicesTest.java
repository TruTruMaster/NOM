package mx.wack.nom.servidor.repositorio;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

import org.junit.jupiter.api.Test;

/** Las consultas de cuentas y sesiones usan un índice (no recorren toda la tabla), con el esquema de verdad. */
class EsquemaIndicesTest {

    private static String plan(Statement st, String sql) throws Exception {
        try (ResultSet r = st.executeQuery("EXPLAIN " + sql)) { r.next(); return r.getString(1); }
    }

    @Test
    void lasConsultasPrincipalesUsanIndice() throws Exception {
        try (Connection c = DriverManager.getConnection("jdbc:h2:mem:indices;DB_CLOSE_DELAY=-1"); Statement st = c.createStatement()) {
            st.execute("RUNSCRIPT FROM 'classpath:/schema-h2.sql'");
            String[][] casos = {
                { "SELECT * FROM usuario WHERE id = 'x'", "usuario" },
                { "SELECT * FROM usuario WHERE nombre_clave = 'x'", "usuario" },
                { "SELECT * FROM sesion WHERE token = 'x'", "sesion" },
                { "SELECT * FROM sesion WHERE id_usuario = 'x'", "sesion" },
                { "SELECT * FROM sesion WHERE expira <= 5", "sesion" },          // el borrado de sesiones vencidas
                { "SELECT * FROM preferencias WHERE id_usuario = 'x'", "preferencias" },
            };
            for (String[] k : casos) {
                String p = plan(st, k[0]).toUpperCase();
                assertTrue(!p.contains("TABLESCAN") && !p.contains("_SCAN_"), k[0] + " recorre toda la tabla: " + p);
            }
        }
    }
}
