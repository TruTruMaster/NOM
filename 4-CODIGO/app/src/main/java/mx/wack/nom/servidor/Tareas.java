package mx.wack.nom.servidor;

import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import jakarta.annotation.PreDestroy;
import mx.wack.nom.core.servicio.Difusor;
import mx.wack.nom.core.servicio.ServicioBinarios;
import mx.wack.nom.core.servicio.ServicioDocumentos;
import mx.wack.nom.core.servicio.ServicioSincronia;

/** Lo que corre solo: al arrancar, cada cierto tiempo y al apagar (en el teléfono lo hace la app). */
@Component
public class Tareas {
    private final ServicioDocumentos documentos;
    private final Difusor difusor;
    private final ServicioSincronia sincronia;
    private final ServicioBinarios binarios;

    public Tareas(ServicioDocumentos documentos, Difusor difusor, ServicioSincronia sincronia, ServicioBinarios binarios) {
        this.documentos = documentos;
        this.difusor = difusor;
        this.sincronia = sincronia;
        this.binarios = binarios;
    }

    /** El almacén de binarios: lo que ya ningún documento usa (ni sus versiones ni la papelera) desde hace 30 días se
        borra, cada 6 horas (la primera, 10 minutos después de arrancar). */
    @Scheduled(fixedDelay = 6 * 3_600_000, initialDelay = 600_000)
    public void limpiarBinarios() { binarios.limpiarTodo(); }

    /** Los espacios conectados con otro servidor se copian cada 5 minutos (si hay red). */
    @Scheduled(fixedDelay = 300_000, initialDelay = 90_000)
    public void sincronizar() { sincronia.cadaRato(); }

    @Scheduled(fixedDelay = 800)
    public void persistir() { documentos.persistir(); }

    @Scheduled(fixedDelay = 10_000)
    public void bitacora() { documentos.escribirBitacora(); }

    @Scheduled(fixedRate = 60_000, initialDelay = 60_000)
    public void autoguardar() { documentos.autoguardar(); }

    @Scheduled(fixedRate = 20_000)
    public void latido() { difusor.latido(); }

    @PreDestroy
    public void alApagar() { documentos.alApagar(); }
}
