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

    /**
     * Los espacios conectados con otro servidor se copian cada 5 minutos (si hay red), en su propio hilo: lo demás
     * de esta lista (guardar los documentos, el latido del canal en vivo) no espera a otro servidor.
     */
    private final java.util.concurrent.ExecutorService hiloCopias = java.util.concurrent.Executors.newSingleThreadExecutor(r -> {
        Thread t = new Thread(r, "nom-copias");
        t.setDaemon(true);
        return t;
    });
    private final java.util.concurrent.atomic.AtomicBoolean copiando = new java.util.concurrent.atomic.AtomicBoolean();

    @Scheduled(fixedDelay = 300_000, initialDelay = 90_000)
    public void sincronizar() {
        if (!copiando.compareAndSet(false, true)) return;
        hiloCopias.execute(() -> {
            try { sincronia.cadaRato(); } finally { copiando.set(false); }
        });
    }

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
