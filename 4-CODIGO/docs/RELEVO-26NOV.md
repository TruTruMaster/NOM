# Relevo del 26 de noviembre de 2026

**Leer este primero.** Sigue a `RELEVO-25NOV-NOCHE.md`. Al empezar estaba publicada la **0.23.3** (interfaz 27, APK 45).
El dueño pidió revisar todo lo de red «como si no se conociera el código» y después toda la app, investigando también
técnicas de ataque recientes. Esta sesión hizo la **0.23.4** (interfaz 27, APK 46).

El detalle de seguridad (qué se encontró, cómo se reproduce y cómo se probó) vive solo en el plan privado
`~/NOM/Auditoria-Seguridad-PLAN.md` (§N y §N.2) y en `~/NOM/trabajo/auditoria-26nov/` (catálogo de ataques y mapa de
cobertura por archivo). Aquí va lo que cambia para quien usa NOM y cómo seguir.

## 1. Hecho (0.23.4)

| Qué | Dónde |
|---|---|
| Mover un documento a otro espacio es del propietario del espacio (quien edita, copia) | `ServicioDocumentos.mover`, `vista-espacios.js`, `editor.js` |
| Presencia (cursor de los demás) solo de los canales propios y con datos acotados | `Difusor`, `DocumentosApi`, `ApiTelefono` |
| Reloj de los cambios acotado por envío (operaciones ordenadas al subir y al fusionar ramas) | `ServicioDocumentos.limiteReloj`, `ServicioSincronia.operaciones` |
| Metadatos acotados (llave de documento cifrado, base y nodos de ramas, copia en otro servidor, enlaces guardados) | `ServicioDocumentos`, `ServicioSincronia`, `ServicioEspacios` |
| Bitácora de cada documento: se lee desde el final y no pasa de 16 MB; cuenta en la cuota | `AlmacenJson.ultimasLineas/recortarLineas`, `RepositorioDocumentosJson` |
| Versión manual: con cuota y sin repetir si no hubo cambios; etiqueta ≤ 120 | `guardarVersion` |
| Copia con otro servidor: cuota y tope por documento al traer; binarios con cuota | `fusionarExterno`, `ServicioSincronia.bajarBinarios` |
| Topes por cuenta: chat (30/min), búsqueda (40/min, 2 a la vez en el servidor), «copiar ahora» (6/min), cambios (1000 por envío y tiempo de procesador por minuto), presencia | `Topes` (nuevo), servicios |
| Direcciones y nombres de servidor que se atienden (400/421) | `FiltroRutas` (nuevo), `Hosts` (nuevo, también en el teléfono) |
| Topes por red con IPv6 agrupada por /64 | `Origen.agrupar` |
| Página: rutas de la API armadas con datos se validan; ids de la dirección validados; política Trusted Types «default» en la página principal | `api.js`, `app.js`, `nucleo.js`, CSP de `FiltroCabeceras` / `index.html` / teléfono |
| Enlaces de documentos: solo web, correo y marcadores; exportar a HTML/Markdown limpio | `modo-texto.js`, `formatos.js` |
| Archivos de Office descargados: relaciones externas que no son web/correo/unidad local pasan a `about:blank`; nivel de lista numérico | `ooxml.js` (`relsSeguras`, `nivelLista`) |
| Colores y anchos de figuras validados | `lienzo.js` (`colorSeguro`) |
| Fórmulas: presupuesto por cálculo (#NUM!) | `formulas.js` |
| Teléfono que comparte: plazos para peticiones lentas, sesión en la dirección solo en el canal en vivo | `ServidorHttp`, `ApiTelefono` |
| Descubrimiento UDP acotado; en la nube no se listan espacios | `Descubrimiento`, `RedApi` |
| App de Android: nombres de archivo limpios al guardar; enlaces solo de los dominios del manifiesto | `MainActivity` |

Pruebas: JUnit 41/41 (8 nuevas: `ProteccionesTest`, `FiltroRutasTest`, `OrigenTest`, `BitacoraTest`). En vivo
(www.nom.qd.je, 0.23.4): `humo.py` 36/36, los cuatro nombres del túnel responden, las direcciones raras dan 400 y la
página trae Trusted Types. Resultados de la interfaz: ver §2.

## 2. Publicación

- Pi: 0.23.4 publicada (respaldo previo `datos-2026-11-26_091000`). Ver §3 para lo que falta.

## 3. Cómo seguir / lo que falta

1. Terminar la revisión archivo por archivo (`trabajo/auditoria-26nov/COBERTURA.md`): quedan las vistas grandes de la
   página (lienzo, hoja, animaciones, importadores de Office) con el método del catálogo.
2. Después del 27-nov: el orden del texto en el servidor con un árbol (para documentos con partes asignadas muy
   grandes), subir H2 a 2.5.x con prueba de migración, versionar la envoltura de llaves para subir sus iteraciones.
3. Lo del dueño: ver el plan privado (dominio de antes, correo, DMARC, copia de la llave del cifrado en reposo).
