---
titulo: Contraseñas y protección de Office
descripcion: Abre archivos de Office con contraseña, descárgalos con contraseña y protege como en Word, Excel y PowerPoint.
orden: 15
aplica: Navegador, App Android, Teléfono que comparte
actualizado: 2026-11-17
version: 0.16.0
---
## Abrir un archivo con contraseña {#abrir}

Si el archivo de Word, Excel o PowerPoint tiene **contraseña para abrir**, al importarlo NOM te la
pide. El archivo se descifra **en tu equipo**: la contraseña no se envía a ningún servidor. Sirve el
cifrado de Office 2010 en adelante y el de Office 2007.

Si la contraseña no es la correcta, NOM lo dice y puedes volver a intentarlo.

## Descargar con contraseña {#descargar}

En **… → Descargar**, elige Word, Excel o PowerPoint y abre **Proteger con contraseña**. El archivo
sale cifrado como lo cifra Office (AES de 256 bits) y Office lo abre pidiendo esa contraseña.

!> Si olvidas la contraseña, **nadie puede abrir el archivo**: ni NOM ni Office la pueden recuperar.

## Proteger como en Office {#proteger}

En **… → Proteger (como en Office)** están las mismas protecciones que en Office. Las que ya trae un
archivo se respetan al abrirlo y se conservan al descargarlo.

| Programa | Protección | Qué hace |
|---|---|---|
| Word | **Restringir edición** | Solo lectura, solo comentarios, solo cambios con control de cambios o solo rellenar formularios. Con o sin contraseña para quitarla |
| Word, Excel y PowerPoint | **Marcar como final** | Avisa que el documento está terminado y lo abre en solo lectura; cualquiera puede seguir editando si lo decide |
| Word, Excel y PowerPoint | **Contraseña para modificar** | Sin la contraseña, el documento se abre en solo lectura |
| Excel | **Proteger hoja** | Las celdas bloqueadas no se cambian; las fórmulas ocultas no se ven en la barra. Puedes permitir cosas como dar formato, insertar filas u ordenar |
| Excel | **Proteger libro** | No se pueden agregar, quitar, renombrar ni mover hojas |

En una hoja de Excel, para dejar celdas libres en una hoja protegida: selecciónalas y usa
**Desbloquear celdas** en su menú (antes de proteger la hoja). **Ocultar la fórmula** hace que la
celda muestre su resultado pero no su fórmula.

!> Estas protecciones son como las de Office: evitan cambios por descuido, **no son cifrado**. Quien
tenga el archivo puede leer su contenido. Para que nadie más lo lea, usa la contraseña para abrir.

## En un espacio compartido {#espacios}

La protección viaja con el documento: todos los miembros la ven. Los permisos del espacio
(*Puede editar*, *Solo lectura*) siguen mandando: alguien con *Solo lectura* no puede editar aunque
conozca la contraseña de la protección.
