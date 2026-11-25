# Relevo del 25 de noviembre de 2026

**Leer este primero.** Sigue a `RELEVO-23NOV.md`. Al empezar estaba publicada la **0.22.2**; esta sesión hizo la
**0.23.0** (interfaz 27, APK 42).

## 1. Hecho

| Qué | Dónde | Comprobado |
|---|---|---|
| Azure fuera de todo | Nube: grupo `rg-nomcolab` borrado (Web App, plan, servidor y base SQL). Repo: sin `.github/workflows/azure.yml`, sin `4-CODIGO/azure/`, sin `schema-sqlserver.sql` ni `mssql-jdbc`; `Origen.java` sin el caso de Azure; ayuda, aviso de privacidad (`actualizado` 25-nov, `version` sigue 13-nov), diagramas 02/13/17, docs, plan de pruebas. Se quitaron `5-DOCUMENTO-MODULO/` y su PDF | `mvn package` y pruebas en verde; seguridad 446/446 |
| Todo se llama NOM | Textos, nombres de archivo, el jar (`nom.jar`), la clase `NomApplication`, el APK (`NOM.apk`), scripts y servicios (`nom`, `nom-pi`, `nom-servidor`…), documento `PP-NOM-01` | Ver §2 para lo que **no** cambia a propósito |
| La Pi con los nombres nuevos | `servidor-propio/pi/migrar-nombre.sh` (una sola vez): usuario `nom` (mismo UID 106), `/var/lib/nom`, `/opt/nom/nom.jar`, `/etc/nom.conf`, `/etc/cloudflared/nom.yml`, servicios s6 `nom-*`, bitácoras `/var/log/nom-*`, orden `nom-pi` y su regla de doas, respaldo diario `nom-respaldo` | Contesta en `www.nom.qd.je` y en `nomcolab.dpdns.org`, mismo id de servidor, pases firmados |
| La laptop | `~/nom-servidor` (antes `~/nomcolab-servidor`), `~/.config/nom` (antes `~/.config/nomcolab`; `pi-doas` ya lee ahí), unidades `nom-respaldo.timer` y `nom-traer-respaldos.timer`, orden `nom` en `~/.local/bin` | `nom-traer-respaldos` trajo el respaldo de la Pi |
| Ensayo individual | `8-ENSAYO/fuentes/generar.py` → `ENTREGA/Ensayo_Rodriguez_Esparza*.docx/.pdf` | 4 páginas con portada / portada + 4 |

## 2. Lo que se queda con el nombre viejo, a propósito

- Cadenas de cifrado (`nom-colab/reposo/`, `nom-colab/e2e/envoltura`): con otro texto no se abren los datos.
- La marca del formato `.nom` (`"f":"nom-colab"`) y el diccionario de `origen.js`: los archivos y apps de antes.
- `nom-colab-equipo:` en Android: el identificador de cada teléfono.
- El dominio `nomcolab.dpdns.org` (y `origen.nomcolab.dpdns.org`): es una dirección real y la nube que conocen los APK
  para los pases.
- El túnel de Cloudflare se llama `nomcolab` en la cuenta de Cloudflare (`tunel.sh`).
- La release de GitHub lleva también una copia `NOM-COLAB.apk` para las apps y enlaces de antes.
- Bitácoras con fecha (`LOG.md`, `CAMBIOS.md` viejo, `RELEVO-*`, novedades de versiones pasadas): son historia.
- `4-CODIGO/entregables/Auditoria_Seguridad_NOM_5IM7_WACK.pdf` cambió de nombre, pero por dentro sigue diciendo
  NOM-COLAB (no hay fuente para regenerarlo).

## 3. Publicado (25-nov, 10:05)

- GitHub: se borró el repo vacío `TruTruMaster/nom`, `NOM-COLAB` se renombró a **`TruTruMaster/NOM`** (los enlaces
  viejos redirigen) y se borró el secreto `AZURE_WEBAPP_PUBLISH_PROFILE`. Token nuevo en `tokensito` (el `gh` de esta PC
  sigue sin sesión: falla la conexión, se usa la API con `curl`).
- Pi: 0.23.0 con `nom pi actualizar`; `humo.py` contra `www.nom.qd.je` 36 de 36. Commit `c5b6afb` en `main`.
- Release `v0.23.0` con `NOM.apk` y `NOM-COLAB.apk` (el mismo archivo); `/app`, el enlace nuevo y el viejo bajan el
  mismo APK. IndexNow avisado (76 direcciones). GitHub tenía «servicio degradado»: hubo que reintentar.

## 4. 0.23.1: nomcolab.dpdns.org oculto

- Ya no se muestra en ningún lado (README, ayuda, aviso de privacidad, novedades). El navegador que lo abre pasa a
  `www.nom.qd.je`: si no guarda nada de NOM, al momento; si guarda algo, por `mudanza.html` (`js/mudanza.js`), que con
  un clic abre la dirección nueva y le pasa todo por `postMessage` sin pisar lo que ya haya allá. Solo esa página
  lleva `Cross-Origin-Opener-Policy: unsafe-none`. Prueba: `pruebas/interfaz/prueba_mudanza.py` (Chrome y Firefox,
  12 de 12; está en la regresión como `mudanza`).
- Sigue como infraestructura (no se apaga ni se deja vencer): la zona de Cloudflare de la que cuelga `www.nom.qd.je`,
  la API para las apps ya instaladas y la firma de los pases (`NOM_NUBE`).
- `Cross-Origin-Opener-Policy` ya no se manda por http en la red local (el navegador la ignora y avisaba en la
  consola: fallaba `cuentas_equipo`); el teléfono que comparte no la manda nunca.

## 5. Cómo seguir

- Instalar la 0.23.0 en el A55 cuando esté en la misma red.
- La regresión de Windows no se corrió (la PC no respondía); los PDF de casos de uso y del ensayo salieron de LibreOffice:
  revisarlos en Word.
