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

## 3. Cómo seguir

- **GitHub (manos del dueño):** borrar el repo vacío `TruTruMaster/nom`, renombrar `NOM-COLAB` a `NOM` y borrar el secreto
  `AZURE_WEBAPP_PUBLISH_PROFILE`. El token de `tokensito` no tiene permiso de administración y `gh` necesita
  `gh auth login`. Después: `git remote set-url origin https://github.com/TruTruMaster/NOM.git`.
- Publicar la 0.23.0 en este orden: Pi (`nom pi actualizar`, humo), `git push`, APK (`publicar-apk.sh`), buscadores.
- La regresión de Windows no se corrió (la PC no respondía); los PDF de casos de uso y del ensayo salieron de LibreOffice:
  revisarlos en Word.
