# Relevo del 23 de noviembre de 2026

**Leer este primero.** Sigue a `RELEVO-22NOV.md`. Al empezar estaba publicada la **0.22.2** (interfaz 27, APK 41).
Esta sesión fue de comprobación del dominio y ajustes de cabeceras: no cambió la paridad con Office ni los entregables.

## 1. Hecho

| Qué | Dónde | Comprobado |
|---|---|---|
| Dominio raíz y TLS | — | `nom.qd.je` ya da 301 a `www` con certificado válido; `www` sigue con certificado de Cloudflare y solo TLS 1.2/1.3 |
| Cabeceras de aislamiento | `FiltroCabeceras.java`, `ServidorHttp.java`, `seguridad.py` | Dos cabeceras nuevas en la web y en el teléfono-servidor; dos casos nuevos en la regresión de seguridad |
| Regresión de seguridad | `pruebas/seguridad.py` | **446/446**. Maven (`package` y `test`) en verde |
| Comprobación en vivo | — | Pruebas de baja tasa con una cuenta desechable, creada y borrada en la misma sesión; sin cambios en los datos de nadie más |

## 2. Cómo seguir

- Los cambios de cabeceras están **sin commit y sin publicar**. Al preparar la siguiente versión, van con el resto y se
  comprueba con `NOM_PRUEBA_RAPIDA=1 python3 4-CODIGO/pruebas/seguridad.py`.
- Lo que sigue para la entrega del 27-nov y los pendientes de paridad están en `RELEVO-22NOV.md` y en
  `PLAN-CONSTRUCCION.md` (fase 24).
- El servidor de desarrollo sigue siendo el jar de la 0.22.0 en `trabajo/srv/nom-dev.jar` (puerto 8096).
