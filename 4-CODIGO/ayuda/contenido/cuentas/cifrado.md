---
titulo: Cifrado: en tu equipo, en el camino y en el servidor
descripcion: Qué va cifrado, dónde, y quién podría leerlo.
orden: 50
aplica: Navegador, App Android, Teléfono que comparte
actualizado: 2026-11-13
version: 0.13.0
---
| Dónde | Cómo va | Quién podría leerlo |
|---|---|---|
| En tu equipo | Cifrado con tu contraseña (bóveda) | Solo tú, con tu contraseña |
| En el camino, por Internet | HTTPS hasta Cloudflare y después por su túnel cifrado hasta el servidor | Nadie en tu Wi-Fi ni en tu proveedor de Internet. Cloudflare (por donde pasa el túnel) y quien administra el servidor sí podrían, salvo en los espacios cifrados de extremo a extremo |
| En el camino, en la red del salón | **Sin cifrar** (http) | Alguien en el mismo Wi-Fi; ver [Conexión sin cifrar](ayuda:cuentas/conexion-sin-cifrar) |
| En el servidor (disco y respaldos) | Cifrado en reposo (AES-256-GCM) | Quien administra el servidor, que tiene la llave |
| En un espacio **cifrado de extremo a extremo** | Cifrado en tu equipo antes de enviarse | Solo los miembros del espacio |

El cifrado en reposo protege si se roban el disco del servidor o una copia de los respaldos. Para
que **ni quien administra ni Cloudflare** puedan leer, crea el espacio con
[cifrado de extremo a extremo](ayuda:cuentas/extremo-a-extremo).
