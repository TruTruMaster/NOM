# nom.qd.je por Vercel

`nom.qd.je` es la raíz de su zona en DigitalPlat: no puede ser CNAME, y Cloudflare for SaaS solo
acepta raíces con el plan Enterprise. Vercel sí acepta la raíz (registro A) y reenvía cada petición,
sin cambiar la dirección que ve la persona, a `https://nomcolab.dpdns.org`, que sigue llegando a la
Raspberry por el túnel de Cloudflare. La casa no se expone y no se abre ningún puerto.

- `vercel.json`: todo se reenvía; Vercel no guarda copias (`x-vercel-enable-rewrite-caching: 0`) y
  agrega en cada petición `x-nom-proxy` con la clave `NOM_PROXY_SECRETO` (variable del proyecto en
  Vercel, reemplaza la que mande cualquiera). El servidor solo cree `X-Real-IP` si la clave coincide
  con la suya (`Origen.deVercel`); así los límites contra abuso ven a cada persona.
- Vercel corta cada petición reenviada a los 120 s: el canal en vivo se vuelve a abrir solo, sin
  avisar (`api.eventos`).
- La clave vive fuera del repositorio (`~/.config/nomcolab/secreto-vercel` en la laptop) y en
  `/etc/nomcolab.conf` de la Pi (la pone `servidor-propio/pi/dominio-corto.sh`).

Publicar (con el token de Vercel en `~/NOM/tokenvercel`):

```sh
cd 4-CODIGO/servidor-propio/vercel
npx vercel@latest deploy --prod --token "$(cat ~/NOM/tokenvercel)"
```
