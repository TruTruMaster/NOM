#!/bin/sh
# dominio-corto.sh: prepara la Raspberry para el nombre corto nom.qd.je (una sola vez, como root).
#
#   doas sh dominio-corto.sh
#
# 1. El túnel nomcolab también atiende www.nom.qd.je (y nom.qd.je) (/etc/cloudflared/nomcolab.yml).
# 2. La nube que firma los pases sigue siendo https://nomcolab.dpdns.org (NOM_NUBE), así los APK ya
#    instalados y los teléfonos que comparten siguen aceptando los pases aunque la cuenta se abra
#    con el nombre corto.
# 3. Instala la orden nomcolab-pi nueva (con «nomcolab-pi dominio», que no pide contraseña).
# 3b. La clave del reenvío de Vercel (NOM_PROXY_SECRETO, del archivo secreto-vercel junto a este
#    script): con ella el servidor cree la IP de cada persona que llega por nom.qd.je.
# 4. Si nom.qd.je ya llega aquí, lo vuelve la dirección principal para los buscadores
#    (NOM_CANONICO). Si todavía no, después basta con:  nomcolab-pi dominio
#
# Los pasos de afuera (Vercel y el DNS de nom.qd.je) están en docs/DOMINIO-Y-BUSCADORES.md.
set -eu
[ "$(id -u)" = 0 ] || { echo "dominio-corto: hay que correrlo como root (doas sh $0)"; exit 1; }
AQUI="$(cd "$(dirname "$0")" && pwd)"
RUN=/run/service
di() { echo "dominio-corto: $*"; }

# 1. el túnel
install -m 640 -o root -g cloudflared "$AQUI/nomcolab-tunel.yml" /etc/cloudflared/nomcolab.yml.n
/usr/local/bin/cloudflared --config /etc/cloudflared/nomcolab.yml.n tunnel ingress validate >/dev/null \
  || { rm -f /etc/cloudflared/nomcolab.yml.n; di "la configuración del túnel no es válida; no cambié nada"; exit 1; }
mv /etc/cloudflared/nomcolab.yml.n /etc/cloudflared/nomcolab.yml
di "túnel: nomcolab.dpdns.org, www.nomcolab.dpdns.org, www.nom.qd.je y nom.qd.je"

# 2. la nube de los pases
reiniciar=0
if ! grep -q '^export NOM_NUBE=' /etc/nomcolab.conf 2>/dev/null; then
  printf '\n# La nube que firma los pases: la que conocen los APK ya instalados (dominio-corto.sh, %s).\nexport NOM_NUBE=https://nomcolab.dpdns.org\n' "$(date +%Y-%m-%d)" >> /etc/nomcolab.conf
  reiniciar=1
  di "pases: firmados como https://nomcolab.dpdns.org"
fi

# 3b. la clave del reenvío de Vercel
if [ -s "$AQUI/secreto-vercel" ]; then
  clave=$(tr -dc 'a-f0-9' < "$AQUI/secreto-vercel")
  [ ${#clave} -ge 32 ] || { di "secreto-vercel no tiene una clave válida"; exit 1; }
  if ! grep -q "^export NOM_PROXY_SECRETO=$clave\$" /etc/nomcolab.conf; then
    grep -v '^export NOM_PROXY_SECRETO=' /etc/nomcolab.conf > /etc/nomcolab.conf.n || true
    printf '# Clave del reenvío de Vercel para nom.qd.je (dominio-corto.sh, %s).\nexport NOM_PROXY_SECRETO=%s\n' "$(date +%Y-%m-%d)" "$clave" >> /etc/nomcolab.conf.n
    cat /etc/nomcolab.conf.n > /etc/nomcolab.conf && rm -f /etc/nomcolab.conf.n
    reiniciar=1
    di "clave de Vercel: puesta"
  fi
  chown root:root /etc/nomcolab.conf && chmod 600 /etc/nomcolab.conf
  rm -f "$AQUI/secreto-vercel"
fi

# 3. la orden nomcolab-pi
install -m 755 "$AQUI/nomcolab-pi" /usr/local/sbin/nomcolab-pi

# reinicios: el túnel siempre (lee el nombre nuevo); el servidor solo si cambió su configuración
s6-svc -r $RUN/nomcolab-tunel
if [ $reiniciar = 1 ]; then
  s6-svc -r $RUN/nomcolab-servidor
  i=0; until wget -q -O /dev/null -T 3 http://127.0.0.1:8080/api/red/info 2>/dev/null; do
    i=$((i + 1)); [ $i -ge 120 ] && { di "el servidor no contesta tras 120 s (nomcolab-pi bitacora)"; exit 1; }; sleep 1
  done
fi
sleep 5

# 4. ¿ya llega nom.qd.je?
/usr/local/sbin/nomcolab-pi dominio || di "cuando nom.qd.je ya apunte a Vercel corre: nomcolab-pi dominio"
/usr/local/sbin/nomcolab-pi estado
