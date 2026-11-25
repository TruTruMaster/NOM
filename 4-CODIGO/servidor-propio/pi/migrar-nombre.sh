#!/bin/sh
# Una sola vez (25-nov-2026): pasa la instalación de la Pi de los nombres «nomcolab» a «nom».
#   usuario nomcolab -> nom (mismo UID y GID: los archivos siguen siendo suyos)
#   /var/lib/nomcolab -> /var/lib/nom · /opt/nomcolab/nom-colab.jar -> /opt/nom/nom.jar
#   /etc/nomcolab.conf -> /etc/nom.conf · /etc/cloudflared/nomcolab.yml -> nom.yml
#   servicios s6 nomcolab-* -> nom-*, bitácoras /var/log/nomcolab-* -> /var/log/nom-*
#   la orden nomcolab-pi -> nom-pi (y su regla de doas), el respaldo diario
# Después corre instalar-pi.sh (que está junto a este archivo) con el jar que ya estaba.
# Se corre como root:  doas sh migrar-nombre.sh
set -eu
[ "$(id -u)" = 0 ] || { echo "migrar: hay que correrlo como root"; exit 1; }
AQUI="$(cd "$(dirname "$0")" && pwd)"
m() { echo "migrar: $*"; }
if ! [ -d /var/lib/nomcolab ]; then m "ya no hay /var/lib/nomcolab: nada que migrar"; exit 0; fi
[ -f "$AQUI/instalar-pi.sh" ] || { m "falta instalar-pi.sh junto a este archivo"; exit 1; }
[ -e /var/lib/nom ] && { m "/var/lib/nom ya existe: no sigo"; exit 1; }

m "respaldo antes de empezar"
/usr/local/sbin/nomcolab-pi respaldo

m "apago los servicios viejos"
s6-rc -d change nomcolab

m "usuario y grupo: nomcolab -> nom (mismo UID/GID)"
for f in /etc/passwd /etc/group /etc/shadow; do cp -p "$f" "$f.antes-de-nom"; done
sed -i 's|^nomcolab:x:\([0-9]*\):\([0-9]*\)::/var/lib/nomcolab:|nom:x:\1:\2::/var/lib/nom:|' /etc/passwd
sed -i 's|^nomcolab:x:\([0-9]*\):nomcolab$|nom:x:\1:nom|' /etc/group
sed -i 's|^nomcolab:|nom:|' /etc/shadow
id nom >/dev/null && getent group nom >/dev/null || { m "no quedó el usuario nom"; exit 1; }

m "carpetas y archivos"
mv /var/lib/nomcolab /var/lib/nom
mv /opt/nomcolab /opt/nom
mv /opt/nom/nom-colab.jar /opt/nom/nom.jar
[ -f /opt/nom/nom-colab.jar.anterior ] && mv /opt/nom/nom-colab.jar.anterior /opt/nom/nom.jar.anterior
mv /etc/nomcolab.conf /etc/nom.conf
sed -i 's|/var/lib/nomcolab|/var/lib/nom|g; s|NOM-COLAB|NOM|g; s|nomcolab-servidor|nom-servidor|g; s|nomcolab-pi|nom-pi|g' /etc/nom.conf
mv /etc/cloudflared/nomcolab.yml /etc/cloudflared/nom.yml
for s in servidor tunel vigia; do [ -d /var/log/nomcolab-$s ] && mv /var/log/nomcolab-$s /var/log/nom-$s; done

m "quito lo viejo de s6 y el respaldo diario viejo"
for s in nomcolab nomcolab-servidor nomcolab-servidor-log nomcolab-tunel nomcolab-tunel-log nomcolab-vigia nomcolab-vigia-log; do
  rm -rf "/etc/s6-rc/source/$s"
done
rm -f /etc/s6-rc/source/default/contents.d/nomcolab /etc/periodic/daily/nomcolab-respaldo

m "instalador con los nombres nuevos"
sh "$AQUI/instalar-pi.sh"

m "quito la orden vieja y su regla de doas"
rm -f /usr/local/sbin/nomcolab-pi /etc/doas.d/zz-nomcolab.conf
m "listo"
