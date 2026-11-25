#!/bin/sh
# Instala el servidor propio de NOM en esta computadora, siempre prendido en
# https://nomcolab.dpdns.org (Cloudflare Tunnel) y también en Tailscale Funnel.
# Requiere: la copia de entrega en ~/NOM (con 2-PAGINA-Servidor), cloudflared (paquete de
# pkg.cloudflare.com) y haber corrido una vez `cloudflared tunnel login`.
set -e
AQUI="$(dirname "$(readlink -f "$0")")"
D="$HOME/nom-servidor"
U="$HOME/.config/systemd/user"
mkdir -p "$D/datos" "$D/respaldos" "$U" "$HOME/.local/bin"
[ -f "$D/nom.jar" ] || cp "$HOME/NOM/2-PAGINA-Servidor/nom.jar" "$D/nom.jar"
for s in respaldar.sh vigia.sh; do cp "$AQUI/$s" "$D/$s" && chmod +x "$D/$s"; done
for u in nom.service nom-tunel.service nom-respaldo.service nom-respaldo.timer \
         nom-vigia.service nom-vigia.timer; do cp "$AQUI/$u" "$U/"; done
ln -sf "$(readlink -f "$AQUI/nom")" "$HOME/.local/bin/nom"
# Que los servicios arranquen con la computadora sin iniciar sesión, y que Tailscale no pida sudo.
sudo loginctl enable-linger "$USER"
command -v tailscale >/dev/null && sudo tailscale set --operator="$USER" || true
# Que no se duerma: ni por la tapa ni por estar sin usar.
sudo mkdir -p /etc/systemd/logind.conf.d
sudo cp "$AQUI/logind-siempre-prendido.conf" /etc/systemd/logind.conf.d/nom-siempre-prendido.conf
sudo systemctl mask sleep.target suspend.target hibernate.target hybrid-sleep.target suspend-then-hibernate.target
[ -f "$HOME/.cloudflared/nom.yml" ] || "$AQUI/tunel.sh"
systemctl --user daemon-reload
systemctl --user enable --now nom.service nom-tunel.service nom-respaldo.timer nom-vigia.timer
command -v tailscale >/dev/null && tailscale funnel --bg 8080 >/dev/null 2>&1 || true
echo "Listo: https://nomcolab.dpdns.org (estado: nom local estado)"
