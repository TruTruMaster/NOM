#!/usr/bin/env bash
# Todas las pruebas en vivo contra www.nom.qd.je, una tras otra (nunca en paralelo: la Pi es el DNS de la casa).
#   pruebas/vivo.sh CARPETA [prueba ...]     (sin pruebas: la tanda completa)
# Al final escribe CARPETA/resumen.txt con los ok/MAL de cada una.
set -u
AQUI="$(cd "$(dirname "$0")" && pwd)"
OUT="${1:?carpeta de salida}"; shift
mkdir -p "$OUT"
PY="$HOME/.cache/nom-pruebas/venv/bin/python"
export PLAYWRIGHT_BROWSERS_PATH="$HOME/.cache/nom-pruebas/navegadores" NOM_BASE="${NOM_BASE:-https://www.nom.qd.je}"
TODAS=(humo demo armados funciones funciones_movil funciones_firefox e2e sin_internet red presentacion pres1 pptx2 diag1 p85 ayuda axe accesibilidad rastreo)
[ $# -eq 0 ] && set -- "${TODAS[@]}"
for t in "$@"; do
  echo "=== $t $(date +%H:%M:%S)"
  case "$t" in
    humo) timeout 900 "$PY" "$AQUI/humo.py" > "$OUT/$t.txt" 2>&1 ;;
    funciones) timeout 1800 "$PY" "$AQUI/interfaz/prueba_funciones.py" "$OUT/cap-$t" > "$OUT/$t.txt" 2>&1 ;;
    funciones_movil) timeout 1800 "$PY" "$AQUI/interfaz/prueba_funciones.py" "$OUT/cap-$t" --movil > "$OUT/$t.txt" 2>&1 ;;
    funciones_firefox) NOM_NAVEGADOR=firefox timeout 1800 "$PY" "$AQUI/interfaz/prueba_funciones.py" "$OUT/cap-$t" > "$OUT/$t.txt" 2>&1 ;;
    axe) timeout 1200 "$PY" "$AQUI/interfaz/prueba_axe.py" "$OUT/axe.json" > "$OUT/$t.txt" 2>&1 ;;
    p85) timeout 1200 "$PY" "$AQUI/interfaz/prueba_85.py" "$OUT/cap-$t" > "$OUT/$t.txt" 2>&1 ;;
    rastreo) timeout 2400 "$PY" "$AQUI/interfaz/rastreo.py" "$OUT/cap-$t" > "$OUT/$t.txt" 2>&1 ;;
    *) timeout 1500 "$PY" "$AQUI/interfaz/prueba_$t.py" "$OUT/cap-$t" > "$OUT/$t.txt" 2>&1 ;;
  esac
  echo "    salida $? · ok $(grep -c '^ *ok ' "$OUT/$t.txt") · MAL $(grep -c 'MAL' "$OUT/$t.txt") · $(tail -1 "$OUT/$t.txt" | cut -c1-150)"
done 2>&1 | tee -a "$OUT/resumen.txt"
