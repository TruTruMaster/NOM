#!/usr/bin/env bash
# Red de regresión: corre TODAS las pruebas automáticas y deja un resumen.
# Sirve para demostrar que lo que ya funcionaba sigue funcionando después de cualquier cambio.
#
#   4-CODIGO/pruebas/regresion.sh                 # todo (unos 60-80 min)
#   4-CODIGO/pruebas/regresion.sh rapido          # sin Office ni ancho de celular (unos 15 min)
#
# Necesita el jar compilado (cd 4-CODIGO/app && mvn -o -q package) y el servidor de pruebas en
# el puerto 8099 con la interfaz en vivo (ver docs/RELEVO-PRUEBAS-0.12.2.md, sección 2). seguridad.py
# y prueba_limites.py arrancan su propio servidor limpio. Los registros quedan en
# ~/.cache/nom-pruebas/registros/<fecha>/ y el resumen en resumen.txt de esa carpeta.
set -u
AQUI="$(cd "$(dirname "$0")" && pwd)"; RAIZ="$(cd "$AQUI/../.." && pwd)"
CACHE="${NOM_PRUEBAS_CACHE:-$HOME/.cache/nom-pruebas}"
export PLAYWRIGHT_BROWSERS_PATH="$CACHE/navegadores"
PY="$CACHE/venv/bin/python"
OUT="$CACHE/registros/$(date +%Y-%m-%d_%H%M)"; mkdir -p "$OUT/capturas"
MODO="${1:-todo}"
RES="$OUT/resumen.txt"; : > "$RES"
curl -s -o /dev/null http://localhost:8099/api/red/info || { echo "No hay servidor en el puerto 8099"; exit 2; }
echo "Versión: $(curl -s http://localhost:8099/api/red/info | sed 's/.*"version":"\([^"]*\)".*/\1/')  ·  $(date)" | tee -a "$RES"
# El QR para unirse solo sale si el servidor se anuncia en la red local (sin NOM_DESCUBRIMIENTO=false);
# si no, prueba_funciones marca dos "MAL" del QR que no son defectos de la app.
curl -s http://localhost:8099/api/red/info | grep -q '"lan":true' || echo "AVISO: el servidor de pruebas no se anuncia en la red local: los 2 casos del QR van a fallar" | tee -a "$RES"

corre() {   # corre <nombre> <comando...>
  local n="$1"; shift
  local t0=$(date +%s)
  ( cd "$RAIZ/4-CODIGO" && timeout 3600 "$@" ) > "$OUT/$n.txt" 2>&1
  local c=$? t=$(( $(date +%s) - t0 ))
  local ult; ult=$(grep -E -i "de [0-9]+|aprobad|correct|resultado|fall|tests run" "$OUT/$n.txt" | tail -2 | tr '\n' ' ' | cut -c1-150)
  printf '%-26s %-6s %4ss  %s\n' "$n" "$([ $c -eq 0 ] && echo BIEN || echo FALLA)" "$t" "$ult" | tee -a "$RES"
}

corre junit            bash -c "cd app && mvn -o -q test"
corre seguridad        python3 pruebas/seguridad.py
corre limites          python3 "$RAIZ/6-PLAN-DE-PRUEBAS/fuentes/prueba_limites.py"
corre carga            python3 pruebas/carga.py --personas 5 --segundos 10
corre cuentas_equipo   "$PY" pruebas/interfaz/prueba_cuentas_equipo.py "$OUT/capturas"
corre sin_internet     "$PY" pruebas/interfaz/prueba_sin_internet.py "$OUT/capturas"
corre e2e              "$PY" pruebas/interfaz/prueba_e2e.py "$OUT/capturas"
corre ayuda            "$PY" pruebas/interfaz/prueba_ayuda.py "$OUT/capturas"
corre red              "$PY" pruebas/interfaz/prueba_red.py "$OUT/capturas"
corre funciones        "$PY" pruebas/interfaz/prueba_funciones.py "$OUT/capturas"
corre rastreo          "$PY" pruebas/interfaz/rastreo.py "$OUT/capturas"
for t in prueba_pres1 prueba_pptx prueba_pptx2 prueba_diag1 prueba_85 prueba_citas prueba_datos prueba_biblioteca prueba_fuentes_api; do
  corre "$t" "$PY" "pruebas/interfaz/$t.py" "$OUT/capturas"
done
if [ "$MODO" != rapido ]; then
  corre funciones_movil  "$PY" pruebas/interfaz/prueba_funciones.py "$OUT/capturas" --movil
  corre rastreo_movil    "$PY" pruebas/interfaz/rastreo.py "$OUT/capturas" --movil
  corre diag_movil       "$PY" pruebas/interfaz/prueba_diag_movil.py "$OUT/capturas"
  corre funciones_firefox env NOM_NAVEGADOR=firefox "$PY" pruebas/interfaz/prueba_funciones.py "$OUT/capturas"
  # los 123 archivos reales (sin argumentos solo usa corpus/, 47); corpus-reales sale de preparar.sh
  CORPUS="$CACHE/corpus/*/* $CACHE/corpus-web/*/* $CACHE/corpus-reales/*/*"
  corre office_fidelidad bash -c "\"$PY\" pruebas/office/fidelidad.py $CORPUS"
  corre office_modelo    bash -c "\"$PY\" pruebas/office/fidelidad.py --modelo $CORPUS"
  corre office_ediciones "$PY" pruebas/office/ediciones.py --office
  # Un Word largo de verdad (78 páginas): escribir sin que se trabe, otra persona, recargar y descargar.
  corre documento_grande "$PY" pruebas/office/documento_grande.py "$CACHE/corpus/docx/ms-informe-gobernanza.docx" "$OUT/documento-grande"
  # Archivos reales editados por dos personas (Chromium y Firefox) y descargados (plantillas de Microsoft).
  R="$CACHE/corpus-reales"
  corre editar_real      "$PY" pruebas/office/editar_real.py "$OUT/editar-real" "$R/xlsx/ms-plantilla-presupuesto-universitario.xlsx" \
                           "$R/xlsx/ms-plantilla-tablas-dinamicas.xlsx" "$R/pptx/ms-plantilla-presentacion-basica.pptx" "$R/pptx/lo-group-rot.pptx" \
                           "$CACHE/corpus/docx/ms-complex2010.docx"
  corre office_raros     "$PY" pruebas/office/raros.py
  corre funciones_excel  "$PY" pruebas/office/funciones_excel.py  # las funciones de la hoja contra Excel 2024 real
  # Lo editado y exportado, con el validador oficial de Microsoft (Open XML SDK): 0 errores nuevos contra el original.
  if [ -x "$HOME/.dotnet/dotnet" ]; then corre office_openxml python3 pruebas/office/validar_openxml.py "$CACHE/office-ediciones"
  else echo "AVISO: sin .NET 8 en ~/.dotnet no corre el validador de Microsoft (ver pruebas/office/validar_openxml.py)" | tee -a "$RES"; fi
fi
echo "Resumen: $RES"
grep -c " FALLA " "$RES" | xargs -I{} echo "{} suite(s) con falla"
