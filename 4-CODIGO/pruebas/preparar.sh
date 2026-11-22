#!/usr/bin/env bash
# Prepara (o reconstruye) todo lo que necesitan las pruebas en esta computadora. Queda en
# ~/.cache/nom-pruebas, fuera del repositorio: si esa carpeta se borra (pasó el 15-nov-2026), se corre
# esto otra vez y en unos minutos queda igual.
#
#   4-CODIGO/pruebas/preparar.sh            # todo lo que falte
#
# 1. Python con Playwright, python-docx, python-pptx, openpyxl, lxml, Pillow y OpenCV (venv propio).
# 2. Los navegadores de prueba: Chromium y Firefox (de Playwright).
# 3. Tres juegos de archivos reales de Office:
#    - corpus/        Open XML SDK de Microsoft y Apache POI (office/corpus.sh)
#    - corpus-web/    archivos de usuarios de las pruebas de LibreOffice y del paquete officer
#    - corpus-reales/ plantillas oficiales de Microsoft (presupuesto, tablas dinámicas, fórmulas, Gantt,
#                     registro escolar, presentación) y más archivos reales de LibreOffice
# 4. axe-core (npm), para la revisión automática de accesibilidad de la interfaz.
# 5. El validador de Microsoft (Open XML SDK) compilado, si hay .NET 8 en ~/.dotnet.
set -u
AQUI="$(cd "$(dirname "$0")" && pwd)"
CACHE="${NOM_PRUEBAS_CACHE:-$HOME/.cache/nom-pruebas}"
mkdir -p "$CACHE"
PY="$CACHE/venv/bin/python"
export PLAYWRIGHT_BROWSERS_PATH="$CACHE/navegadores"

echo "== 1. Python"
[ -x "$PY" ] || python3 -m venv "$CACHE/venv"
"$CACHE/venv/bin/pip" install -q --upgrade pip
"$CACHE/venv/bin/pip" install -q playwright python-docx python-pptx openpyxl lxml pillow opencv-python-headless

echo "== 2. Navegadores"
"$CACHE/venv/bin/playwright" install chromium firefox

baja() {  # baja <destino> <dirección>
  [ -s "$1" ] && return 0
  if curl -fsSL -m 120 --retry 2 -o "$1.n" "$2" && [ -s "$1.n" ]; then mv "$1.n" "$1"; echo "ok   $(basename "$1")"
  else rm -f "$1.n"; echo "FALLÓ $(basename "$1")  <- $2"; fi
}

echo "== 3a. corpus (Microsoft y Apache POI)"
bash "$AQUI/office/corpus.sh" "$CACHE/corpus"

echo "== 3b. corpus-web (LibreOffice y officer)"
LO="https://raw.githubusercontent.com/LibreOffice/core/master"
W="$CACHE/corpus-web"; mkdir -p "$W/docx" "$W/xlsx" "$W/pptx"
for f in fdo74605 fdo76316 fdo77476 fdo77718 FDO78292 FigureAsLabelPicture pictureWatermark tdf104061_tableSectionColumns tdf104420_lostParagraph tdf107111; do
  baja "$W/docx/lo-$f.docx" "$LO/sw/qa/extras/ooxmlexport/data/$f.docx"; done
for f in groupshape-fontname tdf130804 tdf143476_lockedCanvas_position tdf75573_page1frame; do
  baja "$W/docx/lo-$f.docx" "$LO/sw/qa/extras/ooxmlimport/data/$f.docx"; done
for f in accent-color crop-to-shape cshapes multiplelayoutfooter n821567 slide-sections smartart-linear-rule-vert tdf106638 tdf120573 tdf134174 tdf149206 tdf169781 tdf89927; do
  baja "$W/pptx/lo-$f.pptx" "$LO/sd/qa/unit/data/pptx/$f.pptx"; done
for f in cell-anchored-hidden-shapes cell-borders chart_hyperlink column-style-autofilter forcepoint107 forum-mso-de-104083 forum-mso-en4-134670 functions-excel-2010 pivottable_invalid_formats TableEmptyHeaders tdf122471 tdf130959 tdf83671_SmartArt_import totalsRowFunction; do
  baja "$W/xlsx/lo-$f.xlsx" "$LO/sc/qa/unit/data/xlsx/$f.xlsx"; done
OF="https://raw.githubusercontent.com/davidgohel/officer/master/tests/testthat/docs_dir"
baja "$W/docx/officer-table-complex.docx" "$OF/table-complex.docx"
baja "$W/pptx/officer-table-complex.pptx" "$OF/table-complex.pptx"

echo "== 3c. corpus-reales (plantillas de Microsoft y más archivos reales)"
R="$CACHE/corpus-reales"; mkdir -p "$R/xlsx" "$R/pptx" "$CACHE/plantillas"
MS="https://binaries.templates.cdn.office.net/support/templates/en-us"
# Una plantilla (.xltx/.potx) pasa a documento cambiando solo el tipo de la parte principal, como hace
# Office al crear un documento desde una plantilla; las demás partes se copian tal cual.
plantilla() {  # plantilla <archivo del CDN> <destino>
  [ -s "$2" ] && return 0
  baja "$CACHE/plantillas/$1" "$MS/$1" >/dev/null || return 0
  "$PY" - "$CACHE/plantillas/$1" "$2" <<'EOF' && echo "ok   $(basename "$2")"
import sys, zipfile
cambios = {b"spreadsheetml.template.main+xml": b"spreadsheetml.sheet.main+xml", b"presentationml.template.main+xml": b"presentationml.presentation.main+xml"}
with zipfile.ZipFile(sys.argv[1]) as zi, zipfile.ZipFile(sys.argv[2], "w", zipfile.ZIP_DEFLATED) as zo:
    for it in zi.infolist():
        d = zi.read(it.filename)
        if it.filename == "[Content_Types].xml":
            for a, b in cambios.items(): d = d.replace(a, b)
        zo.writestr(it, d)
EOF
}
plantilla tf16400647_win32.xltx "$R/xlsx/ms-plantilla-tablas-dinamicas.xlsx"
plantilla tf16400653_win32.xltx "$R/xlsx/ms-plantilla-presupuesto-universitario.xlsx"
plantilla tf16400663_win32.xltx "$R/xlsx/ms-plantilla-registro-escolar.xlsx"
plantilla tf16400645_win32.xltx "$R/xlsx/ms-plantilla-activos.xlsx"
plantilla tf16400662_win32.xltx "$R/xlsx/ms-plantilla-ventas.xlsx"
plantilla tf16400656_win32.xltx "$R/xlsx/ms-plantilla-tutorial-formulas.xlsx"
plantilla tf16400962_win32.xltx "$R/xlsx/ms-plantilla-gantt.xlsx"
plantilla tf16411248_win32.potx "$R/pptx/ms-plantilla-presentacion-basica.pptx"
for f in accent-color tdf106638 tdf89927 tdf148478_transitionMusic tdf120573 slide-sections smartart-org-chart2 tdf104792-smart-art-animation group-rot predefined-table-style video-loop-mute cshapes n821567 tdf134174; do
  baja "$R/pptx/lo-$f.pptx" "$LO/sd/qa/unit/data/pptx/$f.pptx"; done
for f in forum-mso-de-104083 forum-mso-en4-134670 tdf122471 cell-anchored-hidden-shapes tdf131575 functions-excel-2010 pivottable_tabular_mode tdf83671_SmartArt_import chart_hyperlink totalsRowFunction tdf111980_radioButtons; do
  baja "$R/xlsx/lo-$f.xlsx" "$LO/sc/qa/unit/data/xlsx/$f.xlsx"; done

echo "== 4. axe-core (revisión automática de accesibilidad, WCAG 2.2 AA: interfaz/prueba_axe.py)"
if command -v npm >/dev/null; then
  [ -f "$CACHE/axe/node_modules/axe-core/axe.min.js" ] || (mkdir -p "$CACHE/axe" && cd "$CACHE/axe" && npm install -q --no-audit --no-fund --ignore-scripts axe-core@4.10.3 >/dev/null)
  ls "$CACHE/axe/node_modules/axe-core/axe.min.js" >/dev/null && echo "axe-core listo"
else
  echo "Falta npm (Node.js) para bajar axe-core"
fi

echo "== 5. Validador de Microsoft"
if [ -x "$HOME/.dotnet/dotnet" ]; then
  DOTNET_CLI_TELEMETRY_OPTOUT=1 DOTNET_NOLOGO=1 "$HOME/.dotnet/dotnet" build "$AQUI/office/openxml" -c Release -o "$CACHE/openxml-bin" 2>&1 | tail -2
else
  echo "Falta .NET 8: curl -sSL https://dot.net/v1/dotnet-install.sh | bash -s -- --channel 8.0 --install-dir ~/.dotnet"
fi
echo "Listo: $(ls "$CACHE"/corpus/*/* "$CACHE"/corpus-web/*/* "$CACHE"/corpus-reales/*/* 2>/dev/null | wc -l) archivos de Office en $CACHE"
