#!/bin/sh
# Publica 3-APP-Android/NOM.apk como la release "más reciente" en GitHub, para que el
# botón "Descarga la app" de la página (https://.../releases/latest/download/NOM.apk) dé
# siempre esta versión. El archivo lo sirve la red de GitHub, no la conexión del servidor.
#
# Correrlo después de compilar el APK (cuando ya está copiado en 3-APP-Android/NOM.apk).
# Necesita el token de GitHub en ~/NOM/tokensito.
set -e
REPO="TruTruMaster/NOM"
APK="$HOME/NOM/3-APP-Android/NOM.apk"
TOK="$(cat "$HOME/NOM/tokensito")"
VER="$(grep -oE 'versionName = "[^"]+"' "$HOME/NOM/4-CODIGO/android/app/build.gradle.kts" | head -1 | sed 's/.*"\([^"]*\)".*/\1/')"
TAG="v$VER"
[ -f "$APK" ] || { echo "No existe $APK (compila el APK primero)"; exit 1; }
[ -n "$VER" ] || { echo "No pude leer la versión"; exit 1; }
api() { curl -s -H "Authorization: Bearer $TOK" -H 'Accept: application/vnd.github+json' "$@"; }
# Para que cualquiera compruebe que el APK es el nuestro: su SHA-256 y la huella del certificado con que se firma.
SHA="$(sha256sum "$APK" | cut -d' ' -f1)"
BT="$(ls -d "$HOME"/Android/Sdk/build-tools/* 2>/dev/null | tail -1)"
CERT="$([ -n "$BT" ] && "$BT/apksigner" verify --print-certs "$APK" 2>/dev/null | sed -n 's/^Signer #1 certificate SHA-256 digest: //p' | head -1)"
NOTA="Aplicacion de Android (APK). Descargala con el boton de la pagina o desde el archivo de aqui abajo. En el telefono, abrela y permite instalar de origenes desconocidos.\n\nPara comprobar que el archivo es el original:\n- SHA-256 del APK: $SHA\n- Huella SHA-256 del certificado de firma: ${CERT:-(sin apksigner)}"

echo "Publicando NOM $TAG ($(du -h "$APK" | cut -f1))"
ID="$(api "https://api.github.com/repos/$REPO/releases/tags/$TAG" | python3 -c 'import sys,json;d=json.load(sys.stdin);print(d.get("id","") or "")')"
if [ -z "$ID" ]; then
  BODY="$(NOTA="$NOTA" python3 -c "import json,os;print(json.dumps({'tag_name':'$TAG','target_commitish':'main','name':'NOM $VER','body':os.environ['NOTA'].replace(chr(92)+'n',chr(10)),'make_latest':'true'}))")"
  ID="$(api -X POST "https://api.github.com/repos/$REPO/releases" -d "$BODY" | python3 -c 'import sys,json;print(json.load(sys.stdin)["id"])')"
  echo "  release creada (id $ID)"
else
  api -X PATCH "https://api.github.com/repos/$REPO/releases/$ID" -d "$(NOTA="$NOTA" python3 -c "import json,os;print(json.dumps({'make_latest':'true','body':os.environ['NOTA'].replace(chr(92)+'n',chr(10))}))")" >/dev/null
  echo "  release $TAG ya existía (id $ID), marcada como la más reciente"
  # Si ya tenía el APK, se borra para volver a subirlo.
  OLD="$(api "https://api.github.com/repos/$REPO/releases/$ID/assets" | python3 -c 'import sys,json;[print(a["id"]) for a in json.load(sys.stdin) if a["name"] in ("NOM.apk","NOM-COLAB.apk")]')"
  for a in $OLD; do api -X DELETE "https://api.github.com/repos/$REPO/releases/assets/$a" >/dev/null; echo "  asset anterior borrado"; done
fi
# NOM.apk es el nombre de hoy; NOM-COLAB.apk (el mismo archivo) es para las apps y los enlaces de antes del cambio de
# nombre, que buscan ese archivo en la release más reciente.
for NOMBRE in NOM.apk NOM-COLAB.apk; do
  curl -s -H "Authorization: Bearer $TOK" -H 'Content-Type: application/vnd.android.package-archive' \
    --data-binary @"$APK" "https://uploads.github.com/repos/$REPO/releases/$ID/assets?name=$NOMBRE" \
    | python3 -c 'import sys,json;d=json.load(sys.stdin);print("  subido:",d.get("name"),d.get("size"),d.get("state"))'
done
echo "Listo. Enlace fijo: https://github.com/$REPO/releases/latest/download/NOM.apk"
echo "SHA-256: $SHA · certificado: $CERT"
