# Prepara la computadora con Windows (y Office 2024) para correr la regresión de NOM ahí, sin permisos de administrador:
# Java 17 portátil, los paquetes de Python, los navegadores de Playwright y .NET 8 (el validador de Microsoft).
# Todo queda en %USERPROFILE%\nomoffice\herramientas y en el Python del usuario. Se puede correr otra vez: lo que ya
# está no se vuelve a bajar.
#   powershell -NoProfile -ExecutionPolicy Bypass -File preparar_windows.ps1
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$h = Join-Path $env:USERPROFILE "nomoffice\herramientas"
New-Item -ItemType Directory -Force -Path $h | Out-Null

# 1. Java 17 (Temurin, el mismo que corre el servidor en la Raspberry)
$jdk = Get-ChildItem $h -Directory -Filter "jdk-17*" -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $jdk) {
  Write-Output "== Java 17"
  Invoke-WebRequest -Uri "https://api.adoptium.net/v3/binary/latest/17/ga/windows/x64/jdk/hotspot/normal/eclipse" -OutFile "$h\jdk.zip"
  Expand-Archive -Path "$h\jdk.zip" -DestinationPath $h -Force
  Remove-Item "$h\jdk.zip"
  $jdk = Get-ChildItem $h -Directory -Filter "jdk-17*" | Select-Object -First 1
}
Write-Output ("java: " + $jdk.FullName)

# 2. Paquetes de Python (los mismos que pruebas/preparar.sh)
Write-Output "== Python"
python -m pip install --user -q --upgrade pip
python -m pip install --user -q playwright python-docx python-pptx openpyxl lxml pillow numpy opencv-python-headless cryptography pypdfium2
$env:PLAYWRIGHT_BROWSERS_PATH = Join-Path $env:USERPROFILE ".cache\nom-pruebas\navegadores"
python -m playwright install chromium firefox

# 3. .NET 8 (para el validador de Microsoft, Open XML SDK)
$dn = Join-Path $env:USERPROFILE ".dotnet"
if (-not (Test-Path "$dn\dotnet.exe")) {
  Write-Output "== .NET 8"
  Invoke-WebRequest -Uri "https://dot.net/v1/dotnet-install.ps1" -OutFile "$h\dotnet-install.ps1"
  & "$h\dotnet-install.ps1" -Channel 8.0 -InstallDir $dn | Out-Null
}
Write-Output ("dotnet: " + (Test-Path "$dn\dotnet.exe"))
Write-Output "listo"
