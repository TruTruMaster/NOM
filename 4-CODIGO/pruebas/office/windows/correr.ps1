# correr.ps1 <id> <script.py> [args...]  -> lanza en la sesion de escritorio y espera
# Cada trabajo con su propia tarea programada y su propio archivo (trabajos\<id>.json): varios pueden correr a la vez.
param([string]$id, [Parameter(ValueFromRemainingArguments=$true)][string[]]$resto)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$base = "$env:USERPROFILE\nomoffice"
$pyw = (Get-Command python).Source -replace 'python.exe$','pythonw.exe'
New-Item -ItemType Directory -Force -Path "$base\trabajos" | Out-Null
@{ id = $id; args = $resto } | ConvertTo-Json -Compress | Set-Content -Encoding UTF8 "$base\trabajos\$id.json"
Remove-Item -ErrorAction SilentlyContinue "$base\trabajos\$id.fin","$base\trabajos\$id.log"
$tarea = "nomoffice-$id"
schtasks /Create /F /TN $tarea /SC ONCE /ST 23:59 /IT /TR "`"$pyw`" `"$base\lanzador.py`" $id" | Out-Null
schtasks /Run /TN $tarea | Out-Null
$t0 = Get-Date
while (-not (Test-Path "$base\trabajos\$id.fin")) { Start-Sleep -Milliseconds 500; if (((Get-Date)-$t0).TotalMinutes -gt 60) { "TIEMPO AGOTADO"; break } }
schtasks /Delete /F /TN $tarea 2>$null | Out-Null
Get-Content -Encoding UTF8 "$base\trabajos\$id.log"
"[fin " + (Get-Content "$base\trabajos\$id.fin" -ErrorAction SilentlyContinue) + "]"
