# correr.ps1 <id> <script.py> [args...]  -> lanza en la sesion de escritorio y espera
param([string]$id, [Parameter(ValueFromRemainingArguments=$true)][string[]]$resto)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$base = "$env:USERPROFILE\nomoffice"
$pyw = (Get-Command python).Source -replace 'python.exe$','pythonw.exe'
@{ id = $id; args = $resto } | ConvertTo-Json -Compress | Set-Content -Encoding UTF8 "$base\trabajo.json"
Remove-Item -ErrorAction SilentlyContinue "$base\trabajos\$id.fin","$base\trabajos\$id.log"
schtasks /Create /F /TN nomoffice /SC ONCE /ST 23:59 /IT /TR "`"$pyw`" `"$base\lanzador.py`"" | Out-Null
schtasks /Run /TN nomoffice | Out-Null
$t0 = Get-Date
while (-not (Test-Path "$base\trabajos\$id.fin")) { Start-Sleep -Milliseconds 500; if (((Get-Date)-$t0).TotalMinutes -gt 60) { "TIEMPO AGOTADO"; break } }
Get-Content -Encoding UTF8 "$base\trabajos\$id.log"
"[fin " + (Get-Content "$base\trabajos\$id.fin" -ErrorAction SilentlyContinue) + "]"
