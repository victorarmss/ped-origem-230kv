# ==========================================================================
#  Vigiar_PED.ps1
#  Fica rodando e monitora o arquivo PED - ORIGEM 230kV.xlsx.
#  Quando voce SALVA a PED, ele espera o arquivo ser liberado e chama
#  Publicar_Auto.ps1 (regenera o index.html e faz o push -> Netlify publica).
#  Para iniciar: use "Iniciar Vigia.cmd".  Para parar: feche esta janela.
# ==========================================================================
$base = $PSScriptRoot
$ped  = Join-Path $env:USERPROFILE 'OneDrive - ENGETECNICA\EPC\1700-1701-1702 - ORIGEM\02. Engenharia\01. PED\PED - ORIGEM 230kV.xlsx'
$pub  = Join-Path $base 'Publicar_Auto.ps1'
$log  = Join-Path $base 'publicar_log.txt'
function Log($m){ Add-Content -LiteralPath $log -Value ("[{0}] (vigia) {1}" -f (Get-Date).ToString('dd/MM/yyyy HH:mm:ss'), $m) -Encoding UTF8 }
function FileLocked($p){ try { $s=[System.IO.File]::Open($p,'Open','Read','None'); $s.Close(); return $false } catch { return $true } }

Write-Host "Vigia da PED ativo. Ao salvar a planilha, o dashboard e publicado automaticamente." -ForegroundColor Green
Write-Host "Deixe esta janela aberta (pode minimizar). Feche-a para parar." -ForegroundColor Yellow
Log "Vigia iniciado."

if (-not (Test-Path -LiteralPath $ped)) { Write-Host "ERRO: PED nao encontrada." -ForegroundColor Red; Log "PED nao encontrada."; Start-Sleep 8; return }
$last = (Get-Item -LiteralPath $ped).LastWriteTimeUtc
$pending = $false; $since = $null

while ($true) {
  Start-Sleep -Seconds 5
  if (-not (Test-Path -LiteralPath $ped)) { continue }
  $cur = (Get-Item -LiteralPath $ped).LastWriteTimeUtc
  if ($cur -ne $last) { $last = $cur; $pending = $true; $since = Get-Date; Log "Alteracao detectada na PED."; continue }
  if ($pending -and ((Get-Date) - $since).TotalSeconds -ge 15 -and -not (FileLocked $ped)) {
    $pending = $false
    Write-Host ("[{0}] PED salva - publicando..." -f (Get-Date).ToString('HH:mm:ss')) -ForegroundColor Cyan
    Log "PED estavel e liberada - publicando..."
    & powershell -NoProfile -ExecutionPolicy Bypass -File $pub
    Write-Host ("[{0}] Concluido (veja publicar_log.txt)." -f (Get-Date).ToString('HH:mm:ss')) -ForegroundColor Green
  }
}
