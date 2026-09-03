# ==========================================================================
#  Publicar_Auto.ps1
#  Regenera o index.html a partir da PED e envia ao GitHub (git push).
#  O GitHub Pages, ligado ao repositorio, publica sozinho em seguida.
#  Uso: e chamado automaticamente pelo Vigiar_PED.ps1 (ao salvar a PED),
#  ou manualmente: clique com o botao direito > "Executar com PowerShell".
# ==========================================================================
$ErrorActionPreference = 'Continue'
$base = $PSScriptRoot
$log  = Join-Path $base 'publicar_log.txt'
function Log($m){ Add-Content -LiteralPath $log -Value ("[{0}] {1}" -f (Get-Date).ToString('dd/MM/yyyy HH:mm:ss'), $m) -Encoding UTF8 }

Log "==== Iniciando publicacao ===="
try {
  # 1) Regenera o index.html a partir da PED (sem abrir o navegador)
  & (Join-Path $base 'Atualizar_Dashboard.ps1') -NoOpen 2>&1 | ForEach-Object { Log ("updater: " + $_) }
} catch {
  Log ("ERRO ao regenerar o index.html: " + $_.Exception.Message)
  return
}

# 2) Envia ao GitHub
Set-Location -LiteralPath $base
if (-not (Test-Path (Join-Path $base '.git'))) { Log "AVISO: a pasta ainda nao e um repositorio git. Veja o README.md."; return }

git add index.html 2>&1 | Out-Null
git diff --cached --quiet index.html
if ($LASTEXITCODE -eq 0) { Log "Sem mudancas no index.html - nada a publicar."; return }

$msg = "Atualiza dashboard - " + (Get-Date).ToString('dd/MM/yyyy HH:mm')
(git commit -m $msg 2>&1 | Out-String) | ForEach-Object { Log ("commit: " + $_.Trim()) }
$push = git push 2>&1 | Out-String
Log ("push: " + $push.Trim())
if ($LASTEXITCODE -eq 0) { Log "OK - enviado ao GitHub. O GitHub Pages publica em ~1-2 min." }
else { Log "ERRO no git push (verifique conexao / repositorio remoto)." }
