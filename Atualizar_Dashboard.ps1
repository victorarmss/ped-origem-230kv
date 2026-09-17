# ==========================================================================
#  ACOMPANHAMENTO GERAL DE PROJETOS - Atualizador do Dashboard (1 clique)
#  Lê a planilha PED - ORIGEM 230kV.xlsx e regrava o dashboard_pilar.html
#  com os dados atuais. Depois abre o dashboard no navegador.
# ==========================================================================
param([switch]$NoOpen)
$ErrorActionPreference='Stop'
$base = $PSScriptRoot
$xlsx = Join-Path $env:USERPROFILE 'OneDrive - ENGETECNICA\EPC\1700-1701-1702 - ORIGEM\02. Engenharia\01. PED\PED - ORIGEM 230kV.xlsx'
$html = Join-Path $base 'index.html'
if(-not (Test-Path $xlsx)){ Write-Host "ERRO: nao encontrei '$xlsx'" -ForegroundColor Red; return }
if(-not (Test-Path $html)){ Write-Host "ERRO: nao encontrei '$html'" -ForegroundColor Red; return }

Write-Host "Lendo a planilha..." -ForegroundColor Cyan
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false
try{
  $wb = $xl.Workbooks.Open($xlsx,$false,$true)   # read-only
  $ws = $wb.Worksheets.Item('LISTA GERAL')
  $end = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
  # Col1=obra Col2=disc Col4=descrição Col6=baseline Col8=contratual Col9=reprogramado(DATA PLANEJADA REVISÃO) Col12=STATUS Col14=1a emissão (ITEM removido; Col13=ULTIMA OBSERVACAO)
  $v = $ws.Range($ws.Cells(5,1),$ws.Cells($end,14)).Value2
  $list = New-Object System.Collections.ArrayList
  for($i=1;$i -le ($end-4);$i++){
    $a=$v.GetValue($i,1); if($null -eq $a -or "$a".Trim() -eq ''){continue}
    $b=$v.GetValue($i,6); $p=$v.GetValue($i,8); $e=$v.GetValue($i,14); $s="$($v.GetValue($i,12))"; $rp=$v.GetValue($i,9)
    [void]$list.Add([pscustomobject]@{
      o="$a"; d="$($v.GetValue($i,2))"; n="$($v.GetValue($i,4))"
      p= if($p -is [double]){[int]$p}else{$null}
      b= if($b -is [double]){[int]$b}else{$null}
      e= if($e -is [double]){[int]$e}else{$null}
      rp= if($rp -is [double]){[int]$rp}else{$null}
      s=$s })
  }
  $wb.Close($false)
}finally{
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
$json = $list | ConvertTo-Json -Compress
$emit = @($list | Where-Object { $_.e -ne $null }).Count

# injeta o JSON no lugar de "const DOCS=...;" (robusto, sem regex)
$c = [System.IO.File]::ReadAllText($html,[System.Text.Encoding]::UTF8)
$i1 = $c.IndexOf('const DOCS=')
$i2 = $c.IndexOf('(function()', $i1)
if($i1 -lt 0 -or $i2 -lt 0){ Write-Host "ERRO: marcador 'const DOCS=' nao encontrado no HTML." -ForegroundColor Red; return }
$c = $c.Substring(0,$i1) + 'const DOCS=' + $json + ";`r`n" + $c.Substring($i2)
$enc = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($html,$c,$enc)

Write-Host ("OK! Dashboard atualizado: {0} documentos, {1} ja emitidos." -f $list.Count,$emit) -ForegroundColor Green
if(-not $NoOpen){
  Write-Host "Abrindo no navegador..." -ForegroundColor Cyan
  Start-Process $html
  Write-Host ""
  Write-Host "Para gerar um LINK PUBLICO: arraste o arquivo index.html em https://app.netlify.com/drop" -ForegroundColor Yellow
}
