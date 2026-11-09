$ErrorActionPreference='Stop'
try {
    $root=Split-Path -Parent $MyInvocation.MyCommand.Path
    . (Join-Path $root 'ExportExcel.ps1')
    $dest=Join-Path $root 'KhaoSat'
    Write-Reports -rawFolder (Join-Path $dest 'DuLieuGoc') -outFolder $dest
} catch { Write-Host ('Loi: '+$_.Exception.Message) -ForegroundColor Red; exit 1 }
