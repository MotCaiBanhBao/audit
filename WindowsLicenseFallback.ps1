# WindowsLicenseFallback.ps1 - bounded read-only slmgr /dli fallback.
# Does not activate or install product keys; supports Windows PowerShell 2.0+.
function Get-WindowsSlmgrFallback {
    $r=New-Object PSObject -Property @{
        WindowsState='CAN_XAC_MINH_WINDOWS_LICENSE_TIMEOUT';WindowsEdition='';WindowsChannel='';
        WindowsMethod='SLMGR_DLI_FALLBACK';WindowsNote='WMI Windows khong hoan tat; can doi chieu trong output goc.'
    }
    $script=Join-Path $env:SystemRoot 'System32\slmgr.vbs'
    $native=Join-Path $env:SystemRoot 'Sysnative\slmgr.vbs'
    if(Test-Path -LiteralPath $native){$script=$native}
    if(-not (Test-Path -LiteralPath $script)){
        $r.WindowsNote='Khong tim thay slmgr.vbs de doi chieu khi Windows WMI loi/timeout.'
        Write-RawCommandText 'Windows_SLMGR_DLI' 'cscript.exe //Nologo slmgr.vbs /dli' '' $r.WindowsNote 'NOT_FOUND'
        return $r
    }
    $p=$null
    try {
        $si=New-Object System.Diagnostics.ProcessStartInfo
        $si.FileName=(Join-Path $env:SystemRoot 'System32\cscript.exe')
        if(-not (Test-Path $si.FileName)){$si.FileName='cscript.exe'}
        $si.Arguments='//Nologo "'+$script+'" /dli'
        $si.UseShellExecute=$false
        $si.CreateNoWindow=$true
        $si.RedirectStandardOutput=$true
        $si.RedirectStandardError=$true
        $p=New-Object System.Diagnostics.Process
        $p.StartInfo=$si
        if(-not $p.Start()){throw 'Khong khoi dong cscript.exe'}
        # /dli is a small text response; bounded timeout prevents blocking.
        if(-not $p.WaitForExit(12000)){
            try{$p.Kill()}catch{}
            try{[void]$p.WaitForExit(1000)}catch{}
            $r.WindowsNote='SLMGR /dli timeout 12s; Windows activation CHUA XAC MINH.'
            Write-RawCommandText 'Windows_SLMGR_DLI' ('cscript.exe //Nologo "'+$script+'" /dli') '' '' 'TIMEOUT_12s'
            return $r
        }
        $stdout=[string]$p.StandardOutput.ReadToEnd()
        $stderr=[string]$p.StandardError.ReadToEnd()
        $status='OK'
        if($p.ExitCode -ne 0){$status='EXIT_'+$p.ExitCode}
        Write-RawCommandText 'Windows_SLMGR_DLI' ('cscript.exe //Nologo "'+$script+'" /dli') $stdout $stderr $status
        if($p.ExitCode -ne 0){
            $r.WindowsNote='SLMGR /dli exit code '+$p.ExitCode+'; can doi chieu stderr trong LenhGoc.'
            return $r
        }
        # Do not guess activation from a translated/unknown phrase.
        $licensed=[regex]::IsMatch($stdout,'(?im)^\s*License Status\s*:\s*Licensed\s*$')
        $notLicensed=[regex]::IsMatch($stdout,'(?im)^\s*License Status\s*:\s*(?:Unlicensed|Notification)\s*$')
        $grace=[regex]::IsMatch($stdout,'(?im)^\s*License Status\s*:\s*.*(?:Grace|Extended Grace)')
        if($licensed){$r.WindowsState='DA_KICH_HOAT'}
        elseif($notLicensed){$r.WindowsState='CHUA_KICH_HOAT'}
        elseif($grace){$r.WindowsState='CAN_XAC_MINH_WINDOWS_LICENSE_GRACE'}
        else{$r.WindowsState='CAN_XAC_MINH_SLMGR_KHONG_NHAN_DIEN_NGON_NGU'}
        $r.WindowsNote='SLMGR /dli (nguon du phong doc lap WMI). Xem output nguyen van trong LenhGoc; trang thai kich hoat khong chung minh tinh hop le giay phep.'
        return $r
    }catch{
        $r.WindowsNote='SLMGR /dli loi: '+[string]$_.Exception.Message
        Write-RawCommandText 'Windows_SLMGR_DLI' 'cscript.exe //Nologo slmgr.vbs /dli' '' $r.WindowsNote 'ERROR'
        return $r
    }finally{
        if($p){
            try{if(-not $p.HasExited){$p.Kill()}}catch{}
            try{$p.Dispose()}catch{}
        }
    }
}
