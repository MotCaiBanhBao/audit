# ParallelAudit.ps1: orchestrates THREE independent read-only tasks.
# Windows PowerShell 2.0+; LICENSE has a watchdog so the survey cannot hang.
# Only the parent Survey.ps1 writes KhaoSat\DuLieuGoc and Inventory.xlsx.
function Invoke-ParallelAudit([string]$Root,[scriptblock]$OnLicenseTimeout,[string]$RawOutputDir) {
    $result=New-Object PSObject -Property @{
        License=$null;Security=$null;Extra=$null;Compliance=$null;
        Apps=@(); Errors=@(); Mode='PARALLEL'; Timings=@{}; OfficeFailed=$false; WindowsFailed=$false
    }
    $workDir=Join-Path $env:TEMP ('V5_Audit_'+[guid]::NewGuid().ToString('N'))
    $jobs=@()
    $failed=$false
    $worker=Join-Path $Root 'ParallelAuditWorker.ps1'
    # Separate bounded workers: Office stalls cannot overwrite Windows results.
    $licenseMaxSec=40
    $windowsMaxSec=30
    try {
        if ($env:AUDIT_LICENSE_TIMEOUT_SEC -and
            [int]$env:AUDIT_LICENSE_TIMEOUT_SEC -ge 10 -and
            [int]$env:AUDIT_LICENSE_TIMEOUT_SEC -le 120) {
            $licenseMaxSec=[int]$env:AUDIT_LICENSE_TIMEOUT_SEC
        }
    } catch { $licenseMaxSec=40 }
    try {
        if($env:AUDIT_WINDOWS_TIMEOUT_SEC -and [int]$env:AUDIT_WINDOWS_TIMEOUT_SEC -ge 8 -and [int]$env:AUDIT_WINDOWS_TIMEOUT_SEC -le 90){
            $windowsMaxSec=[int]$env:AUDIT_WINDOWS_TIMEOUT_SEC
        }
    }catch{$windowsMaxSec=30}
    try {
        if(-not (Test-Path -LiteralPath $worker)){throw 'Khong tim thay ParallelAuditWorker.ps1'}
        [void](New-Item -ItemType Directory -Path $workDir -Force -ErrorAction Stop)
        $ps=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $sysNative=Join-Path $env:SystemRoot 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'
        if(Test-Path -LiteralPath $sysNative){$ps=$sysNative}
        if(-not (Test-Path -LiteralPath $ps)){throw 'Windows PowerShell executable not found'}
        foreach($name in @('WINDOWS','OFFICE','SECURITY','POLICIES')){
            $psi=New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName=$ps
            $psi.Arguments='-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "'+$worker+'" -Task '+$name+' -JobDir "'+$workDir+'" -Root "'+$Root+'"'
            $psi.UseShellExecute=$false
            $psi.CreateNoWindow=$true
            $psi.WorkingDirectory=$Root
            $p=[System.Diagnostics.Process]::Start($psi)
            if($null -eq $p){throw ('Khong khoi dong duoc tien trinh '+$name)}
            $job=New-Object PSObject -Property @{
                Name=$name; Proc=$p; ReadLines=0; Completed=$false;
                Began=(Get-Date); Seconds=''; HadError=$false
            }
            $jobs+=,$job
            SurveyLog ('[PARALLEL] Da bat dau '+$name+' (PID '+$p.Id+')') 'STEP'
        }
        # LICENSE watchdog prevents indefinite waits. All other workers continue as before.
        # Every worker logs to its own TEMP file; parent relays log to USB log.
        $nextHeartbeat=(Get-Date).AddSeconds(15)
        while(@($jobs | Where-Object {-not $_.Completed}).Count -gt 0){
            foreach($job in $jobs){
                $logFile=Join-Path $workDir ($job.Name+'.log')
                try {
                    if(Test-Path -LiteralPath $logFile){
                        $lines=@(Get-Content -LiteralPath $logFile -Encoding UTF8 -ErrorAction Stop)
                        if($lines.Count -gt $job.ReadLines){
                            for($i=$job.ReadLines;$i -lt $lines.Count;$i++){
                                $textLine=[string]$lines[$i]
                                $level='INFO'
                                if($textLine -match '\[ERROR\]'){$level='ERROR'}
                                elseif($textLine -match '\[WARN\]'){$level='WARN'}
                                elseif($textLine -match '\[OK\]'){$level='OK'}
                                SurveyLog $textLine $level
                            }
                            $job.ReadLines=$lines.Count
                        }
                    }
                }catch{}
                $maxSec=0
                if($job.Name -eq 'WINDOWS'){$maxSec=$windowsMaxSec}
                elseif($job.Name -eq 'OFFICE'){$maxSec=$licenseMaxSec}
                if(-not $job.Completed -and $maxSec -gt 0 -and
                   -not $job.Proc.HasExited -and
                   ((Get-Date)-$job.Began).TotalSeconds -ge $maxSec) {
                    # Office licensing providers can freeze. Stop only LICENSE;
                    # do NOT cancel security/policies or restart license in Survey.
                    try {$job.Proc.Kill()}catch{}
                    try {[void]$job.Proc.WaitForExit(1200)}catch{}
                    $job.Completed=$true
                    $job.HadError=$true
                    $failed=$true
                    $job.Seconds=[string]$maxSec
                    $result.Timings[$job.Name]=$job.Seconds
                    $result.Errors+=($job.Name+'_TIMEOUT_'+$maxSec+'s')
                    if(Get-Command Write-RawCommandText -ErrorAction SilentlyContinue){
                        Write-RawCommandText ('Worker_'+$job.Name) ('ParallelAuditWorker.ps1 -Task '+$job.Name) '' '' ('TIMEOUT_'+$maxSec+'s')
                    }
                    if($job.Name -eq 'OFFICE'){$result.OfficeFailed=$true}
                    if($job.Name -eq 'WINDOWS'){$result.WindowsFailed=$true}
                    SurveyLog ('[PARALLEL] '+$job.Name+' TIMEOUT '+$maxSec+'s; cac nhom con lai khong bi anh huong.') 'WARN'
                    # Invoke in the *interactive parent* immediately: no need
                    # to wait for SECURITY or POLICIES to finish first. The
                    # child LICENSE process is already terminated.
                    if ($job.Name -eq 'OFFICE' -and $OnLicenseTimeout) {
                        try {
                            [void](& $OnLicenseTimeout)
                        } catch {
                            $result.Errors+=('OFFICE_TIMEOUT_PROMPT_ERROR: '+$_.Exception.Message)
                            SurveyLog ('Khong hien duoc bang nhap Office: '+$_.Exception.Message) 'WARN'
                        }
                    }
                }
                if(-not $job.Completed -and $job.Proc.HasExited){
                    # Drain any final log lines on next loop iteration before completion.
                    $job.Completed=$true
                    $seconds=[math]::Round(((Get-Date)-$job.Began).TotalSeconds,1)
                    $job.Seconds=[string]$seconds
                    $result.Timings[$job.Name]=$job.Seconds
                    if($job.Proc.ExitCode -eq 0 -and (Test-Path -LiteralPath (Join-Path $workDir ($job.Name+'.xml')))){
                        SurveyLog ('[PARALLEL] '+$job.Name+' HOAN_TAT ('+$job.Seconds+'s)') 'OK'
                    }else{
                        $job.HadError=$true
                        $failed=$true
                        $result.Errors+=($job.Name+' worker failed (exit code '+$job.Proc.ExitCode+')')
                        if($job.Name -eq 'OFFICE'){$result.OfficeFailed=$true}
                        if($job.Name -eq 'WINDOWS'){$result.WindowsFailed=$true}
                        SurveyLog ('[PARALLEL] '+$job.Name+' LOI - can doi chieu; cac nhom khac van chay.') 'WARN'
                    }
                }
            }
            if(@($jobs | Where-Object {-not $_.Completed}).Count -gt 0){
                if((Get-Date) -ge $nextHeartbeat){
                    $active=@($jobs | Where-Object {-not $_.Completed} | ForEach-Object {$_.Name}) -join ', '
                    SurveyLog ('[PARALLEL] Dang cho: '+$active+'. WINDOWS '+$windowsMaxSec+'s, OFFICE '+$licenseMaxSec+'s.') 'INFO'
                    $nextHeartbeat=(Get-Date).AddSeconds(15)
                }
                Start-Sleep -Milliseconds 900
            }
        }
        # Drain one last time (worker may have logged after parent first noticed exit).
        foreach($job in $jobs){
            try {
                $logFile=Join-Path $workDir ($job.Name+'.log')
                if(Test-Path $logFile){
                    $lines=@(Get-Content -LiteralPath $logFile -Encoding UTF8)
                    for($i=$job.ReadLines;$i -lt $lines.Count;$i++){ SurveyLog ([string]$lines[$i]) 'INFO' }
                }
            }catch{}
            if($job.HadError){continue}
            try {
                $obj=Import-Clixml -Path (Join-Path $workDir ($job.Name+'.xml')) -ErrorAction Stop
                if($job.Name -eq 'WINDOWS'){$result.License=$obj}
                elseif($job.Name -eq 'OFFICE'){
                    if(-not $result.License){$result.License=$obj}
                    else {
                        foreach($property in @($obj.PSObject.Properties)){
                            if(([string]$property.Name) -match '^(Office|Word|Excel)'){
                                $result.License.PSObject.Properties[$property.Name].Value=$property.Value
                            }
                        }
                    }
                }
                elseif($job.Name -eq 'SECURITY'){
                    $result.Security=$obj
                    $appsPath=Join-Path $workDir 'SECURITY_APPS.csv'
                    if(Test-Path -LiteralPath $appsPath){
                        . (Join-Path $Root 'CsvCompat.ps1')
                        $result.Apps=@(Read-Utf8Csv -FilePath $appsPath)
                    }
                }else{
                    $result.Extra=$obj
                    $result.Compliance=Import-Clixml -Path (Join-Path $workDir 'POLICIES_COMPLIANCE.xml') -ErrorAction Stop
                }
            }catch{
                $failed=$true
                $result.Errors+=('Import '+$job.Name+': '+$_.Exception.Message)
                if($job.Name -eq 'OFFICE'){$result.OfficeFailed=$true}
                if($job.Name -eq 'WINDOWS'){$result.WindowsFailed=$true}
                SurveyLog ('Khong nap duoc ket qua '+$job.Name+': '+$_.Exception.Message) 'WARN'
            }
        }
        $result.Timings['LICENSE']=([string]$result.Timings['WINDOWS'])+'/'+([string]$result.Timings['OFFICE'])
        if($failed){$result.Mode='PARALLEL_MOT_SO_NHOM_LOI'}
        return $result
    }catch{
        SurveyLog ('Khong the khoi dong nhom song song: '+$_.Exception.Message) 'WARN'
        $result.Errors+=('Startup: '+$_.Exception.Message)
        $result.Mode='PARALLEL_STARTUP_ERROR'
        $result.WindowsFailed=$true;$result.OfficeFailed=$true
        # If startup failed, stop orphan worker processes; do not wait forever.
        return $result
    }finally{
        foreach($job in $jobs){
            try {
                if($job.Proc -and -not $job.Proc.HasExited){
                    # Only reached on early error/interruption. Normal path waited for ALL.
                    $job.Proc.Kill()
                    [void]$job.Proc.WaitForExit(1000)
                }
            }catch{}
            try {if($job.Proc){$job.Proc.Dispose()}}catch{}
        }
        # Persist every available raw worker output, INCLUDING partial results
        # from workers killed by timeout. Parent alone writes to the USB.
        if($RawOutputDir){
            try {
                if(-not (Test-Path -LiteralPath $RawOutputDir)){[void](New-Item -ItemType Directory -Path $RawOutputDir -Force)}
                foreach($task in @('WINDOWS','OFFICE','SECURITY','POLICIES')) {
                    $from=Join-Path $workDir ('RAW_'+$task)
                    if(Test-Path -LiteralPath $from){
                        $dest=Join-Path $RawOutputDir $task
                        [void](New-Item -ItemType Directory -Path $dest -Force -ErrorAction SilentlyContinue)
                        Copy-Item -Path (Join-Path $from '*') -Destination $dest -Recurse -Force -ErrorAction SilentlyContinue
                    }
                }
                foreach($f in @(Get-ChildItem -LiteralPath $workDir -ErrorAction SilentlyContinue | Where-Object {-not $_.PSIsContainer})){
                    if($f.Extension -eq '.log' -or $f.Extension -eq '.xml'){
                        Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $RawOutputDir $f.Name) -Force -ErrorAction SilentlyContinue
                    }
                }
            }catch{SurveyLog ('Raw worker copy: '+$_.Exception.Message) 'WARN'}
        }
        # Worker directory includes only transient metadata; keep on failure for troubleshooting.
        if($result.Mode -eq 'PARALLEL'){
            try {Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue}catch{}
        }else{
            SurveyLog ('Nhat ky tien trinh phu (may local, khong USB): '+$workDir) 'WARN'
        }
    }
}
