# READ-ONLY security inventory. Never launches suspicious software or changes settings.
# Windows PowerShell 2.0+ friendly (Windows 7 / 8 / 8.1 / 10 / 11).
if (-not (Get-Command Read-Utf8Csv -ErrorAction SilentlyContinue)) {
    . (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'CsvCompat.ps1')
}
function SecLog([string]$text,[string]$level='INFO') {
    if (Get-Command SurveyLog -ErrorAction SilentlyContinue) { SurveyLog $text $level }
    else { Write-Host ('['+$level+'] '+$text) }
}
function SecClean([object]$value) {
    if ($null -eq $value) { return '' }
    return ([string]$value).Trim() -replace '[\r\n\t]+',' '
}
function SecUnique($arr) {
    return (@($arr | Where-Object { $_ -and ([string]$_).Trim() } | Select-Object -Unique) -join ' | ')
}
function GetInstalledAppList {
    # Registry uninstall keys are much safer than querying Win32_Product (which can trigger MSI repairs).
    $roots=@(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    $items=@()
    foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }
        try {
            foreach($key in @(Get-ChildItem -Path $root -ErrorAction Stop)) {
                try {
                    $r=Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction Stop
                    if ($r.DisplayName -and ([string]$r.DisplayName).Trim()) {
                        $items += (New-Object PSObject -Property @{ Name=(SecClean $r.DisplayName); Version=(SecClean $r.DisplayVersion) })
                    }
                } catch {}
            }
        } catch { SecLog ('Khong doc duoc 1 nhanh installed apps: '+$root) 'WARN' }
    }
    return @($items | Sort-Object Name,Version -Unique)
}
function GetSuspiciousFindings([string]$rulesPath) {
    # Run this scan in a separate PowerShell process. The parent never waits
    # indefinitely for Win7/8/10/11 service enumeration or regex matching.
    $script:SuspiciousScanState='CHUA_BAT_DAU'
    $script:SuspectInstalledApps=@()
    $limitSeconds=25
    $worker=Join-Path (Split-Path -Parent $rulesPath) 'SuspiciousScanWorker.ps1'
    if(-not (Test-Path -LiteralPath $worker)) {
        $script:SuspiciousScanState='LOI_THIEU_FILE_SuspiciousScanWorker.ps1'
        SecLog 'Thieu SuspiciousScanWorker.ps1. Khong the ket luan may sach.' 'WARN'
        return @()
    }
    $folder=Join-Path $env:TEMP ('V5_SusScan_'+[guid]::NewGuid().ToString('N'))
    $process=$null
    try {
        [void](New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop)
        $inputFile=Join-Path $folder 'installed.csv'
        $outputFile=Join-Path $folder 'matches.csv'
        $stageFile=Join-Path $folder 'stage.txt'
        $ps=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        if(Test-Path (Join-Path $env:SystemRoot 'Sysnative\WindowsPowerShell\v1.0\powershell.exe')) {
            $ps=Join-Path $env:SystemRoot 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'
        }
        if(-not (Test-Path -LiteralPath $ps)) { throw 'Khong tim thay Windows PowerShell.' }
        $psi=New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName=$ps
        $psi.Arguments='-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "'+$worker+'" -InputPath "'+$inputFile+'" -RulesPath "'+$rulesPath+'" -OutputPath "'+$outputFile+'" -StatusPath "'+$stageFile+'"'
        $psi.WorkingDirectory=(Split-Path -Parent $worker)
        $psi.UseShellExecute=$false
        $psi.CreateNoWindow=$true
        $process=[System.Diagnostics.Process]::Start($psi)
        if($null -eq $process) { throw 'Khong tao duoc tien trinh quet.' }
        SecLog ('Quet phan mem nghi van (toi da '+$limitSeconds+' giay).') 'INFO'
        # Show live stage updates rather than freezing the console silently.
        $started=Get-Date
        $completed=$false
        $lastStage=''
        do {
            $completed=$process.WaitForExit(1250)
            try {
                if(Test-Path -LiteralPath $stageFile) {
                    $currentStage=Get-Content -LiteralPath $stageFile -ErrorAction Stop | Select-Object -Last 1
                    if($currentStage -and $currentStage -ne $lastStage) {
                        $lastStage=$currentStage
                        SecLog ('Quet phan mem: '+$currentStage) 'INFO'
                    }
                }
            } catch {}
        } while (-not $completed -and ((Get-Date)-$started).TotalSeconds -lt $limitSeconds)
        if(-not $completed) {
            try {$process.Kill(); [void]$process.WaitForExit(3000)} catch {}
        }
        # A slow later stage must not erase successfully collected app names.
        try {
            if(Test-Path -LiteralPath $inputFile) {
                $script:SuspectInstalledApps=@(Read-Utf8Csv -FilePath $inputFile)
            }
        } catch { SecLog ('Khong doc duoc file app tam: '+(SecClean $_.Exception.Message)) 'WARN' }
        if(-not $completed) {
            $stage='KHONG_RO'
            try { if(Test-Path -LiteralPath $stageFile) {$stage=(Get-Content -LiteralPath $stageFile -ErrorAction Stop | Select-Object -Last 1)} } catch {}
            $script:SuspiciousScanState='QUET_QUA_THOI_GIAN ('+$limitSeconds+'s; '+$stage+')'
            SecLog ('QUA THOI GIAN: quet phan mem nghi van, cong doan: '+$stage+'. Tiep tuc kiem tra phan khac.') 'WARN'
            return @()
        }
        if($process.ExitCode -ne 0) {
            $stage='KHONG_RO'
            try { if(Test-Path -LiteralPath $stageFile) {$stage=(Get-Content -LiteralPath $stageFile | Select-Object -Last 1)} } catch {}
            $script:SuspiciousScanState='LOI_QUET_PHAN_MEM: '+$stage
            SecLog ('Loi o buoc quet phan mem: '+$stage+'. Khong ket luan may sach.') 'WARN'
            return @()
        }
        $script:SuspiciousScanState='HOAN_TAT'
        if(Test-Path -LiteralPath $outputFile) {
            return @(Read-Utf8Csv -FilePath $outputFile)
        }
        return @()
    } catch {
        $script:SuspiciousScanState='LOI_QUET_PHAN_MEM: '+(SecClean $_.Exception.Message)
        SecLog ('Khong the quet phan mem nghi van: '+(SecClean $_.Exception.Message)) 'WARN'
        return @()
    } finally {
        if($process) {
            try { if(-not $process.HasExited) {$process.Kill()} } catch {}
            try {$process.Dispose()} catch {}
        }
        try { if(Test-Path -LiteralPath $folder) {Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction SilentlyContinue} } catch {}
    }
}
function GetAVSummary_Legacy($apps) {
    $r=New-Object PSObject -Property @{
        Names='KHONG_XAC_DINH'; RealTime='CAN_XAC_MINH'; Date='CAN_XAC_MINH'
        Source=''; DefenderSig=''; DefenderState='CAN_XAC_MINH'
        Kaspersky='KHONG_PHAT_HIEN'; Note=''
    }
    $registered=@()
    try { $registered=@(Get-WmiObject -Namespace 'root\SecurityCenter2' -Class AntiVirusProduct -ErrorAction Stop);Write-RawEvidence 'SecurityCenter2_AntiVirusProduct' 'Get-WmiObject -Namespace root\SecurityCenter2 -Class AntiVirusProduct' $registered }
    catch { $r.Note='Windows Security Center khong truy cap duoc; co the do quyen hoac WMI.' }
    $names=@($registered | ForEach-Object { SecClean $_.displayName } | Where-Object { $_ })
    $kapps=@($apps | Where-Object { $_.Name -match '(?i)Kaspersky' })
    $knames=@($names | Where-Object { $_ -match '(?i)Kaspersky' })
    if ($kapps.Count -gt 0 -or $knames.Count -gt 0) {
        $r.Kaspersky='CO_CAI / CAN_XAC_MINH_LICENSE'
        if($kapps.Count -gt 0) { $r.Kaspersky='CO_CAI '+(SecUnique @($kapps | ForEach-Object { $_.Version }))+' / CAN_XAC_MINH_LICENSE' }
    }
    $def=$null
    $defQueryError=''
    if (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue) {
        try { $def=Get-MpComputerStatus -ErrorAction Stop;Write-RawEvidence 'Defender_Status' 'Get-MpComputerStatus' $def }
        catch { $defQueryError=SecClean $_.Exception.Message; SecLog ('Defender query: '+$defQueryError) 'WARN' }
    } else { $defQueryError='Get-MpComputerStatus khong co tren may' }
    if($def) {
        if($def.AntivirusSignatureLastUpdated) {
            try { $r.DefenderSig=([datetime]$def.AntivirusSignatureLastUpdated).ToString('yyyy-MM-dd HH:mm') }
            catch { $r.DefenderSig=SecClean $def.AntivirusSignatureLastUpdated }
        }
        if ($def.RealTimeProtectionEnabled -eq $true -and $def.AntivirusEnabled -eq $true) { $r.DefenderState='ACTIVE' }
        elseif ($def.RealTimeProtectionEnabled -eq $false -or $def.AntivirusEnabled -eq $false) { $r.DefenderState='OFF_OR_PASSIVE' }
    }
    $namesFromRegistry=$false
    if($names.Count -eq 0) {
        if($kapps.Count -gt 0) { $names=@('Kaspersky (installed; active status unverified)'); $namesFromRegistry=$true }
        elseif($def) {$names=@('Microsoft Defender Antivirus')}
    }
    if($names.Count -gt 0) {
        $r.Names=SecUnique $names
        if($namesFromRegistry){$r.Source='Registry Uninstall (da cai; chua xac minh active)'}
        elseif($registered.Count -gt 0){$r.Source='Windows Security Center (dang ky; chua xac minh active)'}
        else{$r.Source='Get-MpComputerStatus (Defender)'}
    }
    $otherAV=@($names | Where-Object { $_ -notmatch '(?i)Defender|Microsoft Security Essentials' })
    if($otherAV.Count -eq 0 -and $def) {
        $r.RealTime= if ($r.DefenderState -eq 'ACTIVE') {'ACTIVE (Defender)'} elseif($r.DefenderState -eq 'OFF_OR_PASSIVE') {'OFF/PASSIVE (Defender)'} else {'CAN_XAC_MINH'}
        $r.Date= if($r.DefenderSig){$r.DefenderSig}else{'CAN_XAC_MINH'}
        $r.Source='Get-MpComputerStatus'
    } elseif ($otherAV.Count -gt 0) {
        # The productState integer has no reliable cross-vendor public specification; do not guess its meaning.
        $r.RealTime='CAN_XAC_MINH_TRONG_UNG_DUNG'
        $r.Date='CAN_XAC_MINH_TRONG_UNG_DUNG'
        $r.Note+=' | AV ben thu ba: Windows khong cung cap ngay signature cap nhat chung, xem ung dung/console quan tri.'
    } else {
        $r.RealTime='CAN_XAC_MINH'
        $r.Date='CAN_XAC_MINH'
    }
    if($defQueryError) { $r.Note+=' | Defender: '+$defQueryError+' (khong phai ket luan AV dang tat)' }
    if($names.Count -gt 1) { $r.Note+=' | Co nhieu AV dang ky; can doi chieu san pham dang bao ve.' }
    return $r
}
# AV V5 universal read-only module.
. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'AntivirusAudit.ps1')
function GetAVSummary($apps) { return (GetAVFullSummary $apps) }
function GetLocalAdminAudit {
    $result=New-Object PSObject -Property @{ State='CAN_XAC_MINH'; Elevated='CAN_XAC_MINH' }
    try {
        $identity=[System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal=New-Object System.Security.Principal.WindowsPrincipal($identity)
        $elevated=$principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
        $result.Elevated=if($elevated){'YES'}else{'NO'}
        $sids=@($identity.Groups | ForEach-Object { $_.Value })
        if($elevated) { $result.State='LOCAL_ADMIN_ELEVATED' }
        elseif($sids -contains 'S-1-5-32-544') {
            # With UAC, the Administrators SID can be deny-only in a filtered token.
            $result.State='ADMIN_GROUP_TOKEN_FILTERED_UAC'
        } else { $result.State='STANDARD_USER_HOAC_CHUA_XAC_MINH' }
    } catch { SecLog ('Local Admin: '+$_.Exception.Message) 'WARN' }
    return $result
}
function SecClassifyError([object]$errorObject) {
    $msg=''
    try {$msg=[string]$errorObject.Exception.Message} catch {$msg=[string]$errorObject}
    $code=''
    try { $code=('0x{0:X8}' -f ([uint32]$errorObject.Exception.HResult)) } catch {}
    if($msg -match '(?i)access is denied|access denied|privilege|administrator|0x80070005|0x80041003|unauthorized|permission') {
        return 'THIEU_QUYEN_ADMIN'
    }
    if($msg -match '(?i)invalid namespace|invalid class|not found|0x8004100e|0x80041010') {
        return 'WMI_CLASS_HOAC_NAMESPACE_KHONG_CO'
    }
    if($code -eq '0x80070005' -or $code -eq '0x80041003') { return 'THIEU_QUYEN_ADMIN' }
    if($code -eq '0x8004100E' -or $code -eq '0x80041010') { return 'WMI_CLASS_HOAC_NAMESPACE_KHONG_CO' }
    return 'LOI_TRUY_VAN_WMI'
}
function GetTpmAudit {
    $r=New-Object PSObject -Property @{ State='CAN_XAC_MINH'; Spec=''; Detail=''; Source='Win32_Tpm WMI' }
    try {
        $t=@(Get-WmiObject -Namespace 'root\CIMV2\Security\MicrosoftTpm' -Class Win32_Tpm -ErrorAction Stop);Write-RawEvidence 'Win32_Tpm' 'Get-WmiObject -Namespace root\CIMV2\Security\MicrosoftTpm -Class Win32_Tpm' $t
        if($t.Count -gt 0) {
            $r.Spec=SecClean $t[0].SpecVersion
            if($r.Spec -match '(^|[,;\s])2\.0([,;\s]|$)') { $r.State='TPM_2.0' }
            elseif($r.Spec -match '(^|[,;\s])1\.2([,;\s]|$)') { $r.State='TPM_1.2' }
            else { $r.State='TPM_CO_NHUNG_CHUA_RO_VERSION' }
        } else { $r.State='KHONG_PHAT_HIEN_TPM_QUA_WMI'; $r.Detail='Kiem tra BIOS/UEFI, TPM firmware co the dang tat.' }
    } catch {
        $r.State=SecClassifyError $_
        $r.Detail=SecClean $_.Exception.Message
        # Get-Tpm may exist on Windows 8 and later but usually not Windows 7.
        # It reports presence/ready, NOT the TPM 2.0 spec version.
        if(Get-Command Get-Tpm -ErrorAction SilentlyContinue) {
            try {
                $tpm=Get-Tpm -ErrorAction Stop;Write-RawEvidence 'Get_Tpm' 'Get-Tpm' $tpm
                if($tpm.TpmPresent -eq $true) { $r.State='TPM_CO_NHUNG_CHUA_RO_VERSION'; $r.Source='Get-Tpm fallback' }
                elseif($tpm.TpmPresent -eq $false) { $r.State='KHONG_PHAT_HIEN_TPM'; $r.Source='Get-Tpm fallback' }
            } catch { $r.Detail += ' | Get-Tpm: '+(SecClean $_.Exception.Message) }
        }
    }
    return $r
}
function GetBitlockerAudit {
    # GetConversionStatus and GetProtectionStatus must be interpreted TOGETHER.
    # ProtectionStatus=0 by itself NEVER proves that the volume is decrypted.
    $r=New-Object PSObject -Property @{
        State='CAN_XAC_MINH'; Conversion=''; Protection=''; Percent='';
        Detail=''; Source='Win32_EncryptableVolume WMI'
    }
    try {
        $drive=([string]$env:SystemDrive).TrimEnd('\')
        $volumes=@(Get-WmiObject -Namespace 'root\CIMV2\Security\MicrosoftVolumeEncryption' -Class Win32_EncryptableVolume -ErrorAction Stop);Write-RawEvidence 'BitLocker_EncryptableVolumes' 'Get-WmiObject -Namespace root\CIMV2\Security\MicrosoftVolumeEncryption -Class Win32_EncryptableVolume' $volumes
        $v=@($volumes | Where-Object { ([string]$_.DriveLetter).TrimEnd('\') -eq $drive } | Select-Object -First 1)
        if($v.Count -eq 0) {
            $r.State='KHONG_TIM_THAY_O_WINDOWS_TRONG_WMI'
            $r.Detail='SystemDrive='+$drive
            return $r
        }
        $protect=$v[0].GetProtectionStatus()
        if([int64]$protect.ReturnValue -ne 0) {
            $r.State='LOI_WMI_PROTECTION_STATUS'
            $r.Detail='GetProtectionStatus ReturnValue='+$protect.ReturnValue
            return $r
        }
        try { $convert=$v[0].GetConversionStatus(0) }
        catch { $convert=$v[0].GetConversionStatus() }
        if([int64]$convert.ReturnValue -ne 0) {
            $r.State='LOI_WMI_CONVERSION_STATUS'
            $r.Detail='GetConversionStatus ReturnValue='+$convert.ReturnValue
            return $r
        }
        $p=[int]$protect.ProtectionStatus
        $c=[int]$convert.ConversionStatus
        $r.Percent=[string]$convert.EncryptionPercentage
        switch($p) {
            0 { $r.Protection='PROTECTION_OFF' }
            1 { $r.Protection='PROTECTION_ON' }
            2 { $r.Protection='UNKNOWN' }
            default { $r.Protection='UNKNOWN_'+$p }
        }
        switch($c) {
            0 { $r.Conversion='FULLY_DECRYPTED' }
            1 { $r.Conversion='FULLY_ENCRYPTED' }
            2 { $r.Conversion='ENCRYPTION_IN_PROGRESS' }
            3 { $r.Conversion='DECRYPTION_IN_PROGRESS' }
            4 { $r.Conversion='ENCRYPTION_PAUSED' }
            5 { $r.Conversion='DECRYPTION_PAUSED' }
            default { $r.Conversion='UNKNOWN_'+$c }
        }
        switch($c) {
            0 {
                if($p -eq 0) { $r.State='OFF - FULLY DECRYPTED' }
                else { $r.State='CAN_XAC_MINH - DECRYPTED / PROTECTION '+$p }
            }
            1 {
                if($p -eq 1) { $r.State='ON - FULLY ENCRYPTED' }
                elseif($p -eq 0) { $r.State='PROTECTION OFF - ENCRYPTED' }
                else { $r.State='CAN_XAC_MINH - ENCRYPTED / UNKNOWN PROTECTION' }
            }
            2 { $r.State='ENCRYPTION IN PROGRESS ('+$r.Percent+'%)' }
            3 { $r.State='DECRYPTION IN PROGRESS ('+$r.Percent+'%)' }
            4 { $r.State='ENCRYPTION PAUSED ('+$r.Percent+'%)' }
            5 { $r.State='DECRYPTION PAUSED ('+$r.Percent+'%)' }
            default { $r.State='CAN_XAC_MINH - CONVERSION STATUS '+$c }
        }
        if($c -eq 1 -and $p -eq 0) {
            $r.Detail='O da ma hoa nhung khoa bao ve OFF; co the suspend. KHONG ket luan suspend neu chua co them thong tin.'
        }
        return $r
    } catch {
        $r.State=SecClassifyError $_
        $r.Detail=SecClean $_.Exception.Message
        # manage-bde exists on some editions. Keep the fallback read-only.
        # Its output is localized, so never derive an encryption state from text
        # unless exact English status labels can be recognized together.
        $manage=Join-Path $env:SystemRoot 'System32\manage-bde.exe'
        if(Test-Path -LiteralPath $manage) {
            try {
                $raw=(& $manage -status $env:SystemDrive 2>&1 | Out-String)
                Write-RawCommandText 'BitLocker_ManageBDE' ($manage+' -status '+$env:SystemDrive) $raw '' 'OK'
                $conv=[regex]::Match($raw,'(?im)^\s*Conversion Status\s*:\s*(.+)$')
                $prot=[regex]::Match($raw,'(?im)^\s*Protection Status\s*:\s*(.+)$')
                $pct=[regex]::Match($raw,'(?im)^\s*Percentage Encrypted\s*:\s*([\d\.,]+)')
                if($conv.Success -and $prot.Success) {
                    $r.Source='manage-bde -status (English output)'
                    $r.Conversion=(SecClean $conv.Groups[1].Value).ToUpper().Replace(' ','_')
                    $r.Protection=(SecClean $prot.Groups[1].Value).ToUpper().Replace(' ','_')
                    if($pct.Success) {$r.Percent=SecClean $pct.Groups[1].Value}
                    $cv=$r.Conversion; $ps=$r.Protection
                    if($cv -eq 'FULLY_DECRYPTED' -and $ps -eq 'PROTECTION_OFF') {$r.State='OFF - FULLY DECRYPTED'}
                    elseif($cv -eq 'FULLY_ENCRYPTED' -and $ps -eq 'PROTECTION_ON') {$r.State='ON - FULLY ENCRYPTED'}
                    elseif($cv -eq 'FULLY_ENCRYPTED' -and $ps -eq 'PROTECTION_OFF') {$r.State='PROTECTION OFF - ENCRYPTED'}
                    elseif($cv -eq 'ENCRYPTION_IN_PROGRESS') {$r.State='ENCRYPTION IN PROGRESS ('+$r.Percent+'%)'}
                    elseif($cv -eq 'DECRYPTION_IN_PROGRESS') {$r.State='DECRYPTION IN PROGRESS ('+$r.Percent+'%)'}
                    elseif($cv -eq 'ENCRYPTION_PAUSED') {$r.State='ENCRYPTION PAUSED ('+$r.Percent+'%)'}
                    elseif($cv -eq 'DECRYPTION_PAUSED') {$r.State='DECRYPTION PAUSED ('+$r.Percent+'%)'}
                }
            } catch { $r.Detail+=' | manage-bde: '+(SecClean $_.Exception.Message) }
        }
        return $r
    }
}
function GetSecurityAudit([string]$rulesPath) {
    SecLog 'Dang doc ung dung / process / services va doi chieu dau hieu (watchdog 25s)...' 'STEP'
    $findings=@(GetSuspiciousFindings -rulesPath $rulesPath)
    $apps=@($script:SuspectInstalledApps)
    SecLog ('Da doc '+$apps.Count+' ten ung dung. Trang thai quet: '+$script:SuspiciousScanState) 'INFO'
    $cracks=@($findings | Where-Object { $_.Category -eq 'ACTIVATOR_PATCH' })
    $remote=@($findings | Where-Object { $_.Category -eq 'REMOTE_CONTROL' })
    foreach($f in $findings) { SecLog ('Dau hieu '+$f.Category+': '+$f.Label+' ['+$f.Source+': '+$f.Evidence+']') 'WARN' }
    SecLog 'Dang kiem tra antivirus va ngay cap nhat...' 'STEP'
    $av=GetAVSummary $apps
    if($script:SuspiciousScanState -ne 'HOAN_TAT') {
        # The installed-app list may be incomplete when the bounded scan ends early.
        if($av.Kaspersky -eq 'KHONG_PHAT_HIEN') { $av.Kaspersky='CHUA_XAC_MINH_DO_QUET_PHAN_MEM_CHUA_HOAN_TAT' }
        $av.Note=$av.Note+' | Danh sach phan mem co the chua day du: '+$script:SuspiciousScanState
    }
    SecLog ('AV: '+$av.Names+' | Real-time: '+$av.RealTime+' | Signature: '+$av.Date) 'INFO'
    SecLog 'Dang kiem tra Local Admin, TPM va BitLocker...' 'STEP'
    $admin=GetLocalAdminAudit
    $tpm=GetTpmAudit
    $bitlocker=GetBitlockerAudit
    SecLog ('BitLocker: '+$bitlocker.State+' | Conversion: '+$bitlocker.Conversion+' | Protection: '+$bitlocker.Protection+' | Percent: '+$bitlocker.Percent+' | Source: '+$bitlocker.Source) $(if($bitlocker.State -like 'CAN_XAC_MINH*' -or $bitlocker.State -like 'THIEU_QUYEN*' -or $bitlocker.State -like 'LOI_*'){'WARN'}else{'INFO'})
    if($bitlocker.Detail){ SecLog ('BitLocker chi tiet: '+$bitlocker.Detail) 'WARN' }
    if($tpm.Detail){ SecLog ('TPM chi tiet: '+$tpm.Detail) 'WARN' }
    $suspect='KHONG_PHAT_HIEN_DAU_HIEU_TRONG_DANH_SACH'
    if($script:SuspiciousScanState -ne 'HOAN_TAT') { $suspect=$script:SuspiciousScanState }
    if($cracks.Count -gt 0) { $suspect='CAN_RA_SOAT: '+(SecUnique @($cracks | ForEach-Object { $_.Label })) }
    $remoteDesc=if($remote.Count -gt 0){SecUnique @($remote | ForEach-Object { $_.Label })}elseif($script:SuspiciousScanState -ne 'HOAN_TAT'){'CHUA_KIEM_TRA - '+$script:SuspiciousScanState}else{'KHONG_PHAT_HIEN_TRONG_DANH_SACH'}
    $details=@($findings | ForEach-Object { '['+$_.Category+']['+$_.Source+'] '+$_.Label+': '+$_.Evidence })
    return (New-Object PSObject -Property @{
        Suspicious=$suspect; Findings=(SecUnique $details); Remote=$remoteDesc; Admin=$admin.State
        AdminElevated=$admin.Elevated; Antivirus=$av.Names; AVRealTime=$av.RealTime
        AVUpdated=$av.Date; AVSource=$av.Source; AVNote=$av.Note; DefenderSignature=$av.DefenderSig
        AVVersions=$av.Versions; AVRealTimeDetails=$av.RealTimeDetail;
        AVUpdateDetails=$av.UpdateDetail; AVServiceDetails=$av.Services;
        AVLicenseDetails=$av.LicenseDetail; AVEvidenceDetails=$av.EvidenceDetail;
        AVProductCount=$av.ProductCount; AVScanDetails=$av.ScanDetail; AVEngineDetails=$av.EngineDetail
        Kaspersky=$av.Kaspersky; TPM=$tpm.State; TPMSpec=$tpm.Spec;
        TPMDetail=$tpm.Detail; TPMSource=$tpm.Source;
        BitLocker=$bitlocker.State; BitLockerConversion=$bitlocker.Conversion;
        BitLockerProtection=$bitlocker.Protection; BitLockerPercent=$bitlocker.Percent;
        BitLockerDetail=$bitlocker.Detail; BitLockerSource=$bitlocker.Source
    })
}
