# V5 - read-only antivirus inventory, PowerShell 2.0+ (Windows 7 to 11).
# Do not infer real-time protection or signature date from WSC productState/timestamp.
# Do not execute vendor tools unless their Authenticode signature is verified.
function AVText($x) { if($null -eq $x){return ''};return ([string]$x).Trim() -replace '[\r\n\t]+',' ' }
function AVUnique($values) {return (@($values | Where-Object {$_ -and ([string]$_).Trim()} | Select-Object -Unique) -join ' | ')}
function AVMatchName([string]$name) {
    # Exclude management agents, VPNs and firewalls which are not AV products.
    if($name -match '(?i)Kaspersky (Security Center|Network Agent|VPN|Password Manager)|Kaspersky.*(Plug-in|Management Plugin)|Windows Defender Firewall') {return $false}
    return ($name -match '(?i)Kaspersky (Endpoint Security|Anti-Virus|Internet Security|Standard|Plus|Premium|Small Office Security|Security Cloud|Total Security)|Microsoft Defender Antivirus|Windows Defender$|Microsoft Security Essentials|ReasonLabs|Reason Cybersecurity|RAV Endpoint|Avast|AVG Antivirus|AVG Anti-Virus|Bitdefender|ESET|NOD32|Norton|Symantec Endpoint|McAfee|Trellix|Trend Micro|Sophos|Panda (Dome|Security|Antivirus)|Malwarebytes|Avira|F-Secure|Webroot|SentinelOne|CrowdStrike|Cylance|Emsisoft|Comodo (Internet Security|Antivirus)|360 Total Security|Bkav|K7 (Total Security|Antivirus)|Dr\.Web|AhnLab|G DATA|Seqrite|Quick Heal|Vipre|ESET Endpoint')
}
function AVFindInstalled($scanApps) {
    $result=@()
    foreach($a in @($scanApps)) {
        if($a.Name -and (AVMatchName ([string]$a.Name))) {
            $result+= (New-Object PSObject -Property @{Name=(AVText $a.Name);Version=(AVText $a.Version);Source='Registry uninstall (app scan)'})
        }
    }
    # Reuse the app-inventory worker result whenever available to avoid scanning
    # every uninstall registry key a second time on slow Windows 7 machines.
    if(@($scanApps).Count -gt 0){return @($result | Sort-Object Name,Version -Unique)}
    # Registry fallback if the app-inventory worker did not return any entries.
    $roots=@('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
             'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
             'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
    foreach($root in $roots){
        if(-not(Test-Path -Path $root)){continue}
        try {
            foreach($key in @(Get-ChildItem -Path $root -ErrorAction Stop)) {
                try {
                    # PS2: Get-ItemProperty -LiteralPath is not supported; use -Path.
                    $p=Get-ItemProperty -Path $key.PSPath -ErrorAction Stop
                    if($p.DisplayName -and (AVMatchName ([string]$p.DisplayName))) {
                        $result+= (New-Object PSObject -Property @{Name=(AVText $p.DisplayName);Version=(AVText $p.DisplayVersion);Source='Registry uninstall'})
                    }
                } catch {}
            }
        }catch {SecLog ('AV uninstall registry: '+$root+' : '+(AVText $_.Exception.Message)) 'WARN'}
    }
    return @($result | Sort-Object Name,Version -Unique)
}
function AVFindKescli {
    $roots=@($env:ProgramFiles,${env:ProgramFiles(x86)},$env:ProgramW6432) | Where-Object {$_}
    $candidates=@()
    foreach($root in $roots){
        $vendorDir=Join-Path $root 'Kaspersky Lab'
        if(-not (Test-Path -Path $vendorDir)){continue}
        try {
            $candidates += @(Get-ChildItem -Path $vendorDir -Filter 'kescli.exe' -Recurse -ErrorAction SilentlyContinue | Where-Object {-not $_.PSIsContainer} | ForEach-Object {$_.FullName})
        }catch{}
    }
    foreach($exe in @($candidates | Select-Object -Unique)) {
        try {
            $sig=Get-AuthenticodeSignature -FilePath $exe -ErrorAction Stop
            if(([string]$sig.Status) -eq 'Valid' -and $sig.SignerCertificate -and
               ([string]$sig.SignerCertificate.Subject) -match '(?i)Kaspersky') {
                return $exe
            }
            SecLog ('KESCLI signature not verified, skipped: '+$exe) 'WARN'
        }catch {SecLog ('KESCLI signature check failed: '+(AVText $_.Exception.Message)) 'WARN'}
    }
    return ''
}
function AVRunKescli([string]$file,[string]$action){
    $r=New-Object PSObject -Property @{Success=$false;Output='';Error='';ExitCode='';TimedOut=$false}
    $p=$null
    try {
        $psi=New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName=$file
        $psi.Arguments='--opswat '+$action
        $psi.WorkingDirectory=Split-Path -Parent $file
        $psi.UseShellExecute=$false
        $psi.CreateNoWindow=$true
        $psi.RedirectStandardOutput=$true
        $psi.RedirectStandardError=$true
        $p=[System.Diagnostics.Process]::Start($psi)
        if(-not $p){throw 'Cannot start vendor tool'}
        # Limit only optional vendor probes; no timeout changes to Windows/Office licensing.
        if(-not $p.WaitForExit(12000)) {
            $r.TimedOut=$true; $r.Error='QUA_THOI_GIAN_KESCLI_12S'
            try{$p.Kill()}catch{};return $r
        }
        $r.ExitCode=$p.ExitCode
        $r.Output=AVText ($p.StandardOutput.ReadToEnd())
        $r.Error=AVText ($p.StandardError.ReadToEnd())
        $r.Success=($p.ExitCode -eq 0)
        Write-RawCommandText ('Kaspersky_KESCLI_'+$action) ($file+' --opswat '+$action) $r.Output $r.Error ('EXIT_'+$p.ExitCode)
    }catch{$r.Error=AVText $_.Exception.Message}
    finally{if($p){try{$p.Dispose()}catch{}}}
    return $r
}
function GetAVFullSummary($apps){
    $result=New-Object PSObject -Property @{
        Names='KHONG_XAC_DINH'; RealTime='CAN_XAC_MINH'; Date='CAN_XAC_MINH';Source='';
        Note='';DefenderSig='';DefenderState='CAN_XAC_MINH';Kaspersky='KHONG_PHAT_HIEN';
        Versions='';RealTimeDetail='';UpdateDetail='';Services='';LicenseDetail='';EvidenceDetail='';
        ProductCount='0';ScanDetail='';EngineDetail=''
    }
    $installed=@(AVFindInstalled $apps)
    $wsc=@();$wscError=''
    foreach($ns in @('root\SecurityCenter2','root\SecurityCenter')) {
        try {$wsc=@(Get-WmiObject -Namespace $ns -Class AntiVirusProduct -ErrorAction Stop);Write-RawEvidence 'AV_SecurityCenter_AntiVirusProduct' ('Get-WmiObject -Namespace '+$ns+' -Class AntiVirusProduct') $wsc;break}
        catch{$wscError=AVText $_.Exception.Message}
    }
    $services=@()
    try {$services=@(Get-Service -ErrorAction Stop);Write-RawEvidence 'Services_All' 'Get-Service' $services}catch{SecLog ('AV Get-Service: '+(AVText $_.Exception.Message)) 'WARN'}
    $def=$null;$defErr=''
    if(Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue){
        try {$def=Get-MpComputerStatus -ErrorAction Stop;Write-RawEvidence 'Defender_GetMpComputerStatus' 'Get-MpComputerStatus' $def} catch {$defErr=AVText $_.Exception.Message;SecLog ('Defender query: '+$defErr) 'WARN'}
    }else{$defErr='CMDLET_KHONG_CO_TREN_WINDOWS_NAY'}
    $registeredNames=@($wsc | ForEach-Object {AVText $_.displayName} | Where-Object {$_})
    $items=@()
    foreach($name in @($registeredNames)+@($installed | ForEach-Object {$_.Name})){ if($name -and ($items -notcontains $name)) {$items+= $name} }
    if($def -and -not (@($items | Where-Object {$_ -match '(?i)Defender|Security Essentials'}).Count)) {$items+='Microsoft Defender Antivirus'}
    if($items.Count -eq 0){
        $result.Names='KHONG_PHAT_HIEN_TRONG_CAC_NGUON_DA_QUET'
        $result.Note='Khong phat hien dang ky hoac uninstall; khong chung minh may khong co antivirus.'
        if($wscError){$result.Note+=' | WSC: '+$wscError}
        return $result
    }
    $result.ProductCount=[string]$items.Count
    $result.Names=AVUnique $items
    $result.Source='Windows Security Center | Uninstall Registry | Defender (neu ho tro) | KESCLI (neu ho tro)'
    $versionRows=@();$rtRows=@();$updateRows=@();$serviceRows=@();$licenseRows=@();$evidenceRows=@();$scanRows=@();$engineRows=@()
    $kesTool=''
    if(@($items | Where-Object {$_ -match '(?i)Kaspersky Endpoint Security'}).Count -gt 0) {$kesTool=AVFindKescli}
    $kesRtp=$null;$kesUpdate=$null;$kesScan=$null
    if($kesTool){
        SecLog ('Found signed Kaspersky CLI: '+$kesTool) 'INFO'
        $kesRtp=AVRunKescli $kesTool 'GetRealTimeProtectionState'
        $kesUpdate=AVRunKescli $kesTool 'GetDefinitionState'
        $kesScan=AVRunKescli $kesTool 'GetLastScanTime'
        SecLog ('KESCLI RTP (exit='+$kesRtp.ExitCode+'): '+$kesRtp.Output+' '+$kesRtp.Error) 'INFO'
        SecLog ('KESCLI definition (exit='+$kesUpdate.ExitCode+'): '+$kesUpdate.Output+' '+$kesUpdate.Error) 'INFO'
    }
    foreach($name in $items){
        $inst=@($installed | Where-Object {$_.Name -eq $name})
        $matchingWsc=@($wsc | Where-Object {$_.displayName -eq $name})
        $versions=AVUnique @($inst | ForEach-Object {$_.Version})
        $kind='AV/Endpoint Security - can xac nhan';$rt='CAN_XAC_MINH';$last='CAN_XAC_MINH';$license='CAN_XAC_MINH_TRONG_UNG_DUNG';$scan='CAN_XAC_MINH';$engine='CAN_XAC_MINH'
        $origin=@()
        if($inst.Count -gt 0){$origin+='Uninstall'}
        if($matchingWsc.Count -gt 0){$origin+='WSC'}
        if($name -match '(?i)CrowdStrike|SentinelOne|Cylance') {$kind='Endpoint/EDR; chua ket luan la antivirus'}
        if($name -match '(?i)Defender|Security Essentials'){
            if($def -and $name -notmatch '(?i)Security Essentials'){
                $origin+='Get-MpComputerStatus'
                $versions=AVUnique @($versions,(AVText $def.AMProductVersion))
                if($def.RealTimeProtectionEnabled -eq $true -and $def.AntivirusEnabled -eq $true){$rt='BAT (Defender cmdlet)'}
                elseif($def.RealTimeProtectionEnabled -eq $false -or $def.AntivirusEnabled -eq $false){$rt='TAT_HOAC_PASSIVE (Defender cmdlet)'}
                if($def.AntivirusSignatureLastUpdated){
                    try{$last=([datetime]$def.AntivirusSignatureLastUpdated).ToString('yyyy-MM-dd HH:mm:ss')}
                    catch{$last=AVText $def.AntivirusSignatureLastUpdated}
                    $result.DefenderSig=$last
                }
                $result.DefenderState=$rt
                $engine='Engine '+(AVText $def.AMEngineVersion)+'; Signature '+(AVText $def.AntivirusSignatureVersion)
                $scan='Quick '+(AVText $def.QuickScanEndTime)+'; Full '+(AVText $def.FullScanEndTime)
                $license='TICH_HOP_WINDOWS / CAN_XAC_MINH_LICENSE'
            }else{$origin+='WSC/Registry (khong co cmdlet chi tiet)'}
        }
        if($name -match '(?i)Kaspersky'){
            $result.Kaspersky='CO_CAI_HOAC_DANG_KY / CAN_XAC_MINH_LICENSE'
            if($inst.Count -gt 0){$result.Kaspersky='CO_CAI '+$versions+' / CAN_XAC_MINH_LICENSE'}
            if($name -match '(?i)Endpoint Security' -and $kesTool){
                $origin+='KESCLI_CHU_KY_HOP_LE'
                if($kesRtp.Success -and $kesRtp.Output -match '^\s*1\s*$'){$rt='BAT (KESCLI File Threat Protection)'}
                elseif($kesRtp.Success -and $kesRtp.Output -match '^\s*0\s*$'){$rt='TAT (KESCLI File Threat Protection)'}
                else{$rt='CAN_XAC_MINH_KESCLI';if($kesRtp.Error){SecLog ('KESCLI RTP error: '+$kesRtp.Error) 'WARN'}}
                if($kesUpdate.Success -and $kesUpdate.Output){
                    # Keep vendor output verbatim: format can vary by KES version.
                    $last='KESCLI_RAW: '+$kesUpdate.Output
                }elseif($kesUpdate.Error){SecLog ('KESCLI definition error: '+$kesUpdate.Error) 'WARN'}
                if($kesScan.Success -and $kesScan.Output){$scan='KESCLI_RAW: '+$kesScan.Output}
                elseif($kesScan.Error){SecLog ('KESCLI last scan error: '+$kesScan.Error) 'WARN'}
            }
        }
        $svcMatches=@()
        foreach($s in $services){
            $svcName=[string]$s.Name;$svcDisplay=[string]$s.DisplayName
            $vendorPattern=''
            if($name -match '(?i)Kaspersky'){$vendorPattern='(?i)Kaspersky|^AVP$|^kavfs$'}
            elseif($name -match '(?i)Defender'){$vendorPattern='(?i)WinDefend|Microsoft Defender Antivirus'}
            elseif($name -match '(?i)Security Essentials'){$vendorPattern='(?i)MsMpSvc|Microsoft Antimalware'}
            elseif($name -match '(?i)Reason|RAV Endpoint'){$vendorPattern='(?i)Reason|RAV Endpoint|rsEngine'}
            elseif($name -match '(?i)Malwarebytes'){$vendorPattern='(?i)Malwarebytes|MBAMService'}
            elseif($name -match '(?i)ESET|NOD32'){$vendorPattern='(?i)ESET|^ekrn$'}
            elseif($name -match '(?i)Sophos'){$vendorPattern='(?i)Sophos'}
            elseif($name -match '(?i)McAfee|Trellix'){$vendorPattern='(?i)McAfee|Trellix'}
            elseif($name -match '(?i)Bitdefender'){$vendorPattern='(?i)Bitdefender|^vsserv$'}
            elseif($name -match '(?i)Trend Micro'){$vendorPattern='(?i)Trend Micro'}
            elseif($name -match '(?i)Avast'){$vendorPattern='(?i)Avast'}
            elseif($name -match '(?i)AVG'){$vendorPattern='(?i)AVG'}
            elseif($name -match '(?i)Norton|Symantec'){$vendorPattern='(?i)Norton|Symantec'}
            if($vendorPattern -and ($svcName -match $vendorPattern -or $svcDisplay -match $vendorPattern)){$svcMatches+=($svcName+'='+$s.Status)}
        }
        if($matchingWsc.Count -gt 0){
            foreach($w in $matchingWsc){
                $evidenceRows+=($name+': WSC productState='+(AVText $w.productState)+'; timestamp='+(AVText $w.timestamp)+' (NOT signature date)')
            }
        }
        $versionRows+=($name+': '+$(if($versions){$versions}else{'CHUA_XAC_MINH'}))
        $rtRows+=($name+': '+$rt)
        $updateRows+=($name+': '+$last)
        $serviceRows+=($name+': '+$(if($svcMatches.Count){AVUnique $svcMatches}else{'CHUA_TIM_THAY_SERVICE_PHU_HOP'}))
        $licenseRows+=($name+': '+$license)
        $scanRows+=($name+': '+$scan)
        $engineRows+=($name+': '+$engine)
        $evidenceRows+=($name+': Loai='+$kind+'; Nguon='+(AVUnique $origin))
        SecLog ('Antivirus: '+$name+' | Version='+$versions+' | RTP='+$rt+' | Definition='+$last+' | Services='+(AVUnique $svcMatches)) 'INFO'
    }
    $result.Versions=($versionRows -join "`n")
    $result.RealTimeDetail=($rtRows -join "`n")
    $result.UpdateDetail=($updateRows -join "`n")
    $result.Services=($serviceRows -join "`n")
    $result.LicenseDetail=($licenseRows -join "`n")
    $result.EvidenceDetail=($evidenceRows -join "`n")
    $result.ScanDetail=($scanRows -join "`n")
    $result.EngineDetail=($engineRows -join "`n")
    $confirmed=@($rtRows | Where-Object {$_ -match ': BAT \('})
    if($confirmed.Count -gt 0){$result.RealTime='BAT_XAC_NHAN: '+(AVUnique $confirmed)}
    else{$result.RealTime='CAN_XAC_MINH_TRONG_UNG_DUNG'}
    $observed=@($updateRows | Where-Object {$_ -notmatch ': CAN_XAC_MINH$'})
    if($observed.Count -eq 1){$result.Date=$observed[0]}
    elseif($observed.Count -gt 1){$result.Date='XEM_COT_CAP_NHAT_TUNG_AV'}
    else{$result.Date='CAN_XAC_MINH_TRONG_UNG_DUNG'}
    $result.Note='WSC productState chi luu tham chieu, khong suy luan RTP/ngay update; RTP KESCLI chi phan anh File Threat Protection; ngay KESCLI_RAW can doi chieu phien ban.'
    if($wscError -and $wsc.Count -eq 0){$result.Note+=' | WSC: '+$wscError}
    if($defErr){$result.Note+=' | Defender: '+$defErr}
    if($items.Count -gt 1){$result.Note+=' | Nhieu phan mem duoc phat hien, khong mac dinh tat ca deu dang hoat dong.'}
    return $result
}
