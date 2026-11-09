# Shared read-only Registry helper; PS2 compatible.
if(-not (Get-Command Read-RegistryEvidence -ErrorAction SilentlyContinue)) {
    . (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'RegistryEvidence.ps1')
}
# Extra PA05 evidence, READ ONLY. Compatible with Windows PowerShell 2.0 (Win7).
# No network scans, GPO edits, service changes, package removal or configuration changes.
function CEReg([string]$path,[string]$name) {
    $e=Read-RegistryEvidence -Path $path -Name $name
    if($e.State -eq 'FOUND'){return $e.Value}
    return $null
}
function CEStr([object]$x) { if ($null -eq $x){return ''}; return ([string]$x).Trim() -replace '[\r\n\t]+',' ' }
function CEError([object]$e) {
    $s=CEStr $e.Exception.Message
    if($s -match '(?i)denied|0x80070005|unauthorized') {return 'THIEU_QUYEN: '+$s}
    return 'LOI_TRUY_VAN: '+$s
}
function CEOSBuild {
    try {
        $os=Get-WmiObject -Class Win32_OperatingSystem -ErrorAction Stop | Select-Object -First 1
        $ver=[version]([string]$os.Version)
        return [int]$ver.Major*100+[int]$ver.Minor
    }catch {return 999}
}
function CEGetSMB1Feature([string]$featureName) {
    # This cmdlet is not installed on Windows 7/PowerShell 2.0; read-only when available.
    $r=New-Object PSObject -Property @{State='';Error='';Checked=$false}
    if(-not (Get-Command Get-WindowsOptionalFeature -ErrorAction SilentlyContinue)) {return $r}
    try {
        $x=Get-WindowsOptionalFeature -Online -FeatureName $featureName -ErrorAction Stop
        if($null -ne $x) {$r.State=CEStr $x.State; $r.Checked=$true}
    }catch {$r.Error=CEError $_}
    return $r
}
function CEGetSMB1 {
    $r=New-Object PSObject -Property @{Status='CAN_XAC_MINH';Server='';Client='';Note='';Source='';Error=''}
    try {
        $serverKey='HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters'
        $clientKey='HKLM:\SYSTEM\CurrentControlSet\Services\mrxsmb10'
        $regServer=CEReg $serverKey 'SMB1'
        $regClient=CEReg $clientKey 'Start'
        $notes=@()
        $sources=@()
        $featureServer=$null
        $featureClient=$null
        # Prefer the actual SMB Server setting on Windows versions that support it.
        $serverConfigured=$null
        if(Get-Command Get-SmbServerConfiguration -ErrorAction SilentlyContinue) {
            try {
                $cfg=Get-SmbServerConfiguration -ErrorAction Stop
                if($null -ne $cfg -and $null -ne $cfg.EnableSMB1Protocol) {
                    $serverConfigured=[bool]$cfg.EnableSMB1Protocol
                    $sources+= 'Get-SmbServerConfiguration'
                }
            }catch {$notes+=('Get-SmbServerConfiguration: '+(CEError $_))}
        }
        if($null -ne $serverConfigured -and -not $serverConfigured) {
            $r.Server='DISABLED'
        } else {
            # Need Windows Feature status for ambiguous cases: enabled server config
            # does not always mean SMB1 component is installed.
            if((Get-Command Get-WindowsOptionalFeature -ErrorAction SilentlyContinue)) {
                $featureServer=CEGetSMB1Feature 'SMB1Protocol-Server'
                if($featureServer.Checked){$sources+='WindowsFeature SMB1Protocol-Server='+$featureServer.State}
            }
            if($null -ne $featureServer -and $featureServer.Checked -and $featureServer.State -match '^Disabled') {
                $r.Server='DISABLED_FEATURE'
            }elseif($null -ne $featureServer -and $featureServer.Checked -and $featureServer.State -match 'Pending') {
                $r.Server='CAN_XAC_MINH_CHANGE_PENDING'
            }elseif($null -ne $serverConfigured -and $serverConfigured) {
                if($null -ne $regServer -and [int]$regServer -eq 0){
                    $r.Server='CAN_XAC_MINH_MAU_THUAN_SMB_SERVER'
                    $notes+='Get-SmbServerConfiguration=True nhung Registry SMB1=0; co the chua restart, can doi chieu.'
                }elseif($null -eq $featureServer -or -not $featureServer.Checked){$r.Server='ENABLED_CONFIG_CAN_DOI_CHIEU_FEATURE'}
                else{$r.Server='ENABLED_CONFIG'}
            }elseif($null -ne $regServer -and [int]$regServer -eq 0) {
                $r.Server='DISABLED_CONFIG'
            }elseif($null -ne $regServer -and [int]$regServer -eq 1) {
                $r.Server='ENABLED_CONFIG_CAN_DOI_CHIEU_FEATURE'
            }elseif($null -eq $regServer) {
                if((CEOSBuild) -ge 601 -and (CEOSBuild) -le 603) {$r.Server='KHONG_CAU_HINH - WIN7_8_81_MAC_DINH_CO_THE_BAT'}
                else {$r.Server='KHONG_CAU_HINH - CAN_DOI_CHIEU_FEATURE'}
            }else {$r.Server='CAN_XAC_MINH_GIA_TRI_REGISTRY'}
        }
        if($null -ne $regServer){$sources+=('Registry SMB1='+[string]$regServer)}

        # Client is separate from Server. Check Feature when available; older
        # Windows must fall back to the MrxSmb10 driver state.
        if($null -ne $regClient -and [int]$regClient -eq 4){
            $r.Client='DISABLED_CONFIG'
            # Driver disabled is meaningful, but if Feature is enabled note mismatch.
            $featureClient=CEGetSMB1Feature 'SMB1Protocol-Client'
            if($featureClient.Checked -and $featureClient.State -eq 'Enabled') {
                $r.Client='DISABLED_CONFIG - FEATURE_ENABLED_CAN_DOI_CHIEU'
                $notes+='SMBv1 Client: driver mrxsmb10 Start=4 nhung feature Enabled; kiem tra sau khi restart.'
            }
            if($featureClient.Checked){$sources+='WindowsFeature SMB1Protocol-Client='+$featureClient.State}
        }else{
            if((Get-Command Get-WindowsOptionalFeature -ErrorAction SilentlyContinue)) {
                $featureClient=CEGetSMB1Feature 'SMB1Protocol-Client'
                if($featureClient.Checked){$sources+='WindowsFeature SMB1Protocol-Client='+$featureClient.State}
            }
            if($null -ne $featureClient -and $featureClient.Checked -and $featureClient.State -match '^Disabled') {
                $r.Client='DISABLED_FEATURE'
            }elseif($null -ne $featureClient -and $featureClient.Checked -and $featureClient.State -match 'Pending') {
                $r.Client='CAN_XAC_MINH_CHANGE_PENDING'
            }elseif($null -ne $regClient) {
                if($null -ne $featureClient -and $featureClient.Checked -and $featureClient.State -eq 'Enabled'){
                    $r.Client='ENABLED_CONFIG'
                }else{
                    $r.Client='START_'+[string]$regClient+'_CAN_DOI_CHIEU_FEATURE'
                }
            }elseif(Test-Path $clientKey) {
                $r.Client='CAN_XAC_MINH_CLIENT_START_KHONG_CO'
            }elseif($null -ne $featureClient -and $featureClient.Checked -and $featureClient.State -eq 'Enabled') {
                $r.Client='CAN_XAC_MINH_CLIENT_FEATURE_DRIVER_MAU_THUAN'
            }else{
                $r.Client='CAN_XAC_MINH_CLIENT_KHONG_TIM_THAY_DRIVER'
            }
        }
        if($null -ne $regClient){$sources+=('mrxsmb10 Start='+[string]$regClient)}
        if($null -ne $featureServer -and $featureServer.Error){$notes+=('ServerFeature: '+$featureServer.Error)}
        if($null -ne $featureClient -and $featureClient.Error){$notes+=('ClientFeature: '+$featureClient.Error)}
        $notes+='ENABLED_CONFIG chi la cau hinh, KHONG xac nhan dang co ket noi SMBv1.'
        $notes+='Windows 7/8/8.1: Registry SMB1 khong ton tai co the van cho phep Server SMBv1.'
        $notes+='Khong thay mrxsmb10 khong du de ket luan Client OFF neu chua doi chieu Windows Feature.'
        $r.Status='SERVER: '+$r.Server+' | CLIENT: '+$r.Client
        $r.Source=$sources -join ' ; '
        $r.Note=($notes -join ' | ')+' | Nguon: '+$r.Source
    }catch{
        $r.Error=CEError $_
        $r.Status='CAN_XAC_MINH_SMB1'
    }
    return $r
}
function CEGetAutoRun {
    # No registry value by itself proves the effective result of every AutoPlay/AutoRun policy.
    # Policy scopes and interactive-user preferences are reported separately.
    # Windows 7/8 and Windows 10/11; Get-ItemProperty exists in PowerShell 2.0.
    $r=New-Object PSObject -Property @{
        Status='CAN_XAC_MINH';HKLM='KHONG_CO';HKCU='KHONG_CO';NoDriveAutoRun='';
        UserPreference='KHONG_CO_GIA_TRI';PolicyEvidence='';CheckedUser='';Note='';Error=''
    }
    try {
        $r.CheckedUser=[System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $base='\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer'
        $policyBase='\SOFTWARE\Policies\Microsoft\Windows\Explorer'
        $found=@()
        $roots=@('HKLM','HKCU')
        $driveValues=@{}
        foreach($rootName in $roots){
            $root=$rootName+':'+$base
            $v=CEReg $root 'NoDriveTypeAutoRun'
            if($null -ne $v){
                $num=[int64]$v
                $formatted=$num.ToString()+' (0x'+$num.ToString('X')+')'
                $r.$rootName=$formatted
                $driveValues[$rootName]=$num
                $found+=($rootName+' NoDriveTypeAutoRun='+$formatted)
            }
            $perDrive=CEReg $root 'NoDriveAutoRun'
            if($null -ne $perDrive){$r.NoDriveAutoRun+=($rootName+'='+[string]$perDrive+'; ');$found+=($rootName+' NoDriveAutoRun='+[string]$perDrive)}
            $noAutorun=CEReg $root 'NoAutorun'
            if($null -ne $noAutorun){$found+=($rootName+' NoAutorun='+[string]$noAutorun)}
            $rootPolicy=$rootName+':'+$policyBase
            # Additional Administrative Templates policy source, not just legacy Explorer settings.
            $newMask=CEReg $rootPolicy 'NoDriveTypeAutoRun'
            if($null -ne $newMask){
                $maskPolicy=[int64]$newMask
                $found+=($rootName+' Policy NoDriveTypeAutoRun='+[string]$maskPolicy)
                # Explicit Group Policy is more important than a legacy preference at the same scope.
                $driveValues[$rootName]=$maskPolicy
            }
            $newNoDrives=CEReg $rootPolicy 'NoDriveAutoRun'
            if($null -ne $newNoDrives){$found+=($rootName+' Policy NoDriveAutoRun='+[string]$newNoDrives)}
            $nonVolume=CEReg $rootPolicy 'NoAutoplayfornonVolume'
            if($null -ne $nonVolume){$found+=($rootName+' NoAutoplayfornonVolume='+[string]$nonVolume)}
        }
        # UI preference is not GPO. HKCU describes the account executing the script.
        $userValue=CEReg 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\AutoplayHandlers' 'DisableAutoplay'
        if($null -ne $userValue){
            if([int64]$userValue -eq 1){$r.UserPreference='DISABLED_BY_USER (DisableAutoplay=1)'}
            elseif([int64]$userValue -eq 0){$r.UserPreference='NOT_DISABLED_BY_USER (DisableAutoplay=0)'}
            else{$r.UserPreference='UNKNOWN_VALUE='+[string]$userValue}
        }
        $r.PolicyEvidence=($found -join ' | ')
        # Computer Group Policy generally overrides conflicting user settings.
        $effectiveScope=''
        $mask=$null
        if($driveValues.ContainsKey('HKLM')){$effectiveScope='HKLM';$mask=$driveValues['HKLM']}
        elseif($driveValues.ContainsKey('HKCU')){$effectiveScope='HKCU';$mask=$driveValues['HKCU']}
        if($null -ne $mask){
            if($mask -eq 255){$r.Status='AUTOPLAY_POLICY_ALL_DRIVE_TYPES_DISABLED ('+$effectiveScope+' 0xFF)'}
            else{$r.Status='AUTOPLAY_POLICY_MASK_REVIEW ('+$effectiveScope+' 0x'+$mask.ToString('X')+')'}
        }elseif($r.PolicyEvidence){
            $r.Status='POLICY_PARTIAL_NEEDS_REVIEW'
        }elseif($r.UserPreference -match '^DISABLED_BY_USER'){
            $r.Status='AUTOPLAY_USER_DISABLED_NO_GPO_FOUND'
        }elseif($r.UserPreference -match '^NOT_DISABLED_BY_USER'){
            $r.Status='AUTOPLAY_USER_NOT_DISABLED_NO_GPO_FOUND'
        }else{
            $r.Status='KHONG_TIM_THAY_CAU_HINH_DA_KIEM_TRA'
        }
        $r.Note='Policy NoDriveTypeAutoRun/NoDriveAutoRun/NoAutorun va UI DisableAutoplay la cac nguon KHAC NHAU. '+
        'Gia tri HKCU thuoc '+$r.CheckedUser+', khong nhat thiet la can bo su dung may. '+
        'Khong tim thay khoa KHONG dong nghia AutoPlay dang bat/tat; mask chua phai ket luan bao mat day du.'
        if($r.UserPreference){$r.Note+=' | UI: '+$r.UserPreference}
        if($r.PolicyEvidence){$r.Note+=' | Policy: '+$r.PolicyEvidence}
    }catch{$r.Status='CAN_XAC_MINH_AUTORUN';$r.Error=CEError $_}
    return $r
}
function CEGetTimeSync {
    $r=New-Object PSObject -Property @{Status='CAN_XAC_MINH';Service='';Type='';NtpServer='';Source='';LastSync='';Error='';Note=''}
    try {
        $s=Get-Service -Name W32Time -ErrorAction Stop
        $r.Service=[string]$s.Status
    }catch{$r.Service='KHONG_DOC_DUOC';$r.Error=CEError $_}
    $path='HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters'
    $r.Type=CEStr (CEReg $path 'Type')
    $r.NtpServer=CEStr (CEReg $path 'NtpServer')
    # w32tm reads local service status, never sends an active scan or changes the clock.
    try {
        $exe=Join-Path $env:SystemRoot 'System32\w32tm.exe'
        if(Test-Path $exe) {
            $psi=New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName=$exe;$psi.Arguments='/query /status'
            $psi.UseShellExecute=$false;$psi.CreateNoWindow=$true
            $psi.RedirectStandardOutput=$true;$psi.RedirectStandardError=$true
            $proc=[System.Diagnostics.Process]::Start($psi)
            if($null -ne $proc){
                try {
                    if($proc.WaitForExit(5000)) {
                        $output=$proc.StandardOutput.ReadToEnd()
                        if($proc.ExitCode -eq 0){
                            foreach($ln in ($output -split '\r?\n')){
                                if($ln -match '^\s*Source\s*:\s*(.*)$'){$r.Source=CEStr $matches[1]}
                                if($ln -match '^\s*Last Successful Sync Time\s*:\s*(.*)$'){$r.LastSync=CEStr $matches[1]}
                            }
                            # Localized output may lack English labels; preserve raw instead of claiming sync.
                            if(-not $r.Source){$r.Note='w32tm thanh cong nhung output khac ngon ngu; doc trong CMD: w32tm /query /status'}
                        } else { $r.Note='w32tm tra ma '+$proc.ExitCode+'; xem w32tm /query /status tren may.' }
                    } else {
                        try{$proc.Kill()}catch{}
                        $r.Note='w32tm qua thoi gian 5 giay, khong ket luan mat dong bo.'
                    }
                }finally{$proc.Dispose()}
            }
        }
    }catch{$r.Note+=' | '+(CEError $_)}
    if(-not $r.Source){
        # w32tm /query /source returns a single source value, independent of
        # the localized "Source:" label in /query /status output.
        try {
            $exe=Join-Path $env:SystemRoot 'System32\w32tm.exe'
            if(Test-Path $exe){
                $psi2=New-Object System.Diagnostics.ProcessStartInfo
                $psi2.FileName=$exe
                $psi2.Arguments='/query /source'
                $psi2.UseShellExecute=$false
                $psi2.CreateNoWindow=$true
                $psi2.RedirectStandardOutput=$true
                $psi2.RedirectStandardError=$false
                $proc2=[System.Diagnostics.Process]::Start($psi2)
                if($null -ne $proc2){
                    try {
                        if($proc2.WaitForExit(4000) -and $proc2.ExitCode -eq 0){
                            $sourceText=CEStr $proc2.StandardOutput.ReadToEnd()
                            if($sourceText -and $sourceText -notmatch '\r|\n'){$r.Source=$sourceText}
                        }else{if(-not $proc2.HasExited){try{$proc2.Kill()}catch{}}}
                    }finally{$proc2.Dispose()}
                }
            }
        }catch{$r.Note+=' | W32Time source fallback: '+(CEError $_)}
    }
    $r.Status='W32Time='+$r.Service+'; Type='+$r.Type+'; NtpServer='+$r.NtpServer
    if($r.Source){$r.Status+='; Source='+$r.Source}
    if(-not $r.LastSync){$r.Note+=' | Chua xac nhan lan dong bo thanh cong gan nhat.'}
    return $r
}
function CEGetDomain([object]$cs) {
    $r=New-Object PSObject -Property @{Status='CAN_XAC_MINH';Name='';Error=''}
    if($null -eq $cs){$r.Error='Win32_ComputerSystem khong doc duoc';return $r}
    try {
        if($cs.PartOfDomain -eq $true){$r.Status='DOMAIN';$r.Name=CEStr $cs.Domain}
        elseif($cs.PartOfDomain -eq $false){$r.Status='WORKGROUP';$r.Name=CEStr $cs.Workgroup}
        else{$r.Status='CAN_XAC_MINH';$r.Name=CEStr $cs.Domain}
    }catch{$r.Error=CEError $_}
    return $r
}
function CEGetIPMode([object[]]$configs) {
    $parts=@()
    foreach($cfg in @($configs)){
        if($null -eq $cfg -or -not $cfg.IPEnabled){continue}
        if(-not @($cfg.IPAddress | Where-Object {$_ -match '^\d{1,3}(\.\d{1,3}){3}$'}).Count){continue}
        $mode=if($cfg.DHCPEnabled -eq $true){'DHCP'}elseif($cfg.DHCPEnabled -eq $false){'STATIC_CONFIG'}else{'CHUA_XAC_MINH'}
        $label=CEStr $cfg.Description
        if($label.Length -gt 50){$label=$label.Substring(0,50)}
        $parts+=($label+'='+$mode)
    }
    if($parts.Count -eq 0){return 'KHONG_XAC_DINH_IP_CONFIG'}
    return ($parts -join ' | ')
}
function GetComplianceExtras([object]$cs,[object[]]$configs) {
    $domain=CEGetDomain $cs
    $smb=CEGetSMB1
    $auto=CEGetAutoRun
    $time=CEGetTimeSync
    $mode=CEGetIPMode $configs
    return (New-Object PSObject -Property @{
        DomainStatus=$domain.Status;DomainName=$domain.Name;DomainError=$domain.Error;
        SMB1=$smb.Status;SMB1Server=$smb.Server;SMB1Client=$smb.Client;SMB1Note=$smb.Note;SMB1Error=$smb.Error;
        AutoRun=$auto.Status;AutoRunHKLM=$auto.HKLM;AutoRunHKCU=$auto.HKCU;AutoRunUserPreference=$auto.UserPreference;AutoRunPolicyEvidence=$auto.PolicyEvidence;AutoRunCheckedUser=$auto.CheckedUser;AutoRunNote=$auto.Note;AutoRunError=$auto.Error;
        NTP=$time.Status;NTPService=$time.Service;NTPType=$time.Type;NTPServer=$time.NtpServer;NTPSource=$time.Source;NTPLastSync=$time.LastSync;NTPNote=$time.Note;NTPError=$time.Error;
        IPMode=$mode;
        NetScan='CAN_DOI_CHIEU_SWITCH_CMDB'; NAC='CAN_DOI_CHIEU_SWITCH_NAC';
        Notes='Khong quet mang, khong thay doi cai dat, khong go bo ung dung. Toan bo du lieu la dau hieu tai thoi diem quet.'
    })
}
