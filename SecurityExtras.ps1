# Shared read-only Registry helper; PS2 compatible.
if(-not (Get-Command Read-RegistryEvidence -ErrorAction SilentlyContinue)) {
    . (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'RegistryEvidence.ps1')
}
# SecurityExtras.ps1 - read-only configuration checks for Windows 7/8/8.1/10/11.
# Baseline: Windows PowerShell 2.0 + WMI; optional cmdlets only if present.
# The data describes observed settings, not compliance with a particular organization's policy.

function ExtraLog([string]$message,[string]$level='INFO') {
    if(Get-Command SurveyLog -ErrorAction SilentlyContinue) { SurveyLog $message $level }
    else { Write-Host ('['+$level+'] '+$message) }
}
function ExtraReadReg([string]$key,[string]$name) {
    $e=Read-RegistryEvidence -Path $key -Name $name
    if($e.State -eq 'FOUND'){return $e.Value}
    return $null
}
function ExtraError([object]$errorObj) {
    $s=[string]$errorObj.Exception.Message
    if($s -match '(?i)access is denied|access denied|unauthorized|0x80070005|denied') {return 'THIEU_QUYEN: '+$s}
    if($s -match '(?i)invalid class|invalid namespace|0x8004100e|0x80041010') {return 'WMI_KHONG_HO_TRO: '+$s}
    return 'LOI_TRUY_VAN: '+$s
}
function ExtraJoin($values) {
    $a=@($values | Where-Object {$_ -ne $null -and ([string]$_).Trim() -ne ''} | Select-Object -Unique)
    if($a.Count -eq 0) {return ''}
    return ($a -join ' | ')
}
function ExtraHotfixDate([object]$value) {
    if($null -eq $value) {return $null}
    if($value -is [datetime]) {return $value}
    $s=([string]$value).Trim()
    if(-not $s) {return $null}
    # Some Win32_QuickFixEngineering versions store a 64-bit hex Windows FILETIME.
    if($s -match '^(?i:[0-9a-f]{16})$') {
        try {
            $ticks=[convert]::ToInt64($s,16)
            if($ticks -gt 1000000000000000) {return [datetime]::FromFileTimeUtc($ticks).ToLocalTime()}
        } catch {}
    }
    $d=[datetime]::MinValue
    $cultures=@([System.Globalization.CultureInfo]::CurrentCulture,[System.Globalization.CultureInfo]::InvariantCulture,
                (New-Object System.Globalization.CultureInfo('en-US')),
                (New-Object System.Globalization.CultureInfo('vi-VN')))
    foreach($culture in $cultures) {
        if([datetime]::TryParse($s,$culture,[System.Globalization.DateTimeStyles]::None,[ref]$d)) {return $d}
    }
    return $null
}
function CheckFirewallCompat {
    $result=New-Object PSObject -Property @{
        Status='CAN_XAC_MINH'; Domain='CAN_XAC_MINH'; Private='CAN_XAC_MINH'; Public='CAN_XAC_MINH';
        CurrentProfiles='CAN_XAC_MINH'; Source='HNetCfg.FwPolicy2 (Vista+)'; Error=''
    }
    try {
        $fw=New-Object -ComObject HNetCfg.FwPolicy2 -ErrorAction Stop
        $names=@{1='Domain';2='Private';4='Public'}
        foreach($p in @(1,2,4)) {
            # COM indexed property FirewallEnabled(profile). PS2-compatible invocation.
            $enabled=[bool]$fw.FirewallEnabled($p)
            $s=if($enabled){'ON'}else{'OFF'}
            $n=$names[$p]
            $result.$n=$s
        }
        $mask=[int]$fw.CurrentProfileTypes
        $act=@()
        foreach($p in @(1,2,4)) {if(($mask -band $p) -ne 0){$act+=$names[$p]}}
        $result.CurrentProfiles=ExtraJoin $act
        if(-not $result.CurrentProfiles){$result.CurrentProfiles='KHONG_XAC_DINH'}
        $parts=@()
        foreach($n in @('Domain','Private','Public')) {$parts+=($n+'='+[string]$result.$n)}
        $result.Status=$parts -join '; '
    } catch {
        $result.Error=ExtraError $_
        # NetSecurity cmdlet is not installed on Windows 7. Call it only when available.
        if(Get-Command Get-NetFirewallProfile -ErrorAction SilentlyContinue) {
            try {
                $profiles=@(Get-NetFirewallProfile -ErrorAction Stop)
                foreach($x in $profiles) {
                    $n=[string]$x.Name
                    if(@('Domain','Private','Public') -contains $n) {
                        $value=if($x.Enabled){'ON'}else{'OFF'}
                        $result.$n=$value
                    }
                }
                $result.Status='Domain='+$result.Domain+'; Private='+$result.Private+'; Public='+$result.Public
                $result.Source='Get-NetFirewallProfile (NetSecurity)'
                $result.Error=''
            } catch {$result.Error=ExtraError $_}
        }
    }
    return $result
}
function CheckHotfixCompat {
    $r=New-Object PSObject -Property @{
        Status='CAN_XAC_MINH'; KB=''; Date=''; Latest3=''; Count=''; Source='Win32_QuickFixEngineering WMI'; Error=''
    }
    try {
        $fixes=@(Get-WmiObject -Class Win32_QuickFixEngineering -ErrorAction Stop)
        $r.Count=[string]$fixes.Count
        if($fixes.Count -eq 0){$r.Status='KHONG_TIM_THAY_HOTFIX_TRONG_WMI'; return $r}
        $dated=@()
        $undated=@()
        foreach($f in $fixes) {
            $dt=ExtraHotfixDate $f.InstalledOn
            $id=[string]$f.HotFixID
            if(-not $id){continue}
            if($dt){
                $dated+=(New-Object PSObject -Property @{ KB=$id; Dt=$dt })
            } else {$undated+=$id}
        }
        if($dated.Count -eq 0) {
            $r.Status='CO_'+$fixes.Count+'_HOTFIX_NHUNG_KHONG_CO_NGAY_HOP_LE'
            $r.Latest3=ExtraJoin (@($undated | Select-Object -First 3))
            return $r
        }
        $sorted=@($dated | Sort-Object -Property Dt -Descending)
        $r.KB=$sorted[0].KB
        $r.Date=$sorted[0].Dt.ToString('yyyy-MM-dd')
        $latest=@()
        foreach($f in @($sorted | Select-Object -First 3)) {$latest+=($f.KB+' ('+$f.Dt.ToString('yyyy-MM-dd')+')')}
        $r.Latest3=ExtraJoin $latest
        $r.Status=$r.KB+' / '+$r.Date
    } catch {$r.Error=ExtraError $_; $r.Status='CAN_XAC_MINH_HOTFIX'}
    return $r
}
function CheckUSBCompat {
    $r=New-Object PSObject -Property @{
       Status='CAN_XAC_MINH'; USBSTOR='KHONG_CO_GIA_TRI'; DenyAll='KHONG_CAU_HINH';
       DenyRead='KHONG_CAU_HINH'; DenyWrite='KHONG_CAU_HINH'; WriteProtect='KHONG_CAU_HINH';
       Source='Registry USBSTOR + RemovableStorageDevices policies'; Error=''
    }
    try {
        $start=ExtraReadReg 'HKLM:\SYSTEM\CurrentControlSet\Services\USBSTOR' 'Start'
        if($null -ne $start) {$r.USBSTOR=[string]$start}
        $roots=@('HKLM:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices',
                 'HKCU:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices')
        $deny=@();$reads=@();$writes=@()
        foreach($root in $roots){
            $scope=if($root -like 'HKLM:*'){'MAY'}else{'TAI_KHOAN_CHAY_SCRIPT'}
            $v=ExtraReadReg $root 'Deny_All'
            if($null -ne $v){$deny+=($scope+'='+[string]$v)}
            $disk=$root+'\{53f5630d-b6bf-11d0-94f2-00a0c91efb8b}'
            $v=ExtraReadReg $disk 'Deny_Read'
            if($null -ne $v){$reads+=($scope+'='+[string]$v)}
            $v=ExtraReadReg $disk 'Deny_Write'
            if($null -ne $v){$writes+=($scope+'='+[string]$v)}
        }
        if($deny.Count -gt 0){$r.DenyAll=ExtraJoin $deny}
        if($reads.Count -gt 0){$r.DenyRead=ExtraJoin $reads}
        if($writes.Count -gt 0){$r.DenyWrite=ExtraJoin $writes}
        $wp=ExtraReadReg 'HKLM:\SYSTEM\CurrentControlSet\Control\StorageDevicePolicies' 'WriteProtect'
        if($null -ne $wp){$r.WriteProtect=[string]$wp}
        if([string]$start -eq '4') {$r.Status='USBSTOR_DISABLED (Start=4)'}
        elseif([string]$start -eq '3') {$r.Status='USBSTOR_MANUAL (Start=3; chua du ket luan cho phep USB)'}
        elseif($null -ne $start) {$r.Status='USBSTOR_START='+$start}
        else {$r.Status='KHONG_CO_USBSTOR_VALUE'}
        $blocks=@()
        if(@($deny | Where-Object {$_ -match '=1$'}).Count -gt 0){$blocks+='DENY_ALL_POLICY'}
        if(@($reads | Where-Object {$_ -match '=1$'}).Count -gt 0){$blocks+='DENY_READ_REMOVABLE_DISKS'}
        if(@($writes | Where-Object {$_ -match '=1$'}).Count -gt 0){$blocks+='DENY_WRITE_REMOVABLE_DISKS'}
        if([string]$wp -eq '1'){$blocks+='WRITE_PROTECT'}
        if($blocks.Count -gt 0){$r.Status+='; '+(ExtraJoin $blocks)}
    } catch {$r.Status='CAN_XAC_MINH_USB';$r.Error=ExtraError $_}
    return $r
}
function CheckGuestCompat {
    $r=New-Object PSObject -Property @{Status='CAN_XAC_MINH';Name='';Enabled='';Source='Win32_UserAccount SID *-501';Error=''}
    try {
        # SID RID 501 identifies the built-in Guest account even on localized Windows.
        $accounts=@(Get-WmiObject -Class Win32_UserAccount -Filter 'LocalAccount=True' -ErrorAction Stop)
        $guest=@($accounts | Where-Object {$_.SID -match '-501$'} | Select-Object -First 1)
        if($guest.Count -eq 0){$r.Status='GUEST_KHONG_TIM_THAY_TRONG_WMI';return $r}
        $r.Name=[string]$guest[0].Name
        $r.Enabled=if($guest[0].Disabled){'NO'}else{'YES'}
        $r.Status=if($guest[0].Disabled){'GUEST_DISABLED'}else{'GUEST_ENABLED'}
    } catch {$r.Status='CAN_XAC_MINH_GUEST';$r.Error=ExtraError $_}
    return $r
}
function CheckDNSCompat {
    $r=New-Object PSObject -Property @{Status='CAN_XAC_MINH';PerAdapter='';Source='Win32_NetworkAdapterConfiguration (WMI)';Error=''}
    try {
        $configs=@(Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True' -ErrorAction Stop)
        if($configs.Count -eq 0) {$r.Status='KHONG_CO_CARD_MANG_DANG_DUNG_IP';return $r}
        $all=@();$details=@();$preferred=@()
        foreach($c in $configs) {
            $servers=@($c.DNSServerSearchOrder | Where-Object {$_})
            $label=([string]$c.Description).Trim()
            if($label.Length -gt 68){$label=$label.Substring(0,68)}
            if($servers.Count -gt 0) {
                $all+=@($servers)
                $line=$label+': '+($servers -join ', ')
                $details+=$line
                if(@($c.DefaultIPGateway | Where-Object {$_}).Count -gt 0){$preferred+=@($servers)}
            } else {$details+=($label+': KHONG_CO_DNS_SERVER')}
        }
        $r.PerAdapter=ExtraJoin $details
        if($preferred.Count -gt 0){$r.Status=ExtraJoin $preferred}
        elseif($all.Count -gt 0){$r.Status=(ExtraJoin $all)+' (tat ca card mang; chua xac dinh card chinh)'}
        else {$r.Status='KHONG_CO_DNS_SERVER_TRONG_WMI'}
    } catch {$r.Status='CAN_XAC_MINH_DNS';$r.Error=ExtraError $_}
    return $r
}
function CheckScreenLockCompat {
    $r=New-Object PSObject -Property @{
        Status='CAN_XAC_MINH';Active='KHONG_CAU_HINH';Secure='KHONG_CAU_HINH';TimeOut='KHONG_CAU_HINH';
        MachineInactivity='KHONG_CAU_HINH';User=''; Source='HKCU policy > HKCU Desktop; HKLM Machine Inactivity';
        Error=''; Evidence=''; ScreenSaverExe='KHONG_CAU_HINH'
    }
    try {
        $desktop='HKCU:\Control Panel\Desktop'
        $policy='HKCU:\Software\Policies\Microsoft\Windows\Control Panel\Desktop'
        $r.User=[System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $values=@{}
        $sources=@()
        foreach($name in @('ScreenSaveActive','ScreenSaverIsSecure','ScreenSaveTimeOut','SCRNSAVE.EXE')) {
            $p=Read-RegistryEvidence -Path $policy -Name $name
            if($p.State -eq 'FOUND') {
                $values[$name]=$p.Value; $sources+=('GPO '+$name+'='+[string]$p.Value)
            } else {
                $u=Read-RegistryEvidence -Path $desktop -Name $name
                if($u.State -eq 'FOUND'){
                    $values[$name]=$u.Value; $sources+=('HKCU '+$name+'='+[string]$u.Value)
                }else{
                    $values[$name]=$null
                    $sources+=($name+'='+$u.State+' (policy '+$p.State+')')
                }
            }
        }
        if($null -ne $values['ScreenSaveActive']){$r.Active=[string]$values['ScreenSaveActive']}
        if($null -ne $values['ScreenSaverIsSecure']){$r.Secure=[string]$values['ScreenSaverIsSecure']}
        if($null -ne $values['ScreenSaveTimeOut']){$r.TimeOut=[string]$values['ScreenSaveTimeOut']}
        if($null -ne $values['SCRNSAVE.EXE']){$r.ScreenSaverExe=[string]$values['SCRNSAVE.EXE']}
        $mach=Read-RegistryEvidence -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name 'InactivityTimeoutSecs'
        if($mach.State -eq 'FOUND') {$r.MachineInactivity=[string]$mach.Value}
        $sources+=('Machine inactivity='+$mach.State+':'+$r.MachineInactivity)
        $status=@()
        $seconds=0
        # Machine inactivity limit became supported in Windows 8; do not claim it applies to Win7.
        $os=Get-WmiObject -Class Win32_OperatingSystem -ErrorAction Stop
        $version=[version]([string]$os.Version)
        if($mach.State -eq 'FOUND' -and [int]::TryParse($r.MachineInactivity,[ref]$seconds) -and $seconds -gt 0) {
            if($version.Major -gt 6 -or ($version.Major -eq 6 -and $version.Minor -ge 2)){
                $status+=('MACHINE_INACTIVITY_LOCK='+$seconds+'s (policy)')
            }else{$status+='MACHINE_INACTIVITY_VALUE_NOT_APPLICABLE_WIN7'}
        }
        $secs=0
        if($r.Active -eq '1' -and $r.Secure -eq '1' -and
           [int]::TryParse($r.TimeOut,[ref]$secs) -and $secs -gt 0) {
            # A screensaver executable/valid executable path is still required for actual lock.
            if($r.ScreenSaverExe -eq 'KHONG_CAU_HINH') {
                $status+='SCREENSAVER_SECURE_CONFIG='+$secs+'s - CAN_XAC_MINH_EXE'
            }else{
                $status+='SECURE_SCREENSAVER_LOCK='+$secs+'s (cau hinh)'
            }
        }elseif($r.Active -eq '0' -or $r.Secure -eq '0' -or $r.TimeOut -eq '0'){
            $status+='SCREENSAVER_CONFIG_NOT_SECURE_OR_DISABLED'
        }else{$status+='SCREENSAVER_CHUA_DU_DU_LIEU'}
        $r.Status=$status -join '; '
        $r.Evidence=$sources -join ' | '
        $interactive=''
        try {$cs=Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop; $interactive=[string]$cs.UserName}catch{}
        if($interactive -and $r.User -ine $interactive){
            $r.Status+='; HKCU_CO_THE_LA_ADMIN_KHAC_NGUOI_DUNG'
        }
        if($mach.State -eq 'QUERY_ERROR'){$r.Error='Machine inactivity: '+$mach.Error}
    }catch{$r.Status='CAN_XAC_MINH_SCREEN_LOCK';$r.Error=ExtraError $_}
    return $r
}
function GetExtraSecurityAudit {
    $fw=CheckFirewallCompat
    ExtraLog ('Windows Firewall: '+$fw.Status) $(if($fw.Error){'WARN'}else{'OK'})
    $patch=CheckHotfixCompat
    ExtraLog ('HotFix gan nhat: '+$patch.Status) $(if($patch.Error){'WARN'}else{'OK'})
    $usb=CheckUSBCompat
    ExtraLog ('USB Storage: '+$usb.Status) $(if($usb.Error){'WARN'}else{'OK'})
    $guest=CheckGuestCompat
    ExtraLog ('Guest: '+$guest.Status) $(if($guest.Error){'WARN'}else{'OK'})
    $dns=CheckDNSCompat
    ExtraLog ('DNS: '+$dns.Status) $(if($dns.Error){'WARN'}else{'OK'})
    $lock=CheckScreenLockCompat
    ExtraLog ('Khoa man hinh: '+$lock.Status+' | HKCU user='+$lock.User) $(if($lock.Error){'WARN'}else{'OK'})
    ExtraLog ('ScreenSaver nguon: '+$lock.Evidence) 'INFO'
    $errors=ExtraJoin @($fw.Error,$patch.Error,$usb.Error,$guest.Error,$dns.Error,$lock.Error)
    return (New-Object PSObject -Property @{
        Firewall=$fw.Status;FirewallDomain=$fw.Domain;FirewallPrivate=$fw.Private;FirewallPublic=$fw.Public;
        FirewallCurrentProfiles=$fw.CurrentProfiles;FirewallSource=$fw.Source;
        Hotfix=$patch.Status;HotfixKB=$patch.KB;HotfixDate=$patch.Date;Hotfix3=$patch.Latest3;
        HotfixCount=$patch.Count;HotfixSource=$patch.Source;
        USB=$usb.Status;USBSTOR=$usb.USBSTOR;USBDenyAll=$usb.DenyAll;
        USBDenyRead=$usb.DenyRead;USBDenyWrite=$usb.DenyWrite;USBWriteProtect=$usb.WriteProtect;
        Guest=$guest.Status;GuestName=$guest.Name;GuestEnabled=$guest.Enabled;
        DNS=$dns.Status;DNSPerAdapter=$dns.PerAdapter;
        ScreenLock=$lock.Status;ScreenSaverActive=$lock.Active;ScreenSaverSecure=$lock.Secure;
        ScreenSaverTimeout=$lock.TimeOut;MachineInactivity=$lock.MachineInactivity;ScreenLockUser=$lock.User;
        Errors=$errors
    })
}
