# USB device audit: Windows PowerShell 2.0-5.1, Windows 7/8/8.1/10/11.
# Only reads device/license metadata. Results are stored next to this script.
$ErrorActionPreference = 'Stop'
# SECURITY GATE: fail closed before reading/writing any survey data.
# Works on Windows PowerShell 2.0+ (Windows 7 through Windows 11).
function Test-AuditAdministrator {
    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        return [bool]$principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}
if (-not (Test-AuditAdministrator)) {
    Write-Host ''
    Write-Host '============================================================' -ForegroundColor Red
    Write-Host '  ACCESS DENIED - ADMINISTRATOR PRIVILEGE REQUIRED' -ForegroundColor Red
    Write-Host '============================================================' -ForegroundColor Red
    Write-Host '  Hay chay 1_Run_Survey.bat bang Run as administrator.' -ForegroundColor Yellow
    Write-Host ''
    exit 5
}
$script:root = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $script:root 'CsvCompat.ps1')
$script:folder = Join-Path $script:root 'KhaoSat'
$script:rawFolder = Join-Path $script:folder 'DuLieuGoc'
$script:settings = Join-Path $script:root 'LastLocation.txt'
$legacySettings = Join-Path $script:root 'ViTriGanNhat.txt'
if (-not (Test-Path -LiteralPath $script:settings) -and (Test-Path -LiteralPath $legacySettings)) {
    try { Copy-Item -LiteralPath $legacySettings -Destination $script:settings -ErrorAction Stop } catch { }
}
$script:logFolder = Join-Path $script:folder 'Log'
$script:logFile = $null
$script:errs = @()

# Lightweight ASCII console UI. No GUI modules or PS3+ syntax.
function Show-AuditHeader {
    # Font size is a display preference. Failure is harmless (e.g. Windows Terminal).
    if($env:AUDIT_FONT_PX -and $env:AUDIT_FONT_PX -ne '0') {
        try {
            . (Join-Path $script:root 'ConsoleFont.ps1')
            [void](Set-AuditConsoleFont ([int]$env:AUDIT_FONT_PX))
        } catch { }
    }
    try { $Host.UI.RawUI.WindowTitle = 'Device Audit Tool' } catch {}
    Write-Host ''
    Write-Host '================================================================' -ForegroundColor DarkCyan
    Write-Host '  DEVICE AUDIT TOOL' -ForegroundColor Cyan
    Write-Host '----------------------------------------------------------------' -ForegroundColor DarkCyan
    Write-Host ('  MAY:  '+$env:COMPUTERNAME) -ForegroundColor White
    Write-Host ('  QUYEN: ADMINISTRATOR (DA XAC NHAN)') -ForegroundColor Green
    Write-Host ('  THOI GIAN: '+(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')) -ForegroundColor Gray
    Write-Host '================================================================' -ForegroundColor DarkCyan
    Write-Host ''
}
function Show-AuditStep([string]$message) {
    $match=[regex]::Match($message,'^BUOC ([1-7])(?:[A-D])?/7')
    if (-not $match.Success) { return }
    $n=[int]$match.Groups[1].Value
    $filled=$n*4
    $progress=('=' * $filled)+('-' * (28-$filled))
    Write-Host ''
    Write-Host ('  ['+$progress+']  '+$n+'/7') -ForegroundColor Cyan
    Write-Host ('  '+$message) -ForegroundColor White
    Write-Host ('  '+('-' * 58)) -ForegroundColor DarkGray
}
function Show-AuditFinish([string]$reportPath,[string]$logPath,[int]$errorCount) {
    Write-Host ''
    Write-Host '================================================================' -ForegroundColor DarkCyan
    Write-Host '  KET THUC KHAO SAT' -ForegroundColor Green
    Write-Host '----------------------------------------------------------------' -ForegroundColor DarkCyan
    if ($reportPath) {Write-Host ('  EXCEL: '+$reportPath) -ForegroundColor Green}
    else {Write-Host '  EXCEL: Chua tao duoc - kiem tra Inventory.csv' -ForegroundColor Yellow}
    Write-Host ('  LOG: '+$logPath) -ForegroundColor Gray
    if($errorCount -gt 0){Write-Host ('  CANH BAO: '+$errorCount+' loi truy van - xem log') -ForegroundColor Yellow}
    Write-Host '================================================================' -ForegroundColor DarkCyan
}
function SurveyLog([string]$msg,[string]$level='INFO') {
    $time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = '['+$time+']['+$level+'] '+($msg -replace '[\r\n]+',' ')
    $color = 'Gray'
    if ($level -eq 'OK') { $color = 'Green' }
    if ($level -eq 'WARN') { $color = 'Yellow' }
    if ($level -eq 'ERROR') { $color = 'Red' }
    if ($level -eq 'STEP') { $color = 'Cyan' }
    if ($level -eq 'STEP') { Show-AuditStep $msg }
    Write-Host $line -ForegroundColor $color
    if ($script:logFile) {
        try { Add-Content -LiteralPath $script:logFile -Value $line -Encoding UTF8 -ErrorAction Stop }
        catch { Write-Host ('[WARN] Khong ghi duoc file log: '+$_.Exception.Message) -ForegroundColor Yellow }
    }
}
function Clean([object]$v) {
    if ($null -eq $v) { return '' }
    return ([string]$v).Trim() -replace '[\r\n\t]+',' '
}
function Real([object]$v) {
    $s=Clean $v
    if ($s -eq '') { return $false }
    return ($s -notmatch '(?i)^(default( string)?|to be filled( by o\.?e\.?m\.?)?|o\.?e\.?m\.?|unknown|none|not specified|not available|n/?a|system serial number|system product name|system manufacturer|base board product name|baseboard product name|invalid|0+|f{8,})$')
}
function Wmi([string]$cl) {
    try { return @(Get-WmiObject -Class $cl -ErrorAction Stop) }
    catch { $script:errs += ($cl + ': '+$_.Exception.Message); SurveyLog ('WMI '+$cl+': '+$_.Exception.Message) 'WARN'; return @() }
}
function First($x) { return @($x) | Select-Object -First 1 }
function Distinct($x) {
    $v=@($x | Where-Object { $_ -and (([string]$_).Trim()) } | Select-Object -Unique)
    if ($v.Count -eq 0) {return ''}
    return ($v -join ' | ')
}
function AddField([object]$o,[string]$n,[object]$v) {
    $o | Add-Member -MemberType NoteProperty -Name $n -Value (Clean $v)
}
function AskDefault([string]$name,[string]$old) {
    if ($old) { $s=Read-Host ($name + ' [Enter giu: ' + $old + ']') }
    else { $s=Read-Host $name }
    if (-not $s) { return $old }
    return $s.Trim()
}
function AskPosition {
    Write-Host ''
    Write-Host '  +------------------------------------------------------------+' -ForegroundColor DarkCyan
    Write-Host '  | THONG TIN VI TRI - NHAP MOT LAN CHO MOI MAY              |' -ForegroundColor Cyan
    Write-Host '  +------------------------------------------------------------+' -ForegroundColor DarkCyan
    $prev=@('','')
    if (Test-Path -LiteralPath $script:settings) {
        try {
            $ls=@(Get-Content -LiteralPath $script:settings -Encoding UTF8)
            if ($ls.Count -gt 0) {$prev[0]=$ls[0]}
            if ($ls.Count -gt 1) {$prev[1]=$ls[1]}
        } catch {}
    }
    $floor = AskDefault 'Ten tang (vi du: Tang 1)' $prev[0]
    $room  = AskDefault 'Phong ban (vi du: Tiep nhan)' $prev[1]
    if (-not $floor -or -not $room) { throw 'Ban can nhap Ten tang va Phong ban.' }
    @($floor,$room) | Set-Content -LiteralPath $script:settings -Encoding UTF8
    do {
        $kind=Read-Host 'Loai may: [1] Laptop  [2] PC'
        if ($kind -match '^(1|(?i:laptop))$') { $type='Laptop'; break }
        if ($kind -match '^(2|(?i:pc))$') { $type='PC'; break }
        Write-Host 'Vui long nhap 1 hoac 2.' -ForegroundColor Yellow
    } while ($true)
    $owner=Read-Host 'Nguoi phu trach may (de trong neu chua biet)'
    if (-not $owner -or -not $owner.Trim()) {
        $owner='CHUA_XAC_DINH'
        SurveyLog 'Chua nhap Nguoi phu trach may; se danh dau de kiem tra.' 'WARN'
    }
    return @($floor,$room,$type,$owner.Trim())
}

function Get-QuickOfficeFallback {
    # IMPORTANT: never rerun GetLicenseSummary in the main process!
    # A frozen Office provider would hang a second time and block the report.
    # Quick executable presence check uses only Test-Path (no WMI).
    $quickWord=$false; $quickExcel=$false; $quickOther=$false
    foreach($basePath in @($env:ProgramFiles,${env:ProgramFiles(x86)},$env:ProgramW6432)) {
        if(-not $basePath){continue}
        foreach($rel in @('Microsoft Office\root','Microsoft Office')) {
            foreach($ver in @('Office16','Office15','Office14','Office12')) {
                $dir=Join-Path (Join-Path $basePath $rel) $ver
                if(Test-Path -LiteralPath (Join-Path $dir 'WINWORD.EXE') -PathType Leaf){$quickWord=$true}
                if(Test-Path -LiteralPath (Join-Path $dir 'EXCEL.EXE') -PathType Leaf){$quickExcel=$true}
                foreach($otherExe in @('POWERPNT.EXE','OUTLOOK.EXE','MSACCESS.EXE')) {
                    if(Test-Path -LiteralPath (Join-Path $dir $otherExe) -PathType Leaf){$quickOther=$true}
                }
            }
        }
    }
    # Office may be installed outside the standard folders (App Paths).
    foreach($app in @('WINWORD.EXE','EXCEL.EXE','POWERPNT.EXE','OUTLOOK.EXE','MSACCESS.EXE')) {
        foreach($regRoot in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths',
                              'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\App Paths')) {
            try {
                $key=Get-Item -LiteralPath (Join-Path $regRoot $app) -ErrorAction Stop
                $path=[Environment]::ExpandEnvironmentVariables(([string]$key.GetValue('')).Trim().Trim('"'))
                if($path -and (Test-Path -LiteralPath $path -PathType Leaf)) {
                    if($app -eq 'WINWORD.EXE'){$quickWord=$true}
                    elseif($app -eq 'EXCEL.EXE'){$quickExcel=$true}
                    else{$quickOther=$true}
                }
            } catch { }
        }
    }
    $quickOffice=($quickWord -or $quickExcel -or $quickOther)
    $presence='KHONG_PHAT_HIEN'
    $officeState='KHONG_PHAT_HIEN'
    if($quickOffice){$presence='CO';$officeState='CAN_XAC_MINH_OFFICE_LICENSE'}
    $wordState='KHONG_PHAT_HIEN';$excelState='KHONG_PHAT_HIEN'
    $wordFlag='KHONG_PHAT_HIEN';$excelFlag='KHONG_PHAT_HIEN'
    if($quickWord){$wordFlag='CO';$wordState='CAN_XAC_MINH_WORD_LICENSE'}
    if($quickExcel){$excelFlag='CO';$excelState='CAN_XAC_MINH_EXCEL_LICENSE'}
    SurveyLog 'License worker qua thoi gian/loi; KHONG quet lai WMI. Kiem tra Office nhanh va tiep tuc khao sat.' 'WARN'
    $lic=New-Object PSObject -Property @{
        WindowsState='CAN_XAC_MINH_WINDOWS_LICENSE';WindowsEdition='';WindowsChannel='';
        WindowsMethod='LICENSE_WORKER_TIMEOUT_OR_ERROR';WindowsNote='Chua doc duoc do timeout/loi';
        OfficeInstalled=$presence;WordInstalled=$wordFlag;ExcelInstalled=$excelFlag;
        WordState=$wordState;ExcelState=$excelState;OfficeState=$officeState;
        OfficeProduct='';OfficeMethod='LICENSE_WORKER_TIMEOUT_OR_ERROR';
        OfficeNote='Khong the hoan tat kiem tra license trong gioi han; can xac minh thu cong neu da cai Office.'
    }
    return $lic
}

# Shared manual prompt (Windows PowerShell 2.0 compatible).
# Can run immediately at the LICENSE timeout while other workers continue.
function Invoke-OfficeManualPrompt([object]$lic) {
    if (-not $lic) { return $false }
    if ($lic -and $lic.OfficeState -match '^CAN_XAC_MINH' -and
        ($lic.OfficeInstalled -eq 'CO' -or $lic.WordInstalled -eq 'CO' -or $lic.ExcelInstalled -eq 'CO')) {
        Write-Host ''
        Write-Host '  OFFICE - XAC NHAN TRANG THAI KICH HOAT' -ForegroundColor Yellow
        Write-Host '  Chi chon 1 hoac 2 sau khi da kiem tra Office truc tiep.' -ForegroundColor Gray
        Write-Host '  [1] DA_KICH_HOAT' -ForegroundColor Green
        Write-Host '  [2] CHUA_KICH_HOAT' -ForegroundColor Yellow
        Write-Host '  [3] CAN_XAC_MINH (Enter de bo qua)' -ForegroundColor Gray
        $officeChoice = ''
        do {
            $reply = Read-Host 'Chon 1/2/3 (hoac nhap: co / chua)'
            $answer = ([string]$reply).Trim().ToLowerInvariant()
            # Accept Vietnamese accented input too, while remaining compatible
            # with Windows PowerShell 2.0 / .NET Framework on Windows 7.
            try {
                $decomposed = $answer.Normalize([System.Text.NormalizationForm]::FormD)
                $plain = New-Object System.Text.StringBuilder
                foreach ($ch in $decomposed.ToCharArray()) {
                    if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
                        [void]$plain.Append($ch)
                    }
                }
                $answer = $plain.ToString().Replace('đ','d')
            } catch { }
            $answer = [regex]::Replace($answer,'\s+',' ').Trim()
            if ($answer -match '^(1|y|yes|co|da|da kich hoat|co ban quyen)$') {
                $officeChoice = 'DA_KICH_HOAT'
                break
            }
            if ($answer -match '^(2|0|n|no|chua|khong|chua kich hoat|chua co|chua co ban quyen|chua co license|khong kich hoat|khong co ban quyen|khong co license)$') {
                $officeChoice = 'CHUA_KICH_HOAT'
                break
            }
            if ($answer -match '^(3|\?|khong ro|chua kiem tra|chua xac minh|can xac minh)$' -or -not $answer) {
                $officeChoice = 'CAN_XAC_MINH'
                break
            }
            Write-Host '  Lua chon khong hop le. Nhap 1, 2, 3 hoac Enter.' -ForegroundColor Yellow
        } while ($true)
        if ($officeChoice -eq 'DA_KICH_HOAT' -or $officeChoice -eq 'CHUA_KICH_HOAT') {
            $lic.OfficeState = $officeChoice
            if ($lic.WordInstalled -eq 'CO') { $lic.WordState = $officeChoice }
            if ($lic.ExcelInstalled -eq 'CO') { $lic.ExcelState = $officeChoice }
            $lic.OfficeMethod = 'XAC_NHAN_THU_CONG'
            $lic.OfficeNote = 'Nguoi khao sat xac nhan thu cong Office/Word/Excel: '+$officeChoice+'. Khong phai ket qua tu dong tu cong cu cap phep.'
            SurveyLog ('Office: '+$officeChoice+' (XAC_NHAN_THU_CONG; da ghi vao ban ghi khao sat)') 'WARN'
        } else {
            SurveyLog 'Office: khong xac nhan thu cong; giu CAN_XAC_MINH.' 'WARN'
        }
        return $true
    }
    return $false
}


try {
    if (-not (Test-Path -LiteralPath $script:rawFolder)) { New-Item -ItemType Directory -Path $script:rawFolder -Force | Out-Null }
    if (-not (Test-Path -LiteralPath $script:logFolder)) { New-Item -ItemType Directory -Path $script:logFolder -Force | Out-Null }
    $logPart=($env:COMPUTERNAME -replace '[^a-zA-Z0-9_-]','_')+'_'+(Get-Date -Format 'yyyyMMdd_HHmmss_fff')
    $script:logFile=Join-Path $script:logFolder ('SurveyLog_'+$logPart+'.txt')
    Show-AuditHeader
    SurveyLog 'BAT DAU KHAO SAT' 'STEP'
    SurveyLog ('PowerShell: '+$PSVersionTable.PSVersion.ToString()+' | May: '+$env:COMPUTERNAME+' | OS: '+[Environment]::OSVersion.Version.ToString())
    if($PSVersionTable.PSVersion.Major -lt 2) { throw 'Can Windows PowerShell 2.0 tro len.' }
    # Mandatory admin gate already passed before any directory or file was touched.
    SurveyLog 'Da xac nhan Administrator (elevated); cho phep bat dau khao sat.' 'OK'
    SurveyLog ('Ten may Windows: '+$env:COMPUTERNAME)
    SurveyLog 'BUOC 1/7 - Nhap thong tin tang, phong ban, loai may, nguoi phu trach' 'STEP'
    $location=@(AskPosition)
    SurveyLog ('Vi tri: '+$location[0]+' | Phong: '+$location[1]+' | Loai: '+$location[2])
    SurveyLog 'BUOC 2/7 - Lay thong tin phan cung va he dieu hanh qua WMI' 'STEP'
    SurveyLog 'Dang doc ComputerSystem, ComputerSystemProduct, BIOS, BaseBoard, CPU, Windows...'
    $cs=First (Wmi 'Win32_ComputerSystem')
    $pr=First (Wmi 'Win32_ComputerSystemProduct')
    $bios=First (Wmi 'Win32_BIOS')
    $board=First (Wmi 'Win32_BaseBoard')
    $cpu=First (Wmi 'Win32_Processor')
    $os=First (Wmi 'Win32_OperatingSystem')
    $case=First (Wmi 'Win32_SystemEnclosure')
    SurveyLog 'BUOC 3/7 - Doc o cung va card mang (MAC, IP)' 'STEP'
    $disks=@(Wmi 'Win32_DiskDrive')
    $netcards=@(Wmi 'Win32_NetworkAdapter')
    $configs=@(Wmi 'Win32_NetworkAdapterConfiguration')

    $ram=''
    if ($cs -and $cs.TotalPhysicalMemory) { $ram=[math]::Round(([double]$cs.TotalPhysicalMemory/1GB),1) }
    $diskInfo=@()
    foreach ($disk in $disks) {
        if ($disk.Model) {
            $sz=''
            if ($disk.Size) { $sz=[string]([math]::Round(([double]$disk.Size/1GB),0))+'GB' }
            $diskInfo += ((Clean $disk.Model) + ' ' + $sz)
        }
    }
    $cfgMap=@{}
    foreach ($cfg in $configs) { $cfgMap[[string]$cfg.Index]=$cfg }
    $physical=@($netcards | Where-Object { $_.MACAddress -and ($_.PhysicalAdapter -eq $true) })
    if ($physical.Count -eq 0) { $physical=@($netcards | Where-Object { $_.MACAddress -and $_.NetConnectionID }) }
    $lan=@(); $wifi=@(); $other=@(); $ips=@(); $gw=@()
    foreach ($nic in $physical) {
        $name=(Clean ([string]$nic.Name+' '+[string]$nic.NetConnectionID))
        $mac=(Clean $nic.MACAddress)
        if ($name -match '(?i)wi.?fi|wireless|wlan|802\.11') { $wifi += $mac }
        elseif ($name -match '(?i)ethernet|gigabit|\bgbe\b|\blan\b|netxtreme|realtek.*pcie') { $lan += $mac }
        else { $other += ($name+'='+$mac) }
        $cfg=$cfgMap[[string]$nic.Index]
        if ($cfg -and $cfg.IPEnabled) {
            foreach ($ip in @($cfg.IPAddress)) { if ($ip -match '^\d{1,3}(\.\d{1,3}){3}$') { $ips += $ip } }
            $gw += @($cfg.DefaultIPGateway)
        }
    }
    if (@($ips).Count -eq 0) {
        foreach($c in $configs) {
            if ($c.IPEnabled) { foreach($ip in @($c.IPAddress)) { if($ip -match '^\d{1,3}(\.\d{1,3}){3}$') {$ips += $ip} } }
        }
    }
    SurveyLog ('Da doc '+$disks.Count+' o cung va '+$netcards.Count+' card mang.') 'OK'
    SurveyLog 'BUOC 4/7 - Kiem tra tinh day du du lieu' 'STEP'
    $missing=@()
    if ($location[3] -eq 'CHUA_XAC_DINH') { $missing += 'Nguoi phu trach may' }
    if (-not (Real $cs.Manufacturer)) { $missing += 'Hang may' }
    if (-not (Real $cs.Model) -and -not (Real $pr.Name)) { $missing += 'Model may' }
    if (-not (Real $bios.SerialNumber) -and -not (Real $pr.IdentifyingNumber)) { $missing += 'Serial may' }
    if (-not (Real $board.Product)) { $missing += 'Model mainboard' }
    if (-not (Real $pr.UUID)) { $missing += 'UUID' }
    if (-not (Real $board.Manufacturer)) { $missing += 'Hang mainboard' }
    if (-not (Real $board.SerialNumber)) { $missing += 'Serial mainboard' }
    if ($diskInfo.Count -eq 0) { $missing += 'Thong tin o luu tru' }
    if (-not (Real $cpu.Name)) { $missing += 'CPU' }
    if (-not (Real $ram)) { $missing += 'RAM' }
    if (-not (Real $os.BuildNumber)) { $missing += 'Windows build' }
    if (-not (Real (Distinct (@($lan)+@($wifi)+@($other))))) { $missing += 'MAC card mang' }

    # Parallelize only independent audit groups; step 1-4 data collection stays sequential.
    # Child processes NEVER write to USB; only the main process writes snapshot/Excel once.
    $parallelMode='PARALLEL'
    $parallelResult=$null
    $auditTimings=@{}
    $script:licenseTimeoutFallback=$null
    $script:officeManualAlreadyAsked=$false
    $useParallel=$true
    if($env:AUDIT_PARALLEL -eq '0'){$useParallel=$false}
    if($useParallel){
        SurveyLog 'BUOC 5/7 - CHAY SONG SONG 3 NHOM: LICENSE | SECURITY | POLICIES' 'STEP'
        try {
            . (Join-Path $script:root 'ParallelAudit.ps1')
            $parallelResult=Invoke-ParallelAudit -Root $script:root -OnLicenseTimeout {
                # This callback runs in the interactive parent, not inside the
                # -NonInteractive LICENSE worker. Other workers keep running.
                $script:licenseTimeoutFallback=Get-QuickOfficeFallback
                if ($script:licenseTimeoutFallback.OfficeInstalled -eq 'CO' -or
                    $script:licenseTimeoutFallback.WordInstalled -eq 'CO' -or
                    $script:licenseTimeoutFallback.ExcelInstalled -eq 'CO') {
                    SurveyLog 'LICENSE qua thoi gian; Office co cai, hien bang nhap ngay.' 'WARN'
                    # Mark as handled before Read-Host to prevent a second prompt
                    # if the console input fails or is interrupted.
                    $script:officeManualAlreadyAsked=$true
                    [void](Invoke-OfficeManualPrompt $script:licenseTimeoutFallback)
                } else {
                    SurveyLog 'LICENSE qua thoi gian; khong phat hien Office, khong hoi thu cong.' 'INFO'
                }
            }
            $parallelMode=$parallelResult.Mode
            $auditTimings=$parallelResult.Timings
            foreach($e in @($parallelResult.Errors)){ $script:errs += ('Parallel: '+$e) }
            $lic=$parallelResult.License
            $security=$parallelResult.Security
            $extra=$parallelResult.Extra
            $compliance=$parallelResult.Compliance
            $script:SuspectInstalledApps=@($parallelResult.Apps)
        }catch{
            $script:errs += ('Parallel: '+$_.Exception.Message)
            SurveyLog ('Khong chay duoc song song: '+$_.Exception.Message) 'WARN'
            $parallelMode='PARALLEL_FATAL_ERROR'
            $lic=$null;$security=$null;$extra=$null;$compliance=$null
            $script:SuspectInstalledApps=@()
        }
    } else {
        SurveyLog 'BUOC 5/7 - CHE DO LAN LUOT (AUDIT_PARALLEL=0)' 'STEP'
        SurveyLog 'BUOC 5/7 - Kiem tra kich hoat Windows va Office' 'STEP'
        $lic=$null
        try {
            . (Join-Path $script:root 'LicenseCheck.ps1')
            $lic=GetLicenseSummary
        } catch { $script:errs += ('Ban quyen: '+$_.Exception.Message); SurveyLog ('Loi kiem tra ban quyen: '+$_.Exception.Message) 'WARN' }

        SurveyLog 'BUOC 5B/7 - Kiem tra bao mat, phan mem nghi van, antivirus, TPM, BitLocker' 'STEP'
        $security=$null
        try {
            . (Join-Path $script:root 'SecurityAudit.ps1')
            $security=GetSecurityAudit -rulesPath (Join-Path $script:root 'SuspiciousSoftwareRules.csv')
        } catch {
            $script:errs += ('SecurityAudit: '+$_.Exception.Message)
            SurveyLog ('Khong the hoan thanh kiem tra bao mat: '+$_.Exception.Message) 'WARN'
        }
        if (-not $security) {
            $security=New-Object PSObject -Property @{
                Suspicious='CAN_XAC_MINH'; Findings=''; Remote='CAN_XAC_MINH'; Admin='CAN_XAC_MINH'
                AdminElevated='CAN_XAC_MINH'; Antivirus='CAN_XAC_MINH'; AVRealTime='CAN_XAC_MINH'
                AVUpdated='CAN_XAC_MINH'; AVSource=''; AVNote=''; DefenderSignature='';
                AVProductCount='CAN_XAC_MINH';AVVersions='CAN_XAC_MINH';AVRealTimeDetails='CAN_XAC_MINH';
                AVUpdateDetails='CAN_XAC_MINH';AVServiceDetails='CAN_XAC_MINH';
                AVLicenseDetails='CAN_XAC_MINH';AVEvidenceDetails='CAN_XAC_MINH';
                AVScanDetails='CAN_XAC_MINH';AVEngineDetails='CAN_XAC_MINH';
                Kaspersky='CAN_XAC_MINH'; TPM='CAN_XAC_MINH'; TPMSpec=''; TPMDetail=''; TPMSource='';
                BitLocker='CAN_XAC_MINH'; BitLockerConversion=''; BitLockerProtection='';
                BitLockerPercent=''; BitLockerDetail='SecurityAudit khong chay duoc'; BitLockerSource=''
            }
        }
        SurveyLog 'BUOC 5C/7 - Firewall, HotFix, USB Storage, Guest, DNS va Screen Lock' 'STEP'
        $extra=$null
        try {
            . (Join-Path $script:root 'SecurityExtras.ps1')
            $extra=GetExtraSecurityAudit
        } catch {
            $script:errs += ('SecurityExtras: '+$_.Exception.Message)
            SurveyLog ('Loi kiem tra cau hinh bo sung: '+$_.Exception.Message) 'WARN'
        }
        if(-not $extra) {
            $extra=New-Object PSObject -Property @{
              Firewall='CAN_XAC_MINH';FirewallDomain='';FirewallPrivate='';FirewallPublic='';FirewallCurrentProfiles='';FirewallSource='';
              Hotfix='CAN_XAC_MINH';HotfixKB='';HotfixDate='';Hotfix3='';HotfixCount='';HotfixSource='';
              USB='CAN_XAC_MINH';USBSTOR='';USBDenyAll='';USBDenyRead='';USBDenyWrite='';USBWriteProtect='';
              Guest='CAN_XAC_MINH';GuestName='';GuestEnabled='';DNS='CAN_XAC_MINH';DNSPerAdapter='';
              ScreenLock='CAN_XAC_MINH';ScreenSaverActive='';ScreenSaverSecure='';ScreenSaverTimeout='';MachineInactivity='';
              ScreenLockUser='';Errors='SecurityExtras khong chay duoc'
            }
        }
        SurveyLog 'BUOC 5D/7 - Domain, SMBv1, Autorun, NTP, DHCP/Static (read-only)' 'STEP'
        $compliance=$null
        try {
            . (Join-Path $script:root 'ComplianceExtras.ps1')
            $compliance=GetComplianceExtras -cs $cs -configs $configs
        } catch {
            $script:errs += ('ComplianceExtras: '+$_.Exception.Message)
            SurveyLog ('Loi kiem tra PA05 bo sung: '+$_.Exception.Message) 'WARN'
        }
        if(-not $compliance) {
            $compliance=New-Object PSObject -Property @{
                DomainStatus='CAN_XAC_MINH';DomainName='';DomainError='LOI_MODULE';
                SMB1='CAN_XAC_MINH';SMB1Server='';SMB1Client='';SMB1Note='';SMB1Error='LOI_MODULE';
                AutoRun='CAN_XAC_MINH';AutoRunHKLM='';AutoRunHKCU='';AutoRunUserPreference='';AutoRunPolicyEvidence='';AutoRunCheckedUser='';AutoRunNote='';AutoRunError='LOI_MODULE';
                NTP='CAN_XAC_MINH';NTPService='';NTPType='';NTPServer='';NTPSource='';NTPLastSync='';NTPNote='';NTPError='LOI_MODULE';
                IPMode='CAN_XAC_MINH';NetScan='CAN_DOI_CHIEU_SWITCH_CMBD';NAC='CAN_DOI_CHIEU_SWITCH_NAC';Notes='LOI_MODULE'
            }
        }
        SurveyLog ('Domain: '+$compliance.DomainStatus+' '+$compliance.DomainName+' | SMBv1: '+$compliance.SMB1)
        SurveyLog ('AutoRun: '+$compliance.AutoRun+' | NTP: '+$compliance.NTP)
        if($compliance.SMB1Error -or $compliance.AutoRunError -or $compliance.NTPError) {
            $script:errs+=@($compliance.SMB1Error,$compliance.AutoRunError,$compliance.NTPError | Where-Object {$_})
        }
        $parallelMode='SEQUENTIAL'
    }
    if($useParallel){
        if(-not $lic){
            # Reuse the fallback prepared at the 40-second timeout (and possibly
            # already confirmed by the operator). Never perform another WMI scan.
            if ($script:licenseTimeoutFallback) {
                $lic=$script:licenseTimeoutFallback
            } else {
                $lic=Get-QuickOfficeFallback
            }
        }
        if(-not $security){
            SurveyLog 'Security worker khong co du lieu; danh dau CAN_XAC_MINH.' 'WARN'
            $security=New-Object PSObject -Property @{
                Suspicious='CAN_XAC_MINH';Findings='';Remote='CAN_XAC_MINH';Admin='CAN_XAC_MINH';
                AdminElevated='CAN_XAC_MINH';Antivirus='CAN_XAC_MINH';AVRealTime='CAN_XAC_MINH';
                AVUpdated='CAN_XAC_MINH';AVSource='';AVNote='Worker chua co ket qua';DefenderSignature='';
                AVProductCount='CAN_XAC_MINH';AVVersions='CAN_XAC_MINH';AVRealTimeDetails='CAN_XAC_MINH';
                AVUpdateDetails='CAN_XAC_MINH';AVServiceDetails='CAN_XAC_MINH';
                AVLicenseDetails='CAN_XAC_MINH';AVEvidenceDetails='CAN_XAC_MINH';
                AVScanDetails='CAN_XAC_MINH';AVEngineDetails='CAN_XAC_MINH';
                Kaspersky='CAN_XAC_MINH';TPM='CAN_XAC_MINH';TPMSpec='';TPMDetail='Worker chua co ket qua';TPMSource='';
                BitLocker='CAN_XAC_MINH';BitLockerConversion='';BitLockerProtection='';
                BitLockerPercent='';BitLockerDetail='Worker chua co ket qua';BitLockerSource=''
            }
        }
        if(-not $extra){
            $extra=New-Object PSObject -Property @{
              Firewall='CAN_XAC_MINH';FirewallDomain='';FirewallPrivate='';FirewallPublic='';FirewallCurrentProfiles='';FirewallSource='';
              Hotfix='CAN_XAC_MINH';HotfixKB='';HotfixDate='';Hotfix3='';HotfixCount='';HotfixSource='';
              USB='CAN_XAC_MINH';USBSTOR='';USBDenyAll='';USBDenyRead='';USBDenyWrite='';USBWriteProtect='';
              Guest='CAN_XAC_MINH';GuestName='';GuestEnabled='';DNS='CAN_XAC_MINH';DNSPerAdapter='';
              ScreenLock='CAN_XAC_MINH';ScreenSaverActive='';ScreenSaverSecure='';ScreenSaverTimeout='';MachineInactivity='';
              ScreenLockUser='';Errors='Policies worker chua co ket qua'
            }
        }
        if(-not $compliance){
            $compliance=New-Object PSObject -Property @{
                DomainStatus='CAN_XAC_MINH';DomainName='';DomainError='Policies worker chua co ket qua';
                SMB1='CAN_XAC_MINH';SMB1Server='';SMB1Client='';SMB1Note='';SMB1Error='';
                AutoRun='CAN_XAC_MINH';AutoRunHKLM='';AutoRunHKCU='';AutoRunUserPreference='';AutoRunPolicyEvidence='';AutoRunCheckedUser='';AutoRunNote='';AutoRunError='';
                NTP='CAN_XAC_MINH';NTPService='';NTPType='';NTPServer='';NTPSource='';NTPLastSync='';NTPNote='';NTPError='';
                IPMode='CAN_XAC_MINH';NetScan='CAN_DOI_CHIEU_SWITCH_CMDB';NAC='CAN_DOI_CHIEU_SWITCH_NAC';Notes='Policies worker chua co ket qua'
            }
        }
        if($extra.Errors){$script:errs += ('SecurityExtras: '+$extra.Errors)}
        if($compliance.SMB1Error -or $compliance.AutoRunError -or $compliance.NTPError){
            $script:errs+=@($compliance.SMB1Error,$compliance.AutoRunError,$compliance.NTPError | Where-Object {$_})
        }
        SurveyLog ('BUOC 5 HOAN TAT - '+$parallelMode+' | Windows: '+$lic.WindowsState+' | Office: '+$lic.OfficeState) 'STEP'
    }
    # The timeout callback may already have shown the prompt while other
    # workers were running. Never ask the same Office question twice.
    if (-not $script:officeManualAlreadyAsked) {
        $script:officeManualAlreadyAsked=[bool](Invoke-OfficeManualPrompt $lic)
    }
    # Warn about an obvious mismatch; never overwrite the user's survey input.
    try {
        $chassis=@($case.ChassisTypes)
        $isPortable=$false
        foreach($typ in $chassis) { if(@(8,9,10,11,12,14,18,21,30,31,32) -contains [int]$typ) { $isPortable=$true } }
        if($chassis.Count -gt 0 -and $location[2] -eq 'Laptop' -and (-not $isPortable) -and ($chassis -contains 3 -or $chassis -contains 4 -or $chassis -contains 6 -or $chassis -contains 7 -or $chassis -contains 15)) {
            SurveyLog 'Canh bao: ban chon Laptop nhung ChassisTypes goi y Desktop/PC. Hay doi chieu loai may.' 'WARN'
            $missing += 'Loai may can doi chieu'
        }
    } catch {}
    $winRelease=''
    try {
        $cv=Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
        if($cv.DisplayVersion) {$winRelease=Clean $cv.DisplayVersion}
        elseif($cv.ReleaseId) {$winRelease=Clean $cv.ReleaseId}
    } catch {}
    $windowsLabel=(Clean $os.Caption)
    if($winRelease) {$windowsLabel += ' '+$winRelease}
    if($os.BuildNumber) {$windowsLabel += ' (Build '+(Clean $os.BuildNumber)+')'}
    $maker=Clean $cs.Manufacturer
    if (-not (Real $maker) -and (Real $pr.Vendor)) { $maker=Clean $pr.Vendor }
    $model=Clean $cs.Model
    if (-not (Real $model) -and (Real $pr.Name)) { $model=Clean $pr.Name }
    $sku=Clean $pr.SKUNumber
    if (-not (Real $sku) -and (Real $cs.SystemSKUNumber)) { $sku=Clean $cs.SystemSKUNumber }
    $serial=Clean $pr.IdentifyingNumber
    if (-not (Real $serial)) { $serial=Clean $bios.SerialNumber }
    # Win32_SystemEnclosure.SMBIOSAssetTag is SMBIOS Type 3 (chassis Asset Tag).
    # Do NOT substitute Win32_BIOS.SMBIOSBIOSVersion (Type 0 BIOS version).
    $asset=Clean $case.SMBIOSAssetTag
    if(-not $case) {
        $asset='CAN_XAC_MINH_WMI_ASSET_TAG'
        SurveyLog 'Khong truy van duoc Win32_SystemEnclosure; khong the xac minh Asset Tag.' 'WARN'
    } elseif(-not $asset) {
        $asset='KHONG_CO_GIA_TRI_SMBIOS'
        SurveyLog 'Win32_SystemEnclosure.SMBIOSAssetTag rong; khong suy ra Asset Tag tu BIOS version hay Serial.' 'WARN'
    } elseif(-not (Real $asset)) {
        SurveyLog ('SMBIOS Asset Tag la gia tri mac dinh: '+$asset+'; can doi chieu tem tai san.') 'WARN'
    }

    if($extra.Errors){$script:errs += ('SecurityExtras: '+$extra.Errors)}
    $row=New-Object PSObject
    AddField $row 'Thời gian' (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    AddField $row 'Tầng' $location[0]
    AddField $row 'Phòng ban' $location[1]
    AddField $row 'Người phụ trách máy' $location[3]
    AddField $row 'Loại máy' $location[2]
    AddField $row 'Tên máy' $env:COMPUTERNAME
    AddField $row 'Hãng máy' $maker
    AddField $row 'Model máy' $model
    AddField $row 'Product Number - SKU' $sku
    AddField $row 'Serial máy - Service Tag' $serial
    AddField $row 'Serial BIOS' $bios.SerialNumber
    AddField $row 'Mã tài sản SMBIOS' $asset
    AddField $row 'UUID' $pr.UUID
    AddField $row 'Hãng mainboard' $board.Manufacturer
    AddField $row 'Model mainboard' $board.Product
    # Preserve exactly what SMBIOS reports (including 'Default string').
    # Distinct field name requested for the export; same Win32_BaseBoard.Product source.
    AddField $row 'BaseBoard Product' $board.Product
    AddField $row 'Serial mainboard' $board.SerialNumber
    AddField $row 'Phiên bản BIOS' $bios.SMBIOSBIOSVersion
    AddField $row 'CPU' $cpu.Name
    AddField $row 'RAM (GB)' $ram
    AddField $row 'Ổ lưu trữ (model - GB)' (Distinct $diskInfo)
    AddField $row 'Windows (loại/edition)' $os.Caption
    AddField $row 'Windows build' $os.BuildNumber
    AddField $row 'Windows version' $os.Version
    AddField $row 'Windows OS - Version/Build' $windowsLabel
    AddField $row 'Windows Release' $winRelease
    AddField $row 'Quyền Admin Local Audit' $security.Admin
    AddField $row 'Tiến trình đang chạy Admin' $security.AdminElevated
    AddField $row 'Rà soát phần mềm Crack / Lạ' $security.Suspicious
    AddField $row 'Dấu hiệu nghi vấn chi tiết' $security.Findings
    AddField $row 'Công cụ điều khiển từ xa' $security.Remote
    AddField $row 'Antivirus đang dùng' $security.Antivirus
    AddField $row 'Trạng thái Antivirus Real-time' $security.AVRealTime
    AddField $row 'Ngày cập nhật mẫu virus' $security.AVUpdated
    AddField $row 'Nguồn dữ liệu Antivirus' $security.AVSource
    AddField $row 'Ghi chú Antivirus' $security.AVNote
    # Keep verbose AV diagnostics in the scan log, not the Excel report.
    if($security.AVNote -and $script:logFile) {
        try {
            $avLine='['+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')+'][DETAIL] Antivirus: '+(Clean $security.AVNote)
            Add-Content -LiteralPath $script:logFile -Value $avLine -Encoding UTF8 -ErrorAction Stop
        } catch {}
    }
    AddField $row 'Ngày mẫu virus Defender (tham khảo)' $security.DefenderSignature
    AddField $row 'Antivirus - Số sản phẩm phát hiện' $security.AVProductCount
    AddField $row 'Antivirus - Phiên bản từng sản phẩm' $security.AVVersions
    AddField $row 'Antivirus - Real-time từng sản phẩm' $security.AVRealTimeDetails
    AddField $row 'Antivirus - Cập nhật từng sản phẩm' $security.AVUpdateDetails
    AddField $row 'Antivirus - Dịch vụ từng sản phẩm' $security.AVServiceDetails
    AddField $row 'Antivirus - Giấy phép từng sản phẩm' $security.AVLicenseDetails
    AddField $row 'Antivirus - Nguồn kiểm tra chi tiết' $security.AVEvidenceDetails
    AddField $row 'Antivirus - Lần quét gần nhất từng sản phẩm' $security.AVScanDetails
    AddField $row 'Antivirus - Engine / Signature từng sản phẩm' $security.AVEngineDetails
    AddField $row 'Chip TPM 2.0' $security.TPM
    AddField $row 'Phiên bản TPM' $security.TPMSpec
    AddField $row 'Mã hóa BitLocker (ổ Windows)' $security.BitLocker
    AddField $row 'BitLocker Conversion' $security.BitLockerConversion
    AddField $row 'BitLocker Protection' $security.BitLockerProtection
    AddField $row 'BitLocker Encrypted Percent' $security.BitLockerPercent
    AddField $row 'BitLocker Source' $security.BitLockerSource
    AddField $row 'TPM Source' $security.TPMSource
    AddField $row 'Windows PowerShell' $PSVersionTable.PSVersion.ToString()
    AddField $row 'Kiểu thực thi Step 5' $parallelMode
    AddField $row 'Step 5 LICENSE (giây)' $auditTimings['LICENSE']
    AddField $row 'Step 5 SECURITY (giây)' $auditTimings['SECURITY']
    AddField $row 'Step 5 POLICIES (giây)' $auditTimings['POLICIES']
    AddField $row 'Domain / Workgroup' $compliance.DomainStatus
    AddField $row 'Tên Domain / Workgroup' $compliance.DomainName
    AddField $row 'SMBv1 trạng thái cấu hình' $compliance.SMB1
    AddField $row 'SMBv1 Server' $compliance.SMB1Server
    AddField $row 'SMBv1 Client' $compliance.SMB1Client
    AddField $row 'Ghi chú SMBv1' $compliance.SMB1Note
    AddField $row 'Autorun / Autoplay Policy' $compliance.AutoRun
    AddField $row 'Autorun Registry HKLM' $compliance.AutoRunHKLM
    AddField $row 'Autorun Registry HKCU' $compliance.AutoRunHKCU
    AddField $row 'Ghi chú Autorun' $compliance.AutoRunNote
    AddField $row 'AutoPlay UI (HKCU)' $compliance.AutoRunUserPreference
    AddField $row 'Autorun - Chính sách chi tiết' $compliance.AutoRunPolicyEvidence
    AddField $row 'AutoPlay - Tài khoản kiểm tra' $compliance.AutoRunCheckedUser
    AddField $row 'NTP W32Time' $compliance.NTP
    AddField $row 'NTP W32Time Service' $compliance.NTPService
    AddField $row 'NTP Client Type' $compliance.NTPType
    AddField $row 'NTP Server Config' $compliance.NTPServer
    AddField $row 'NTP Source' $compliance.NTPSource
    AddField $row 'NTP Last Successful Sync' $compliance.NTPLastSync
    AddField $row 'NTP Notes' $compliance.NTPNote
    AddField $row 'IP DHCP hoặc Static' $compliance.IPMode
    AddField $row 'Rà soát IP/MAC toàn mạng' $compliance.NetScan
    AddField $row 'NAC Port Security' $compliance.NAC
    AddField $row 'Windows Firewall Status' $extra.Firewall

    AddField $row 'Firewall Current Profiles' $extra.FirewallCurrentProfiles
    AddField $row 'Firewall Source' $extra.FirewallSource
    AddField $row 'Windows HotFix gần nhất' $extra.Hotfix
    AddField $row 'HotFix KB' $extra.HotfixKB
    AddField $row 'HotFix InstalledOn' $extra.HotfixDate
    AddField $row 'HotFix 3 gần nhất' $extra.Hotfix3
    AddField $row 'HotFix Count' $extra.HotfixCount
    AddField $row 'HotFix Source' $extra.HotfixSource
    AddField $row 'USB Storage Policy' $extra.USB
    AddField $row 'USBSTOR Start' $extra.USBSTOR
    AddField $row 'USB Deny_All Policy' $extra.USBDenyAll
    AddField $row 'USB Deny_Read Policy' $extra.USBDenyRead
    AddField $row 'USB Deny_Write Policy' $extra.USBDenyWrite
    AddField $row 'USB WriteProtect' $extra.USBWriteProtect
    AddField $row 'Guest Account Status' $extra.Guest
    AddField $row 'Guest Account Name' $extra.GuestName
    AddField $row 'Guest Account Enabled' $extra.GuestEnabled
    AddField $row 'DNS Client Servers' $extra.DNS
    AddField $row 'DNS per Adapter' $extra.DNSPerAdapter
    AddField $row 'Screen Saver Lock Status' $extra.ScreenLock
    AddField $row 'ScreenSaveActive' $extra.ScreenSaverActive
    AddField $row 'ScreenSaverIsSecure' $extra.ScreenSaverSecure
    AddField $row 'ScreenSaveTimeOut' $extra.ScreenSaverTimeout
    AddField $row 'Machine Inactivity Limit Seconds' $extra.MachineInactivity
    AddField $row 'ScreenLock Registry User' $extra.ScreenLockUser

    # One visible Office column (suite, Word/Excel presence, activation status).
    $officeBits=@()
    if($lic.OfficeProduct) { $officeBits+=('San pham: '+(Clean $lic.OfficeProduct)) }
    $officeBits+=('Office: '+(Clean $lic.OfficeState))
    $wordDisplay=(Clean $lic.WordInstalled)
    $excelDisplay=(Clean $lic.ExcelInstalled)
    if($lic.WordInstalled -eq 'CO' -and $lic.WordState) { $wordDisplay+=' ['+(Clean $lic.WordState)+']' }
    if($lic.ExcelInstalled -eq 'CO' -and $lic.ExcelState) { $excelDisplay+=' ['+(Clean $lic.ExcelState)+']' }
    $officeBits+=('Word: '+$wordDisplay)
    $officeBits+=('Excel: '+$excelDisplay)
    if ($lic.OfficeMethod -eq 'XAC_NHAN_THU_CONG') { $officeBits+=('Nguon: XAC_NHAN_THU_CONG') }
    AddField $row 'Phần mềm Office' ($officeBits -join ' | ')
    AddField $row 'Phần mềm Kaspersky' $security.Kaspersky
    AddField $row 'Windows kích hoạt' $lic.WindowsState
    AddField $row 'Windows License Source' $lic.WindowsMethod
    AddField $row 'Windows License Note' $lic.WindowsNote
    AddField $row 'Office License Method' $lic.OfficeMethod
    AddField $row 'Office License Note' $lic.OfficeNote
    AddField $row 'Word đã cài' $lic.WordInstalled
    AddField $row 'Word kích hoạt' $lic.WordState
    AddField $row 'Excel đã cài' $lic.ExcelInstalled
    AddField $row 'Excel kích hoạt' $lic.ExcelState
    AddField $row 'Sản phẩm Office' $lic.OfficeProduct
    AddField $row 'MAC LAN' (Distinct $lan)
    AddField $row 'MAC Wi-Fi' (Distinct $wifi)
    AddField $row 'MAC khác' (Distinct $other)
    AddField $row 'Địa chỉ IPv4' (Distinct $ips)
    AddField $row 'Gateway' (Distinct $gw)
    # The Office license decision must not affect unrelated evidence.
    # Mark actually absent readings for review; do NOT fill blanks with fake data.
    if (-not (Real $security.AVSource)) { $missing += 'Nguon du lieu Antivirus' }
    if (-not (Real $extra.DNSPerAdapter)) { $missing += 'DNS per Adapter' }
    if (-not (Real (Distinct $gw))) { $missing += 'Gateway' }
    # Defender date can legitimately be blank when Defender is not the active
    # antivirus. Review it only when a Defender date is expected.
    if ($security.AVRealTime -match '(?i)Defender' -and -not $security.DefenderSignature) {
        $missing += 'Ngay mau virus Defender (tham khao)'
    }
    # Verification quality is separate from the actual security condition.
    # Never label 'DAY_DU' if a relevant check is inconclusive or worker failed.
    $reviewReasons=@()
    if($missing.Count -gt 0){$reviewReasons+='Thieu thong tin nhan dang'}
    if($script:errs.Count -gt 0){$reviewReasons+='Loi truy van'}
    if($parallelMode -match 'ERROR|LOI'){ $reviewReasons+='Tien trinh kiem tra song song loi' }
    if($security.Suspicious -match '^(CAN_RA_SOAT|QUET_QUA_THOI_GIAN|LOI_|CAN_XAC_MINH)'){$reviewReasons+='Phan mem nghi van / chua quet du'}
    if($security.AVRealTime -notmatch '^ACTIVE \(Defender\)$'){$reviewReasons+='Trang thai antivirus chua xac minh hoac can ra soat'}
    if($security.AVUpdated -match '^(CAN_XAC_MINH|CHUA_|LOI_)'){$reviewReasons+='Ngay cap nhat antivirus'}
    if($security.TPM -notmatch '^TPM_2\.0$'){$reviewReasons+='TPM (doi chieu phien ban / trang thai)'}
    if($security.BitLocker -notmatch '^ON - FULLY ENCRYPTED$'){$reviewReasons+='BitLocker (trang thai hoac cau hinh can ra soat)'}
    if($security.Admin -match 'CAN_XAC_MINH|CHUA_XAC_MINH'){$reviewReasons+='Quyen Local Admin'}
    if($lic.WindowsState -ne 'DA_KICH_HOAT'){$reviewReasons+='Ban quyen Windows'}
    if($lic.OfficeState -match '^CAN_XAC_MINH|^CHUA_|^THOI_GIAN_AN_HAN|^CAN_RA_SOAT_OFFICE_LICENSE'){$reviewReasons+='Ban quyen Office'}
    # SMB1 has a composite status starting with SERVER:, so test each component.
    if($compliance.SMB1Server -match 'CAN_XAC_MINH|CAN_DOI_CHIEU|KHONG_CAU_HINH|ENABLED_CONFIG' -or
       $compliance.SMB1Client -match 'CAN_XAC_MINH|CAN_DOI_CHIEU|KHONG_TIM_THAY|^START_|ENABLED_CONFIG' -or
       -not $compliance.SMB1Server -or -not $compliance.SMB1Client) {
        $reviewReasons+='SMBv1 Server / Client'
    }
    if($compliance.AutoRun -notmatch '^AUTOPLAY_POLICY_ALL_DRIVE_TYPES_DISABLED'){$reviewReasons+='Autorun / Autoplay (chua xac nhan day du policy)'}
    if(-not $compliance.NTPSource -or -not $compliance.NTPLastSync){$reviewReasons+='NTP chua xac nhan dong bo'}
    if($compliance.DomainStatus -match 'CAN_XAC_MINH'){$reviewReasons+='Domain / Workgroup'}
    if($compliance.IPMode -match 'KHONG_XAC_DINH|CHUA_XAC_MINH'){$reviewReasons+='IP DHCP / Static'}
    if($extra.Firewall -match 'CAN_XAC_MINH|OFF'){$reviewReasons+='Windows Firewall'}
    if($extra.Hotfix -match 'CAN_XAC_MINH|KHONG_TIM_THAY|KHONG_CO'){$reviewReasons+='Windows HotFix'}
    if($extra.USB -match 'CAN_XAC_MINH|USBSTOR_MANUAL|KHONG_CO_USBSTOR'){$reviewReasons+='USB Storage Policy'}
    if($extra.Guest -match 'CAN_XAC_MINH|GUEST_ENABLED|KHONG_TIM_THAY'){$reviewReasons+='Guest Account'}
    if($extra.DNS -match 'CAN_XAC_MINH|KHONG_CO_DNS|KHONG_CO_CARD'){$reviewReasons+='DNS Server'}
    if($extra.ScreenLock -notmatch 'MACHINE_INACTIVITY_LOCK=|SECURE_SCREENSAVER_LOCK='){$reviewReasons+='Screen Lock / Screen Saver'}
    if($extra.ScreenLock -match 'HKCU_CO_THE_LA_ADMIN_KHAC'){$reviewReasons+='Screen Saver khac tai khoan nguoi dung'}
    $reviewReasons=@($reviewReasons | Select-Object -Unique)
    AddField $row 'Mục cần đối chiếu' (Distinct $reviewReasons)
    AddField $row 'Trường còn thiếu' (Distinct $missing)


    # The suspicious scan worker has already read uninstall registry under a 25s watchdog.
    # Export its full installed-app list; never invoke Win32_Product (MSI repairs).
    $allApps=@($script:SuspectInstalledApps | Where-Object {$_.Name})
    $appStatus='CHUA_LAY_DUOC_DANH_SACH'
    $appFile=''
    $appCount=[string]$allApps.Count
    if($allApps.Count -gt 0) {
        try {
            $appFolder=Join-Path $script:folder 'AppInventory'
            if(-not (Test-Path $appFolder)){[void](New-Item -ItemType Directory -Path $appFolder -Force)}
            $appName='InstalledApps_'+($env:COMPUTERNAME -replace '[^a-zA-Z0-9_-]','_')+'_'+(Get-Date -Format 'yyyyMMdd_HHmmss_fff')+'.csv'
            $appFullPath=Join-Path $appFolder $appName
            Write-Utf8Csv -Records @($allApps | Sort-Object Name,Version -Unique) -FilePath $appFullPath
            $appStatus='DA_XUAT_DANH_SACH_REGISTRY'
            $appFile='AppInventory\'+$appName
            SurveyLog ('Danh sach ung dung: '+$appCount+' entries -> '+$appFile) 'OK'
        }catch{
            $appStatus='LOI_XUAT_APP_INVENTORY'
            $script:errs+=('AppInventory: '+$_.Exception.Message)
            SurveyLog ('Loi xuat danh sach ung dung: '+$_.Exception.Message) 'WARN'
        }
    } else {
        SurveyLog 'Khong lay duoc danh sach ung dung day du; co the quet qua thoi gian/bi gioi han quyen.' 'WARN'
    }
    AddField $row 'Trạng thái danh sách ứng dụng' $appStatus
    AddField $row 'Số ứng dụng tìm thấy' $appCount
    AddField $row 'File danh sách ứng dụng' $appFile
    if($appStatus -ne 'DA_XUAT_DANH_SACH_REGISTRY'){
        $row.'Mục cần đối chiếu'=((@($row.'Mục cần đối chiếu') + @('Danh sach ung dung chua day du') | Where-Object {$_}) -join ' | ')
    }
    SurveyLog 'BUOC 6/7 - Ghi du lieu goc vao USB' 'STEP'
    $filePart=(($env:COMPUTERNAME -replace '[^a-zA-Z0-9_-]','_') + '_' + (Get-Date -Format 'yyyyMMdd_HHmmss_fff') + '_' + [guid]::NewGuid().ToString('N').Substring(0,6))
    $snapshot=Join-Path $script:rawFolder ('Device_'+$filePart+'.csv')
    Write-Utf8Csv -Records @($row) -FilePath $snapshot
    # Snapshot invariant: fields exist regardless of whether Office asked a
    # manual question. Record these checks in the log for quick diagnostics.
    try {
        $checkFields=@('UUID','Hãng mainboard','Serial mainboard','CPU','RAM (GB)',
          'Ổ lưu trữ (model - GB)','Nguồn dữ liệu Antivirus','DNS per Adapter',
          'Gateway','Mục cần đối chiếu','Trường còn thiếu')
        $auditFlags=@()
        foreach($fn in $checkFields) {
            $field=$row.PSObject.Properties[$fn]
            if($null -eq $field){$auditFlags+=($fn+':MISSING_COLUMN')}
            elseif(-not ([string]$field.Value).Trim()){$auditFlags+=($fn+':EMPTY')}
        }
        if($auditFlags.Count -gt 0 -and $script:logFile) {
            Add-Content -LiteralPath $script:logFile -Encoding UTF8 -Value (
                '['+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')+'][DETAIL] Snapshot empty fields: '+($auditFlags -join ' | '))
        }
    }catch{}
    SurveyLog ('Office da ghi vao du lieu goc: '+$lic.OfficeState+' | Nguon: '+$lic.OfficeMethod) 'OK'
    SurveyLog ('Da luu du lieu goc: '+$snapshot) 'OK'

    SurveyLog 'BUOC 7/7 - Tong hop CSV va Excel' 'STEP'
    . (Join-Path $script:root 'ExportExcel.ps1')
    Write-Reports -rawFolder $script:rawFolder -outFolder $script:folder

    SurveyLog ('May: '+$env:COMPUTERNAME+' | Nguoi phu trach: '+$location[3]) 'OK'
    SurveyLog ('Hang/Model: '+$maker+' / '+$model+' | Serial: '+$serial)
    SurveyLog ('SMBIOS Asset Tag (Win32_SystemEnclosure.SMBIOSAssetTag): '+$asset)
    SurveyLog ('BIOS Version (Win32_BIOS.SMBIOSBIOSVersion): '+(Clean $bios.SMBIOSBIOSVersion))
    SurveyLog ('BaseBoard Product (Win32_BaseBoard.Product): '+(Clean $board.Product))
    if (-not (Real $board.Product)) { SurveyLog 'BaseBoard Product chua co gia tri hop le tu SMBIOS: can doi chieu tem/mainboard.' 'WARN' }
    SurveyLog ('Windows: '+(Clean $os.Caption)+' | Build: '+(Clean $os.BuildNumber))
    SurveyLog ('MAC LAN: '+(Distinct $lan)+' | MAC Wi-Fi: '+(Distinct $wifi))
    SurveyLog ('Admin Local: '+$security.Admin+' | TPM: '+$security.TPM+' | BitLocker: '+$security.BitLocker)
    SurveyLog ('BitLocker Conversion: '+$security.BitLockerConversion+' | Protection: '+$security.BitLockerProtection+' | Encrypted: '+$security.BitLockerPercent+'%')
    SurveyLog ('Antivirus: '+$security.Antivirus+' | Real-time: '+$security.AVRealTime+' | Cap nhat mau virus: '+$security.AVUpdated)
    SurveyLog ('Phan mem nghi van: '+$security.Suspicious)
    if($security.BitLockerDetail){SurveyLog ('BitLocker - Chan doan loi: '+$security.BitLockerDetail) 'WARN'}
    if($security.TPMDetail){SurveyLog ('TPM - Chan doan loi: '+$security.TPMDetail) 'WARN'}
    if($reviewReasons.Count -gt 0){SurveyLog ('Muc can doi chieu (chi ghi log): '+(Distinct $reviewReasons)) 'WARN'}
    SurveyLog ('AutoPlay UI HKCU: '+$compliance.AutoRunUserPreference+' | User: '+$compliance.AutoRunCheckedUser)
    SurveyLog ('SMBv1: '+$compliance.SMB1+' | Autorun: '+$compliance.AutoRun+' | NTP: '+$compliance.NTP)
    SurveyLog ('Firewall: '+$extra.Firewall+' | Guest: '+$extra.Guest)
    SurveyLog ('HotFix: '+$extra.Hotfix+' | DNS: '+$extra.DNS)
    SurveyLog ('USB Storage: '+$extra.USB+' | Khoa man hinh: '+$extra.ScreenLock)
    if($security.Findings) { SurveyLog ('Chi tiet dau hieu: '+$security.Findings) 'WARN' }

    if ($missing.Count -gt 0) { SurveyLog ('Thong tin can doi chieu: '+(Distinct $missing)) 'WARN' }
    if ($script:errs.Count -gt 0) {
        SurveyLog ('Co '+$script:errs.Count+' loi truy van. Chi tiet chi nam trong log:') 'WARN'
        foreach($oneError in $script:errs) { SurveyLog ('QUERY_ERROR: '+[string]$oneError) 'WARN' }
    }
    if ($script:lastReportPath) { SurveyLog ('Bao cao: '+$script:lastReportPath) 'OK' }
    else { SurveyLog 'Chua tao duoc XLSX, hay xem Inventory.csv.' 'WARN' }
    SurveyLog ('Nhat ky: '+$script:logFile) 'OK'
    SurveyLog '===== HOAN TAT =====' 'OK'
    Show-AuditFinish $script:lastReportPath $script:logFile $script:errs.Count
} catch {
    SurveyLog ('LOI: '+$_.Exception.Message) 'ERROR'
    SurveyLog 'Neu da co file DuLieuGoc, hay chay 2_Rebuild_Excel.bat de tong hop lai.' 'WARN'
    if ($script:logFile) { SurveyLog ('Nhat ky loi: '+$script:logFile) 'ERROR' }
    exit 1
}
