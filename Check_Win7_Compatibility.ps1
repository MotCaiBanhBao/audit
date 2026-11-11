# Read-only compatibility preflight for Device Audit Tool.
# Uses Windows PowerShell 2.0 syntax; DOES NOT run the survey.
$ErrorActionPreference='Continue'
$root=Split-Path -Parent $MyInvocation.MyCommand.Path
$fail=0
$warn=0
function Result([string]$name,[string]$state,[string]$detail) {
    $color='Green'
    if($state -eq 'FAIL') {$script:fail++; $color='Red'}
    if($state -eq 'WARN') {$script:warn++; $color='Yellow'}
    Write-Host ('['+$state+'] '+$name+': '+$detail) -ForegroundColor $color
}
Write-Host '===============================================================' -ForegroundColor Cyan
Write-Host ' DEVICE AUDIT TOOL - WINDOWS 7 / PS2 COMPATIBILITY CHECK' -ForegroundColor Cyan
Write-Host ' Khong thuc hien khao sat, khong thay doi cau hinh he thong.'
Write-Host '==============================================================='
try {
    $v=$PSVersionTable.PSVersion
    if($v -and $v.Major -ge 2) {Result 'PowerShell' 'PASS' ([string]$v)}
    else {Result 'PowerShell' 'FAIL' 'Can Windows PowerShell 2.0 tro len'}
} catch {Result 'PowerShell' 'FAIL' ([string]$_.Exception.Message)}
try {
    $os=Get-WmiObject -Class Win32_OperatingSystem -ErrorAction Stop | Select-Object -First 1
    $ver=[string]$os.Version
    if($ver -match '^6\.1\.') {Result 'Windows' 'PASS' ('Windows 7: '+$ver+' | '+$os.OSArchitecture)}
    else {Result 'Windows' 'WARN' ('Khong phai Windows 7: '+$ver+' | '+$os.Caption)}
    Result 'WMI Win32_OperatingSystem' 'PASS' 'Da truy van thanh cong'
} catch {Result 'WMI Win32_OperatingSystem' 'FAIL' ([string]$_.Exception.Message)}
try {
    $identity=[System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal=New-Object System.Security.Principal.WindowsPrincipal($identity)
    if($principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Result 'Administrator' 'PASS' 'Da chay nang quyen'
    } else {Result 'Administrator' 'FAIL' 'Can Run as administrator de chay khao sat'}
} catch {Result 'Administrator' 'FAIL' ([string]$_.Exception.Message)}
foreach($cmd in @('Get-WmiObject','Import-Clixml','Export-Clixml','ConvertTo-Csv','ConvertFrom-Csv','Add-Type')) {
    try {
        if(Get-Command $cmd -ErrorAction SilentlyContinue){Result $cmd 'PASS' 'Co san'}
        else {Result $cmd 'FAIL' 'Khong co lenh'}
    } catch {Result $cmd 'FAIL' ([string]$_.Exception.Message)}
}
try {
    Add-Type -AssemblyName WindowsBase -ErrorAction Stop | Out-Null
    $pkg=[type]'System.IO.Packaging.Package'
    if($pkg){Result 'Excel XLSX engine' 'PASS' 'WindowsBase / System.IO.Packaging OK'}
    else {Result 'Excel XLSX engine' 'FAIL' 'Khong nap duoc System.IO.Packaging.Package'}
} catch {Result 'Excel XLSX engine' 'FAIL' ([string]$_.Exception.Message)}
$required=@('Survey.ps1','LicenseCheck.ps1','ParallelAudit.ps1','ParallelAuditWorker.ps1',
    'ExportExcel.ps1','ManualEdits.ps1','SecurityAudit.ps1','SecurityExtras.ps1',
    'ComplianceExtras.ps1','CsvCompat.ps1','SuspiciousScanWorker.ps1',
    'ConsoleFont.ps1','ReviewReasons.ps1','RawEvidence.ps1','WindowsLicenseFallback.ps1',
    '1_Run_Survey.bat','2_Rebuild_Excel.bat')
foreach($f in $required) {
    if(-not (Test-Path -LiteralPath (Join-Path $root $f) -PathType Leaf)) {
        Result 'Required file' 'FAIL' ('Thieu '+$f)
    }
}
# PSParser is available in Windows PowerShell 2.0; it checks syntax without
# running any script. It does not prove that all runtime calls will succeed.
try {
    $files=@(Get-ChildItem -Path $root -Filter '*.ps1' -ErrorAction Stop)
    $parseFailures=0
    foreach($f in $files) {
        try {
            $source=[System.IO.File]::ReadAllText($f.FullName,[System.Text.Encoding]::UTF8)
            $parseErrors=$null
            $tokens=[System.Management.Automation.PSParser]::Tokenize($source,[ref]$parseErrors)
            if($parseErrors -and @($parseErrors).Count -gt 0) {
                $parseFailures++
                $first=$parseErrors | Select-Object -First 1
                Result ('Parse '+$f.Name) 'FAIL' ([string]$first.Message)
            }
        } catch {
            $parseFailures++
            Result ('Parse '+$f.Name) 'FAIL' ([string]$_.Exception.Message)
        }
    }
    if($parseFailures -eq 0) {
        Result 'PowerShell syntax' 'PASS' ('Da parse '+$files.Count+' file .ps1 bang engine hien tai')
    }
} catch {Result 'PowerShell syntax' 'FAIL' ([string]$_.Exception.Message)}
Write-Host '---------------------------------------------------------------'
if($fail -gt 0) {
    Write-Host ('KET QUA: CHUA DU DIEU KIEN CHAY ('+$fail+' loi, '+$warn+' canh bao)') -ForegroundColor Red
    exit 1
}
if($warn -gt 0) {
    Write-Host ('KET QUA: KIEM TRA CO DIEU KIEN ('+$warn+' canh bao). Can test khao sat thuc te.') -ForegroundColor Yellow
    exit 0
}
Write-Host 'KET QUA: DIEU KIEN CO BAN DAT. Van can test mot luot khao sat.' -ForegroundColor Green
exit 0
