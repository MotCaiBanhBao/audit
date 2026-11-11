# ParallelAuditWorker.ps1 - Windows PowerShell 2.0+ - read-only, no Excel/USB writes.
# Separate process per task: License, Security, or Policies.
param(
    [Parameter(Mandatory=$true)][string]$Task,
    [Parameter(Mandatory=$true)][string]$JobDir,
    [Parameter(Mandatory=$true)][string]$Root
)
$ErrorActionPreference='Stop'
$env:AUDIT_RAW_OUTPUT_DIR=Join-Path $JobDir ('RAW_'+$Task)
. (Join-Path $Root 'RawEvidence.ps1')
if (@('WINDOWS','OFFICE','SECURITY','POLICIES') -notcontains $Task) {exit 3}
$logPath=Join-Path $JobDir ($Task+'.log')
$outputPath=Join-Path $JobDir ($Task+'.xml')
function WorkerLog([string]$message,[string]$level='INFO') {
    $line='['+(Get-Date -Format 'HH:mm:ss')+']['+$Task+']['+$level+'] '+($message -replace '[\r\n]+',' ')
    try { [System.IO.File]::AppendAllText($logPath,($line+[Environment]::NewLine),(New-Object System.Text.UTF8Encoding -ArgumentList $false)) } catch {}
}
function SurveyLog([string]$msg,[string]$level='INFO'){ WorkerLog $msg $level }
try {
    WorkerLog 'BAT_DAU'
    $data=$null
    if ($Task -eq 'WINDOWS') {
        . (Join-Path $Root 'LicenseCheck.ps1')
        WorkerLog 'Windows: bat dau kiem tra doc lap'
        $data=GetLicenseSummary -WindowsOnly
        WorkerLog ('Windows='+$data.WindowsState) 'OK'
    } elseif ($Task -eq 'OFFICE') {
        . (Join-Path $Root 'LicenseCheck.ps1')
        WorkerLog 'Office: bat dau kiem tra doc lap'
        $data=GetLicenseSummary -OfficeOnly
        WorkerLog ('Office='+$data.OfficeState) 'OK'
    } elseif ($Task -eq 'SECURITY') {
        . (Join-Path $Root 'CsvCompat.ps1')
        . (Join-Path $Root 'SecurityAudit.ps1')
        $audit=GetSecurityAudit -rulesPath (Join-Path $Root 'SuspiciousSoftwareRules.csv')
        # Apps live only in local TEMP. Parent will export them on USB once.
        $data=$audit
        $apps=@($script:SuspectInstalledApps | Where-Object {$_.Name})
        if($apps.Count -gt 0){
            Write-Utf8Csv -Records $apps -FilePath (Join-Path $JobDir 'SECURITY_APPS.csv')
        }
        WorkerLog ('Hoan tat SecurityAudit; app entries='+@($script:SuspectInstalledApps).Count) 'OK'
    } else {
        . (Join-Path $Root 'SecurityExtras.ps1')
        . (Join-Path $Root 'ComplianceExtras.ps1')
        WorkerLog 'Dang kiem tra Firewall/HotFix/USB/Guest/DNS/Screen Lock'
        $extras=GetExtraSecurityAudit
        WorkerLog 'Dang kiem tra Domain/SMB1/Autorun/NTP/DHCP'
        $cs=$null; $configs=@()
        try {$cs=Get-WmiObject Win32_ComputerSystem -ErrorAction Stop | Select-Object -First 1;Write-RawEvidence 'WMI_ComputerSystem_Policies' 'Get-WmiObject Win32_ComputerSystem' $cs}
        catch {WorkerLog ('WMI ComputerSystem: '+$_.Exception.Message) 'WARN'}
        try {$configs=@(Get-WmiObject Win32_NetworkAdapterConfiguration -ErrorAction Stop);Write-RawEvidence 'WMI_NetworkConfiguration_Policies' 'Get-WmiObject Win32_NetworkAdapterConfiguration' $configs}
        catch {WorkerLog ('WMI NetworkAdapterConfiguration: '+$_.Exception.Message) 'WARN'}
        $compliance=GetComplianceExtras -cs $cs -configs $configs
        # Keep separate flat objects to remain compatible with PS2 CliXml depth.
        $data=$extras
        $complianceTemp=Join-Path $JobDir 'POLICIES_COMPLIANCE.xml.tmp'
        $complianceFinal=Join-Path $JobDir 'POLICIES_COMPLIANCE.xml'
        $compliance | Export-Clixml -Path $complianceTemp -ErrorAction Stop
        [System.IO.File]::Move($complianceTemp,$complianceFinal)
        WorkerLog 'Hoan tat security policies/compliance' 'OK'
    }
    if ($null -eq $data) {throw 'Worker returned empty result'}
    Write-RawEvidence ($Task+'_FullResult') ('Worker '+$Task+' returned fields') $data
    # Atomic per-worker result: write .tmp first, then publish .xml when complete.
    $tempPath=$outputPath+'.tmp'
    $data | Export-Clixml -Path $tempPath -ErrorAction Stop
    [System.IO.File]::Move($tempPath,$outputPath)
    WorkerLog 'DA_GHI_KET_QUA_TAM - OK' 'OK'
    exit 0
} catch {
    WorkerLog ('LOI: '+$_.Exception.Message) 'ERROR'
    exit 1
}
