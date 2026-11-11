# Child process: bounded suspicious application/process/service name matching.
# Read-only; compatible with Windows PowerShell 2.0 (Windows 7+) and .NET 3.5+.
param(
    [string]$InputPath,
    [string]$RulesPath,
    [string]$OutputPath,
    [string]$StatusPath
)
$ErrorActionPreference='Stop'
. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'RawEvidence.ps1')
. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'CsvCompat.ps1')
function SetStage([string]$stage) {
    try { Set-Content -LiteralPath $StatusPath -Value $stage -Encoding ASCII -ErrorAction SilentlyContinue } catch {}
}
function Flat([object]$value) {
    if ($null -eq $value) { return '' }
    return ([string]$value).Trim() -replace '[\r\n\t]+',' '
}
try {
    SetStage 'LOAD_RULES'
    $rules=@()
    if(Test-Path -LiteralPath $RulesPath) {
        $rules=@(Read-Utf8Csv -FilePath $RulesPath)
    }
    # Gather installed apps HERE, too, so the parent watchdog also bounds
    # slow registry reads. Avoid Win32_Product/MSI repair side effects.
    SetStage 'ENUMERATE_INSTALLED_APPS'
    $apps=New-Object System.Collections.ArrayList
    $rawInstalled=New-Object System.Collections.ArrayList
    $roots=@(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    foreach($root in $roots) {
        if(-not (Test-Path $root)) { continue }
        SetStage ('INSTALLED_APPS: '+$root)
        try {
            foreach($key in @(Get-ChildItem -Path $root -ErrorAction Stop)) {
                try {
                    $detail=Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction Stop
                    [void]$rawInstalled.Add($detail)
                    if($detail.DisplayName -and ([string]$detail.DisplayName).Trim()) {
                        [void]$apps.Add((New-Object PSObject -Property @{
                            Name=(Flat $detail.DisplayName);Version=(Flat $detail.DisplayVersion);Publisher=(Flat $detail.Publisher)
                        }))
                    }
                } catch {}
            }
        } catch {}
    }
    Write-RawEvidence 'UninstallRegistry_Full' 'Get-ChildItem registry Uninstall; Get-ItemProperty each key' @($rawInstalled)
    SetStage ('WRITE_INSTALLED_APPS_'+$apps.Count)
    if($apps.Count -gt 0) {
        Write-Utf8Csv -Records @($apps | Sort-Object Name,Version -Unique) -FilePath $InputPath
    }
    $candidates=New-Object System.Collections.ArrayList
    foreach($app in $apps) {
        if($app.Name) {
            [void]$candidates.Add((New-Object PSObject -Property @{Source='InstalledApp';Value=(Flat $app.Name)}))
        }
    }
    SetStage 'ENUMERATE_PROCESSES'
    try {
        $processes=@(Get-Process -ErrorAction Stop)
        Write-RawEvidence 'Get_Process_Full' 'Get-Process' $processes
        foreach($p in $processes) {
            if($p.ProcessName) {
                [void]$candidates.Add((New-Object PSObject -Property @{Source='Process';Value=(Flat $p.ProcessName)}))
            }
        }
    } catch { SetStage ('PROCESS_WARNING: '+(Flat $_.Exception.Message)) }
    SetStage 'ENUMERATE_SERVICES'
    # Get-Service avoids the Win32_Service WMI provider, which can stall on some PCs.
    try {
        $allServices=@(Get-Service -ErrorAction Stop)
        Write-RawEvidence 'Get_Service_Full' 'Get-Service' $allServices
        foreach($s in $allServices) {
            [void]$candidates.Add((New-Object PSObject -Property @{
                Source='Service';Value=((Flat $s.ServiceName)+' '+(Flat $s.DisplayName)).Trim()
            }))
        }
    } catch { SetStage ('SERVICE_WARNING: '+(Flat $_.Exception.Message)) }
    SetStage ('MATCH_RULES_'+$rules.Count+'_CANDIDATES_'+$candidates.Count)
    $found=New-Object System.Collections.ArrayList
    foreach($rule in $rules) {
        $pattern=[string]$rule.Pattern
        if(-not $pattern) { continue }
        try {
            $rx=New-Object System.Text.RegularExpressions.Regex -ArgumentList @(
                $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
            )
        } catch { continue }
        foreach($candidate in $candidates) {
            if($rx.IsMatch([string]$candidate.Value)) {
                [void]$found.Add((New-Object PSObject -Property @{
                    Category=(Flat $rule.Category);Label=(Flat $rule.Label);
                    Evidence=(Flat $candidate.Value);Source=[string]$candidate.Source
                }))
            }
        }
    }
    SetStage 'WRITE_RESULTS'
    if($found.Count -gt 0) {
        Write-Utf8Csv -Records @($found | Sort-Object Category,Label,Evidence,Source -Unique) -FilePath $OutputPath
    }
    SetStage 'DONE'
    exit 0
} catch {
    SetStage ('ERROR: '+(Flat $_.Exception.Message))
    exit 1
}
