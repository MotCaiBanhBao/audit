# License inventory: READ-ONLY. Works in Windows PowerShell 2.0-5.1.
# No full product keys, email addresses, tenant IDs, or token files are collected.
function LicenseTrace([string]$m) {
    if(Get-Command SurveyLog -ErrorAction SilentlyContinue){SurveyLog ('License: '+$m) 'INFO'}
}

# Normalize only explicit activation results. Grace/unknown/missing is NOT
# equivalent to an unlicensed product. Numeric states follow Microsoft SPP:
# 0=Unlicensed, 1=Licensed, 5=Notification.
function Convert-OfficeLicenseState([string]$value) {
    if (-not $value) {return 'CAN_XAC_MINH'}
    $state=([string]$value).Trim()
    if ($state -eq '1') {return 'DA_KICH_HOAT'}
    if ($state -eq '0' -or $state -eq '5') {return 'CHUA_KICH_HOAT'}
    if ($state -eq '2' -or $state -eq '3' -or $state -eq '6') {return 'THOI_GIAN_AN_HAN'}
    if ($state -eq '4') {return 'CAN_RA_SOAT_OFFICE_LICENSE'}
    if ($state -match '(?i)UNLICENSED|NOT.LICENSED|NOTIFICATIONS?|EXPIRED|NO.LICENSE') {return 'CHUA_KICH_HOAT'}
    if ($state -match '(?i)NON.GENUINE.GRACE') {return 'CAN_RA_SOAT_OFFICE_LICENSE'}
    if ($state -match '(?i)GRACE') {return 'THOI_GIAN_AN_HAN'}
    if ($state -match '(?i)(^|[^A-Z])LICENSED([^A-Z]|$)') {return 'DA_KICH_HOAT'}
    return 'CAN_XAC_MINH'
}
function Get-OfficeLicenseDecision([object[]]$licenseRows) {
    if (-not $licenseRows -or $licenseRows.Count -eq 0) {return 'CAN_XAC_MINH'}
    $states=@()
    foreach ($entry in $licenseRows) { $states += (Convert-OfficeLicenseState ([string]$entry.State)) }
    if ($states -contains 'DA_KICH_HOAT') {return 'DA_KICH_HOAT'}
    # Do not automatically mark unlicensed if another candidate is in a grace
    # or unrecognized status; additional verification is needed.
    if ($states -contains 'CAN_XAC_MINH') {return 'CAN_XAC_MINH'}
    if ($states -contains 'CAN_RA_SOAT_OFFICE_LICENSE') {return 'CAN_RA_SOAT_OFFICE_LICENSE'}
    if ($states -contains 'THOI_GIAN_AN_HAN') {return 'THOI_GIAN_AN_HAN'}
    if ($states -contains 'CHUA_KICH_HOAT') {return 'CHUA_KICH_HOAT'}
    return 'CAN_XAC_MINH'
}
# Click-to-Run ProductReleaseIds identifies the installed SKU (not a license).
# The legacy SoftwareLicensingProduct store can include dozens of unrelated
# subscription/trial SKUs with LicenseStatus=0. Never use all such rows as a
# vote that the installed Office is unlicensed.
function Get-OfficeInstalledReleaseIds {
    $ids=@()
    $inventoryKeys=@(
        'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Inventory\Office\16.0'
    )
    foreach ($path in $inventoryKeys) {
        try {
            $v=Get-ItemProperty -Path $path -ErrorAction Stop
            Write-RawEvidence 'Office_ReleaseRegistry' ('Get-ItemProperty -Path '+$path) $v
            foreach ($raw in @($v.OfficeProductReleaseIds)) {
                foreach ($part in (([string]$raw) -split '[,;\s]+')) {
                    if ($part -and $part -ne 'System.Object[]') {$ids+= $part.Trim()}
                }
            }
        } catch { }
    }
    # The Inventory key is preferred. Configuration is only a fallback.
    if ($ids.Count -eq 0) {
        foreach ($path in @(
            'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration'
        )) {
            try {
                $v=Get-ItemProperty -Path $path -ErrorAction Stop
                Write-RawEvidence 'Office_ReleaseRegistry' ('Get-ItemProperty -Path '+$path) $v
                foreach ($raw in @($v.ProductReleaseIds)) {
                    foreach ($part in (([string]$raw) -split '[,;\s]+')) {
                        if ($part -and $part -ne 'System.Object[]') {$ids+= $part.Trim()}
                    }
                }
            } catch { }
        }
    }
    return @($ids | Where-Object { $_ } | Select-Object -Unique)
}
function Convert-OfficeReleaseToLegacyToken([string]$releaseId) {
    $releaseId=([string]$releaseId).Trim()
    if (-not $releaseId) {return ''}
    # Examples: ProPlus2019Volume -> ProPlus2019VL,
    # ProfessionalRetail -> ProfessionalR, HomeBusiness2021Retail -> HomeBusiness2021R.
    $token=[regex]::Replace($releaseId,'(?i)Retail$','R')
    $token=[regex]::Replace($token,'(?i)Volume$','VL')
    return $token
}
function Get-OfficeLegacyProductToken([string]$productName) {
    $m=[regex]::Match($productName,'(?i)\bOffice\s*\d+\s*,\s*Office\d+([a-z0-9]+?)(?:_|\s+edition\b|$)')
    if ($m.Success) {return $m.Groups[1].Value}
    return ''
}
function Test-OfficeLegacyProductMatch([string]$productName,[string[]]$releaseIds) {
    if (-not $releaseIds -or $releaseIds.Count -eq 0) {return $false}
    $token=Get-OfficeLegacyProductToken $productName
    if (-not $token) {return $false}
    foreach ($release in $releaseIds) {
        $target=Convert-OfficeReleaseToLegacyToken $release
        if ($target -and $token -ieq $target) {return $true}
        # Some volume licenses use ...VL instead of ...Volume.
        if ($target -and $token -ieq $release) {return $true}
    }
    return $false
}
function Test-OfficePlaceholderLicense([string]$name,[string]$description) {
    # Office may register trial, grace and subscription templates that are not
    # the actual product entitlement. They are not negative license evidence.
    if ($name -match '(?i)_(?:Subscription|SubTest|SubTrial|Sub\d+|Grace|Trial|Demo|BypassTrial)') {return $true}
    if ($description -match '(?i)TIMEBASED_SUB|RETAIL\s*\((?:Trial|Free|Grace)\)') {return $true}
    return $false
}

# vNext is per Windows/Office user, not a device-wide proof of entitlement.
# Compare the audit process SID to the interactive logon user before using
# an empty local vNext result to assess current-user activation. If they do
# not match, leave license status unresolved and ask the operator.
function Get-OfficeAccountContext {
    $value='UNKNOWN'
    try {
        $identity=[System.Security.Principal.WindowsIdentity]::GetCurrent()
        $currentSid=''
        if ($identity -and $identity.User) {$currentSid=[string]$identity.User.Value}
        $cs=Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop | Select-Object -First 1
        $interactive=[string]$cs.UserName
        if ($currentSid -and $interactive) {
            $interactiveSid=''
            try {
                $account=New-Object System.Security.Principal.NTAccount -ArgumentList $interactive
                $interactiveSid=[string]$account.Translate([System.Security.Principal.SecurityIdentifier]).Value
            } catch { }
            if ($interactiveSid) {
                if ($interactiveSid -eq $currentSid) {return 'MATCH'}
                return 'OTHER'
            }
            # Some disconnected domain accounts cannot be translated. A
            # matching qualified name is still good evidence of same user.
            if ($identity.Name -and $identity.Name -ieq $interactive) {return 'MATCH'}
            return 'UNKNOWN'
        }
    } catch { }
    return $value
}
function Test-OfficeVNextEmpty([string]$raw) {
    # Only a successful vnextdiag listing with BOTH vNext and device sections
    # explicitly empty is conclusive about no local vNext tokens in this context.
    # No output, script failures, or blocked execution must NOT return true.
    if (-not $raw) {return $false}
    $perUser=[regex]::IsMatch($raw,'(?is)vNext\s+licenses\s+found\s*=+\s*No licenses found\.')
    $device=[regex]::IsMatch($raw,'(?is)Device\s+licenses\s+found\s*=+\s*No licenses found\.')
    return ($perUser -and $device)
}

# Bounded invocation of vendor diagnostics.  NEVER let ospp/vNext block the survey.
# Read-only, compatible with Windows PowerShell 2.0 / .NET 3.5+.
function Invoke-OfficeDiagnosticBounded([string]$exe,[string]$arguments,[int]$maxSeconds,[string]$label) {
    $p=$null
    try {
        $p=New-Object System.Diagnostics.Process
        $si=New-Object System.Diagnostics.ProcessStartInfo
        $si.FileName=$exe
        $si.Arguments=$arguments
        $si.UseShellExecute=$false
        $si.CreateNoWindow=$true
        $si.RedirectStandardOutput=$true
        $si.RedirectStandardError=$true
        $p.StartInfo=$si
        if (-not $p.Start()) { LicenseTrace ($label+': khoi dong that bai'); return $null }
        # If the command blocks or fills output buffers, time out safely.
        if (-not $p.WaitForExit($maxSeconds*1000)) {
            LicenseTrace ($label+': TIMEOUT '+$maxSeconds+'s; bo qua nguon nay')
            Write-RawCommandText ('Office_'+$label) ($exe+' '+$arguments) '' '' ('TIMEOUT_'+$maxSeconds+'s')
            try { $p.Kill() } catch { }
            try { [void]$p.WaitForExit(1500) } catch { }
            return $null
        }
        $stdout=$p.StandardOutput.ReadToEnd()
        $stderr=$p.StandardError.ReadToEnd()
        if ($p.ExitCode -ne 0) {
            LicenseTrace ($label+': exit code '+$p.ExitCode+'; '+($stderr -replace '[\r\n]+',' '))
            Write-RawCommandText ('Office_'+$label) ($exe+' '+$arguments) ([string]$stdout) ([string]$stderr) ('EXIT_'+$p.ExitCode)
            return $null
        }
        return [string]$stdout
    } catch {
        LicenseTrace ($label+': loi thuc thi: '+$_.Exception.Message)
        Write-RawCommandText ('Office_'+$label) ($exe+' '+$arguments) '' ([string]$_.Exception.Message) 'ERROR'
        return $null
    } finally {
        if ($p) {
            try {if (-not $p.HasExited) {$p.Kill()}}catch{}
            try {$p.Dispose()}catch{}
        }
    }
}

function GetLicenseSummary([switch]$WindowsOnly,[switch]$OfficeOnly) {
    LicenseTrace 'Bat dau truy van Windows activation qua WMI' 
    $r = New-Object PSObject
    $r | Add-Member NoteProperty WindowsState 'CAN_XAC_MINH'
    $r | Add-Member NoteProperty WindowsEdition ''
    $r | Add-Member NoteProperty WindowsChannel ''
    $r | Add-Member NoteProperty WindowsMethod 'SoftwareLicensingProduct (WMI)'
    $r | Add-Member NoteProperty WindowsNote ''
    $r | Add-Member NoteProperty WordInstalled 'KHONG_PHAT_HIEN'
    $r | Add-Member NoteProperty ExcelInstalled 'KHONG_PHAT_HIEN'
    $r | Add-Member NoteProperty WordState 'CAN_XAC_MINH'
    $r | Add-Member NoteProperty ExcelState 'CAN_XAC_MINH'
    $r | Add-Member NoteProperty OfficeInstalled 'CHUA_XAC_MINH'
    $r | Add-Member NoteProperty OfficeState 'CAN_XAC_MINH'
    $r | Add-Member NoteProperty OfficeProduct ''
    $r | Add-Member NoteProperty OfficeMethod ''
    $r | Add-Member NoteProperty OfficeNote ''

    if(-not $OfficeOnly) {
    # Windows only (not Office). Determine the CURRENT installed edition before
    # interpreting LicenseStatus. A zero/notification status on an unrelated
    # edition must not be used to assert the current OS is unlicensed.
    # PowerShell 2.0 compatible. Read-only; never activates/modifies licenses.
    try {
        $allWins = @(Get-WmiObject -Class SoftwareLicensingProduct -Filter "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f'" -ErrorAction Stop)
        Write-RawEvidence 'Windows_SoftwareLicensingProduct' 'Get-WmiObject SoftwareLicensingProduct -Filter Windows ApplicationID' $allWins
        $wins = @($allWins | Where-Object {
            $_.PartialProductKey -and $_.Name -match '(?i)Windows' -and $_.LicenseIsAddon -ne $true
        })
        $editionID = ''
        try {
            $osReg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
            Write-RawEvidence 'Registry_WindowsCurrentVersion' 'Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' $osReg
            if ($osReg.EditionID) { $editionID = ([string]$osReg.EditionID).Trim() }
        } catch { }
        $selected = @()
        $selectionReason = ''
        if ($editionID) {
            # Example: Windows(R), Professional edition -> EditionID Professional.
            # Match exact edition name to avoid Enterprise/EnterpriseS ambiguity.
            $editionRx = '(?i),\s*'+[regex]::Escape($editionID)+'\s+edition\b'
            $selected = @($wins | Where-Object { ([string]$_.Name) -match $editionRx })
            if ($selected.Count -gt 0) { $selectionReason = 'EditionID='+$editionID }
        }
        if ($selected.Count -eq 0 -and $wins.Count -eq 1) {
            $selected = @($wins)
            $selectionReason = 'OnlyOneInstalledWindowsKey'
        }
        if ($selected.Count -eq 0) {
            # No reliable edition match. A licensed row may still identify the
            # active Windows product. Otherwise, never guess from other editions.
            $licensedAll = @($wins | Where-Object { $_.LicenseStatus -eq 1 })
            if ($licensedAll.Count -eq 1) {
                $selected = @($licensedAll)
                $selectionReason = 'SingleLicensedWindowsKey'
            }
        }
        $activated = @($selected | Where-Object { $_.LicenseStatus -eq 1 })
        if ($activated.Count -gt 0) {
            $chosen = $activated[0]
            $r.WindowsState = 'DA_KICH_HOAT'
            $r.WindowsNote = 'Windows LicenseStatus=1 (Licensed); '+$selectionReason+'. Trang thai kich hoat khong chung minh tinh hop le cua giay phep.'
        } elseif ($selected.Count -gt 0) {
            $chosen = $selected[0]
            $codes = @($selected | ForEach-Object { [int]$_.LicenseStatus } | Select-Object -Unique)
            $codeText = @($codes | ForEach-Object { [string]$_ }) -join ','
            $allUnlicensed = @($selected | Where-Object { $_.LicenseStatus -ne 0 -and $_.LicenseStatus -ne 5 }).Count -eq 0
            if ($allUnlicensed) {
                # 0=Unlicensed, 5=Notification (Microsoft WMI). With a matching
                # installed edition/key, these directly indicate not activated.
                $r.WindowsState = 'CHUA_KICH_HOAT'
                $r.WindowsNote = 'Windows LicenseStatus='+$codeText+' (0=Unlicensed, 5=Notification); '+$selectionReason+'.'
            } else {
                # 2,3,4,6=grace / non-genuine grace; do not silently label as
                # permanent activation or conflate with plain unlicensed.
                $r.WindowsState = 'CAN_XAC_MINH_WINDOWS_LICENSE'
                $r.WindowsNote = 'Windows LicenseStatus='+$codeText+' (co the dang grace / trang thai bat thuong); '+$selectionReason+'; doi chieu slmgr /dli va /xpr.'
            }
        } else {
            $r.WindowsState = 'CAN_XAC_MINH_WINDOWS_LICENSE'
            if ($wins.Count -gt 0) {
                $states = @($wins | ForEach-Object { [string]$_.LicenseStatus } | Select-Object -Unique) -join ','
                $r.WindowsNote = 'Co '+$wins.Count+' ban ghi Windows co product key (LicenseStatus='+$states+'), nhung khong chon duoc edition hien tai (EditionID='+$editionID+'). Chay slmgr /dli de doi chieu.'
            } else {
                $r.WindowsState = 'CAN_XAC_MINH_KHONG_CO_BAN_GHI_WINDOWS'
                $r.WindowsNote = 'WMI khong thay ban ghi Windows co PartialProductKey, khong du can cu ket luan chua kich hoat.'
            }
        }
        if ($chosen) {
            $r.WindowsEdition = [string]$chosen.Name
            $r.WindowsChannel = [string]$chosen.Description
        }
    } catch {
        Write-RawEvidence 'Windows_SoftwareLicensingProduct' 'Get-WmiObject SoftwareLicensingProduct -Filter Windows ApplicationID' $null ([string]$_.Exception.Message)
        $r.WindowsState = 'CAN_XAC_MINH_WMI_LICENSE_ERROR'
        $r.WindowsNote = [string]$_.Exception.Message
    }

    LicenseTrace ('Windows xong: '+$r.WindowsState)
    } # Windows phase (independent of Office)
    if($WindowsOnly){return $r}
    # Phase A: verify an actual Office executable. Old App Paths registry entries
    # are NOT sufficient: their default executable must still exist on disk.
    # If no Office application is found, do not ask for manual confirmation.
    $bases = @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramW6432)
    $officeDirs = @()
    foreach ($base in $bases) {
        if (-not $base) { continue }
        foreach ($version in @('Office16','Office15','Office14','Office12')) {
            foreach ($rel in @('Microsoft Office\root','Microsoft Office')) {
                $candidate = Join-Path (Join-Path $base $rel) $version
                if (Test-Path -LiteralPath $candidate -PathType Container) { $officeDirs += $candidate }
            }
        }
    }
    $officeDirs = @($officeDirs | Select-Object -Unique)
    $applicationFiles = @('WINWORD.EXE','EXCEL.EXE','POWERPNT.EXE','OUTLOOK.EXE','MSACCESS.EXE')
    $installedApps = @()
    foreach ($d in $officeDirs) {
        foreach ($app in $applicationFiles) {
            if (Test-Path -LiteralPath (Join-Path $d $app) -PathType Leaf) {
                $installedApps += $app
            }
        }
    }
    foreach ($app in $applicationFiles) {
        foreach ($regRoot in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths',
                               'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\App Paths')) {
            $reg = Join-Path $regRoot $app
            try {
                $key = Get-Item -LiteralPath $reg -ErrorAction Stop
                Write-RawEvidence 'Office_AppPathRegistry' ('Get-Item -LiteralPath '+$reg) $key
                $exePath = [string]$key.GetValue('')
                if (-not $exePath) {continue}
                $exePath = [Environment]::ExpandEnvironmentVariables($exePath.Trim().Trim('"'))
                if (Test-Path -LiteralPath $exePath -PathType Leaf) {
                    $installedApps += $app
                    $officeDirs += (Split-Path $exePath -Parent)
                }
            } catch { }
        }
    }
    $officeDirs = @($officeDirs | Where-Object { $_ } | Select-Object -Unique)
    $installedApps = @($installedApps | Select-Object -Unique)
    if ($installedApps -contains 'WINWORD.EXE') {$r.WordInstalled='CO'} else {$r.WordState='KHONG_PHAT_HIEN'}
    if ($installedApps -contains 'EXCEL.EXE') {$r.ExcelInstalled='CO'} else {$r.ExcelState='KHONG_PHAT_HIEN'}
    if ($installedApps.Count -eq 0) {
        $r.OfficeInstalled='KHONG_PHAT_HIEN'
        $r.OfficeState='KHONG_PHAT_HIEN'
        $r.WordState='KHONG_PHAT_HIEN'
        $r.ExcelState='KHONG_PHAT_HIEN'
        $r.OfficeMethod='SCAN_UNG_DUNG'
        $r.OfficeNote='Khong tim thay chuong trinh Office tai cac duong dan/App Paths da kiem tra; khong ket luan tuyet doi chua cai.'
        LicenseTrace 'Khong phat hien Office; bo qua cau hoi kich hoat.'
        return $r
    }
    $r.OfficeInstalled='CO'
    LicenseTrace ('Office da cai: '+($installedApps -join ', '))

    # Detect the installed suite generation to exclude stale licensing records
    # from unrelated Office versions (e.g. an old Office 2013 key on Office 2019).
    $installedMajors = @()
    foreach ($d in $officeDirs) {
        if ($d -match '(?i)Office(12|14|15|16)(?:\\|$)') {
            $installedMajors += $matches[1]
        }
    }
    $installedMajors = @($installedMajors | Select-Object -Unique)
    # ProductReleaseIds is the installed SKU, unlike WMI which also contains
    # license templates for products not installed on this device.
    $officeReleaseIds=@(Get-OfficeInstalledReleaseIds)
    $m365Install = $false
    foreach ($release in $officeReleaseIds) {
        if ($release -match '(?i)O365|M365|Microsoft365') {$m365Install=$true;break}
    }
    if ($officeReleaseIds.Count -gt 0) {
        LicenseTrace ('Office SKU da cai: '+($officeReleaseIds -join ', '))
    } else {
        LicenseTrace 'Khong lay duoc ClickToRun ProductReleaseIds; WMI chua du xac nhan SKU dang cai.'
    }

    # Phase B: query WMI for traditional Office, but only score records that
    # match the installed SKU; never treat generic WMI subscription templates
    # as proof that a Microsoft 365 subscription is inactive.
    $wmiRows=@(); $osppRows=@(); $vnextRows=@()
    $vnextEmpty=$false
    $vnextExecuted=$false
    $accountContext='UNKNOWN'
    if ($m365Install) {$accountContext=Get-OfficeAccountContext}
    LicenseTrace ('Office ngu canh tai khoan: '+$accountContext+' (khong luu ten tai khoan)')
    if ($m365Install) {
        LicenseTrace 'Phat hien SKU Microsoft 365: khong dung LicenseStatus=0 cua cac mau WMI TIMEBASED_SUB de ket luan Office chua kich hoat.'
    }
    if (-not $m365Install) {
        try {
            LicenseTrace 'Kiem tra license Office qua WMI, doi chieu SKU da cai'
            $wmi=@(Get-WmiObject -Class SoftwareLicensingProduct -Filter "ApplicationID='0ff1ce15-a989-479d-af46-f275c6370663'" -ErrorAction Stop)
            Write-RawEvidence 'Office_SoftwareLicensingProduct' 'Get-WmiObject SoftwareLicensingProduct -Filter Office ApplicationID' $wmi
            foreach ($item in $wmi) {
                $name=([string]$item.Name).Trim()
                $description=([string]$item.Description).Trim()
                if ($name -notmatch '(?i)^Office\s+\d+\b' -or $name -match '(?i)Visio|Project') {continue}
                $versionOk=$installedMajors.Count -eq 0
                foreach ($major in $installedMajors) {
                    if ($name -match ('(?i)^Office\s*'+[regex]::Escape($major)+'\b')) {$versionOk=$true;break}
                }
                if (-not $versionOk -or $null -eq $item.LicenseStatus) {continue}
                if ($item.LicenseIsAddon -eq $true) {continue}
                # Known installed SKU: require a product-token match, not just
                # "Office 16". Without a known SKU, a licensed key is useful
                # evidence but a sea of status=0 WMI templates is NOT.
                if ($officeReleaseIds.Count -gt 0) {
                    if (-not (Test-OfficeLegacyProductMatch $name $officeReleaseIds)) {continue}
                } elseif (-not $item.PartialProductKey -or [int]$item.LicenseStatus -ne 1) {
                    continue
                }
                # Trial/subscription placeholder records don't prove that the
                # currently installed Office is unlicensed.
                if (([int]$item.LicenseStatus -ne 1) -and
                    (Test-OfficePlaceholderLicense $name $description)) {continue}
                $wmiRows += (New-Object PSObject -Property @{
                    Product=$name;State=([string]$item.LicenseStatus);
                    Method='WMI_MATCHED_SKU';Description=$description
                })
            }
        } catch { LicenseTrace ('Khong doc duoc WMI Office license: '+[string]$_.Exception.Message) }
        if ($wmiRows.Count -eq 0 -and ($installedMajors -contains '14')) {
            # Office 2010 can use a dedicated WMI provider on older Windows.
            try {
                $oldWmi=@(Get-WmiObject -Class OfficeSoftwareProtectionProduct -ErrorAction Stop)
                Write-RawEvidence 'OfficeSoftwareProtectionProduct' 'Get-WmiObject OfficeSoftwareProtectionProduct' $oldWmi
                foreach ($item in $oldWmi) {
                    $name=[string]$item.Name
                    if ($name -match '(?i)Office|Word|Excel' -and $name -notmatch '(?i)Visio|Project' -and $null -ne $item.LicenseStatus) {
                        $wmiRows += (New-Object PSObject -Property @{
                            Product=$name;State=([string]$item.LicenseStatus);
                            Method='Office WMI';Description=([string]$item.Description)
                        })
                    }
                }
            } catch { }
        }
    }
    $wmiDecision=Get-OfficeLicenseDecision $wmiRows
    if ($wmiRows.Count -gt 0) {
        LicenseTrace ('Office WMI: '+$wmiRows.Count+' ban ghi; '+$wmiDecision)
    }

    # OSPP is a second source for traditional/volume licensing; the actual
    # Office licensing scripts can be absent in a valid installation.
    if (-not $m365Install -and $wmiDecision -eq 'CAN_XAC_MINH') {
        foreach ($d in $officeDirs) {
            $ospp=Join-Path $d 'ospp.vbs'
            if (-not (Test-Path -LiteralPath $ospp -PathType Leaf)) {continue}
            try {
                LicenseTrace ('Kiem tra Office ospp.vbs: '+$d)
                $out=Invoke-OfficeDiagnosticBounded 'cscript.exe' ('//Nologo "'+$ospp+'" /dstatusall') 8 'OSPP'
                if ($null -eq $out) {continue}
                Write-RawCommandText 'Office_OSPP' ('cscript //Nologo '+$ospp+' /dstatusall') ([string]$out) '' 'OK'
                $hits=[regex]::Matches($out,'(?is)LICENSE NAME\s*:\s*([^\r\n]+)[\s\S]*?LICENSE STATUS\s*:\s*([^\r\n]+)')
                foreach ($hit in $hits) {
                    $osppRows += (New-Object PSObject -Property @{
                        Product=$hit.Groups[1].Value.Trim();
                        State=$hit.Groups[2].Value.Trim();Method='ospp'
                    })
                }
            } catch { LicenseTrace 'Khong doc duoc ket qua ospp.vbs.' }
        }
    }

    # The vNext licensing model is tied to the signed-in Office user. Run only
    # for Microsoft 365 or when old sources cannot determine activation.
    if ($m365Install -or ($wmiDecision -eq 'CAN_XAC_MINH' -and (Get-OfficeLicenseDecision $osppRows) -eq 'CAN_XAC_MINH')) {
        foreach ($d in $officeDirs) {
            $diag=Join-Path $d 'vnextdiag.ps1'
            if (-not (Test-Path -LiteralPath $diag -PathType Leaf)) {continue}
            try {
                LicenseTrace ('Kiem tra Office vnextdiag: '+$d)
                # Device Audit Tool BAT launches PowerShell with a process-only
                # execution policy. Enterprise Group Policy remains authoritative.
                $psPath=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
                $raw=Invoke-OfficeDiagnosticBounded $psPath ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "'+$diag+'" -action list') 10 'VNEXT'
                if ($null -eq $raw) {continue}
                Write-RawCommandText 'Office_VNEXT' ($psPath+' -File '+$diag+' -action list') ([string]$raw) '' 'OK'
                $vnextExecuted=$true
                if (Test-OfficeVNextEmpty $raw) {
                    $vnextEmpty=$true
                    LicenseTrace ('vNext: khong co token license trong ngu canh dang quet; tai khoan '+$accountContext)
                }
                $blocks=[regex]::Matches($raw,'(?s)\{[^{}]*\}')
                foreach ($block in $blocks) {
                    $p=[regex]::Match($block.Value,'(?i)"(?:Product|ProductName|Name)"\s*:\s*"([^"]+)"')
                    $st=[regex]::Match($block.Value,'(?i)"(?:LicenseState|License State|State)"\s*:\s*"([^"]+)"')
                    if ($p.Success -and $st.Success) {
                        $vnextRows += (New-Object PSObject -Property @{
                            Product=$p.Groups[1].Value;State=$st.Groups[1].Value;Method='vnextdiag'
                        })
                    }
                }
                if ($vnextRows.Count -eq 0) {
                    $sections=[regex]::Matches($raw,'(?ims)^\s*(?:Product Name|Product)\s*:\s*([^\r\n]+)(.*?)(?=^\s*(?:Product Name|Product)\s*:|\z)')
                    foreach ($section in $sections) {
                        $st=[regex]::Match($section.Groups[2].Value,'(?im)^\s*(?:License\s*State|LicenseState|State)\s*:\s*([^\r\n]+)')
                        if ($st.Success) {
                            $vnextRows += (New-Object PSObject -Property @{
                                Product=$section.Groups[1].Value.Trim();State=$st.Groups[1].Value.Trim();Method='vnextdiag'
                            })
                        }
                    }
                }
            } catch { LicenseTrace 'Khong chay/doc duoc vnextdiag (co the do Execution Policy).' }
        }
    }

    # Exclude Visio/Project: their activation does not activate Word/Excel.
    $nameRx='(?i)Office|O365|M365|Microsoft365|Word|Excel|ProPlus|HomeBusiness|HomeStudent|Standard|Professional|Mondo|Personal|Business'
    $vnextRows=@($vnextRows | Where-Object { ([string]$_.Product) -notmatch '(?i)Visio|Project' -and ([string]$_.Product) -match $nameRx })
    $osppRows=@($osppRows | Where-Object { ([string]$_.Product) -notmatch '(?i)Visio|Project' -and ([string]$_.Product) -match $nameRx })
    if ($officeReleaseIds.Count -gt 0) {
        $osppRows=@($osppRows | Where-Object {
            Test-OfficeLegacyProductMatch ([string]$_.Product) $officeReleaseIds
        })
    }

    $selectedRows=@()
    $decision='CAN_XAC_MINH'
    if ($m365Install) {
        # Do not trust old KMS/Trial WMI records for a Microsoft 365 license.
        $selectedRows=@($vnextRows)
        $decision=Get-OfficeLicenseDecision $selectedRows
        # With Microsoft 365 the absence of *current-user* vNext/device tokens
        # is usable only when the audit actually ran as that interactive user.
        # Otherwise IT_admin may have no license while the staff user does.
        if ($decision -eq 'CAN_XAC_MINH' -and $vnextEmpty -and
            $vnextRows.Count -eq 0 -and $accountContext -eq 'MATCH') {
            $decision='CHUA_KICH_HOAT'
            $r.OfficeMethod='VNEXT_KHONG_TIM_THAY_LICENSE_TAI_KHOAN_DANG_DUNG'
            $selectedRows=@(New-Object PSObject -Property @{
                Product=($officeReleaseIds -join ', ');
                State='NO_LOCAL_VNEXT_OR_DEVICE_LICENSE';
                Method='vnextdiag - same interactive Windows user'
            })
            LicenseTrace 'Office Microsoft 365: khong tim thay license vNext/device cho tai khoan dang dang nhap; tu ghi CHUA_KICH_HOAT theo ngu canh nay.'
        }
    } else {
        $osppDecision=Get-OfficeLicenseDecision $osppRows
        if ($osppDecision -ne 'CAN_XAC_MINH') {
            # OSPP reports installed Office keys; prioritize it if it gives a
            # decisive result even when historical WMI entries disagree.
            $selectedRows=@($osppRows)
            $decision=$osppDecision
        } elseif ($wmiDecision -ne 'CAN_XAC_MINH') {
            $selectedRows=@($wmiRows)
            $decision=$wmiDecision
        } elseif ($vnextRows.Count -gt 0) {
            $selectedRows=@($vnextRows)
            $decision=Get-OfficeLicenseDecision $selectedRows
        }
    }

    if ($selectedRows.Count -gt 0) {
        $r.OfficeProduct=(@($selectedRows | Select-Object -ExpandProperty Product -Unique) -join ' | ')
        $r.OfficeMethod=(@($selectedRows | Select-Object -ExpandProperty Method -Unique) -join ', ')
    }
    if (-not $r.OfficeProduct -and $officeReleaseIds.Count -gt 0) {
        $r.OfficeProduct=($officeReleaseIds -join ' | ')
    }
    if ($decision -ne 'CAN_XAC_MINH') {
        $r.OfficeState=$decision
        if ($r.WordInstalled -eq 'CO') {$r.WordState=$decision}
        if ($r.ExcelInstalled -eq 'CO') {$r.ExcelState=$decision}
        $r.OfficeNote='Xac nhan tu dong trang thai Office qua '+$r.OfficeMethod+'. Trang thai kich hoat khong tu chung minh tinh hop le cua giay phep. Chi tiet: '+(@($selectedRows | ForEach-Object { ([string]$_.Product)+' = '+([string]$_.State) }) -join '; ')
        if ($m365Install -and $vnextEmpty -and $accountContext -eq 'MATCH' -and $decision -eq 'CHUA_KICH_HOAT') {
            $r.OfficeNote='vNextDiag khong thay license vNext hoac device cho cung tai khoan Windows dang dang nhap va dang khao sat. Office chua kich hoat TRONG NGU CANH TAI KHOAN NAY; khong ket luan cac tai khoan Windows khac khong co license.'
        }
        LicenseTrace ('Office tu dong: '+$decision+'; nguon='+$r.OfficeMethod+'; khong hoi thu cong.')
    } else {
        $r.OfficeMethod='WMI/OSPP/VNEXT_CHUA_DU_BANG_CHUNG'
        $r.OfficeState='CAN_XAC_MINH_OFFICE_LICENSE'
        if ($r.WordInstalled -eq 'CO') {$r.WordState='CAN_XAC_MINH_WORD_LICENSE'}
        if ($r.ExcelInstalled -eq 'CO') {$r.ExcelState='CAN_XAC_MINH_EXCEL_LICENSE'}
        $r.OfficeNote='Phat hien Office da cai nhung khong doc duoc trang thai kich hoat du tin cay; chi truong hop nay moi hoi xac nhan thu cong.'
        if ($m365Install) {
            if ($vnextEmpty -and $accountContext -eq 'OTHER') {
                $r.OfficeNote+=' vNext: NO_LICENSE cho tai khoan khao sat; nguoi dung Windows dang dang nhap KHAC tai khoan khao sat. Khong the ket luan license cua nguoi dung.'
            } elseif ($vnextEmpty -and $accountContext -eq 'UNKNOWN') {
                $r.OfficeNote+=' vNext: NO_LICENSE cho tai khoan khao sat, nhung khong xac dinh duoc co cung tai khoan Windows dang dang nhap hay khong.'
            } elseif (-not $vnextExecuted) {
                $r.OfficeNote+=' vNextDiag khong chay duoc/khong tim thay; co the do policy hoac thieu cong cu.'
            } else {
                $r.OfficeNote+=' vNextDiag chua tra du ket qua de ket luan license.'
            }
        }
        if ($officeReleaseIds.Count -gt 0) {$r.OfficeNote+=' SKU cai dat: '+($officeReleaseIds -join ', ')+'.'}
        LicenseTrace 'Office da cai, khong xac minh duoc license: se hoi thu cong.'
    }
    LicenseTrace ('Ket qua Office='+$r.OfficeState)
    return $r
}
