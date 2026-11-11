# Produces XLSX directly using .NET Open Packaging (WindowsBase), no Excel required.
# Compatible with Windows PowerShell 2.0 on Windows 7.
if (-not (Get-Command Read-Utf8Csv -ErrorAction SilentlyContinue)) {
    . (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'CsvCompat.ps1')
}
if (-not (Get-Command Merge-AuditManualEdits -ErrorAction SilentlyContinue)) {
    . (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'ManualEdits.ps1')
}
function ReportMessage([string]$message,[string]$level='INFO') {
    if (Get-Command SurveyLog -ErrorAction SilentlyContinue) {
        SurveyLog $message $level
    } else {
        $c='Gray'
        if($level -eq 'OK') {$c='Green'}
        if($level -eq 'WARN') {$c='Yellow'}
        if($level -eq 'ERROR') {$c='Red'}
        Write-Host $message -ForegroundColor $c
    }
}
function XmlText([object]$v) {
    $t=([string]$v)
    if ($t.Length -gt 32000) { $t=$t.Substring(0,32000) }
    $t=[regex]::Replace($t,'[\x00-\x08\x0B\x0C\x0E-\x1F]',' ')
    return [System.Security.SecurityElement]::Escape($t)
}
function ColLetter([int]$index) {
    $s=''
    while ($index -gt 0) {
        $index--
        $s=[char](65+($index%26))+$s
        $index=[int][math]::Floor($index/26)
    }
    return $s
}
# Create package URIs explicitly; do not rely on New-Object constructor
# argument binding, which can select the one-argument URI overload in legacy PS.
# Uri.TryCreate(string, UriKind, out Uri) is available in .NET Framework 2.0+.
function New-PartUri([string]$value) {
    $result=$null
    if(-not [System.Uri]::TryCreate($value,[System.UriKind]::Relative,[ref]$result)) {
        throw ('URI khong hop le trong XLSX: '+$value)
    }
    if($null -eq $result -or $result.IsAbsoluteUri) {
        throw ('URI can la relative trong XLSX: '+$value)
    }
    return $result
}
function WritePart($part,[string]$payload) {
    $stream=$part.GetStream([System.IO.FileMode]::Create,[System.IO.FileAccess]::Write)
    $enc=New-Object System.Text.UTF8Encoding -ArgumentList $false
    $w=New-Object System.IO.StreamWriter -ArgumentList @($stream,$enc)
    try {$w.Write($payload)} finally {$w.Dispose()}
}
# Better-looking XLSX on Win7/PowerShell 2.0 through raw SpreadsheetML.
# KhaoSat is the only XLSX worksheet; style helper also supports legacy role names.
function MakeSheet($headers,$rows,[bool]$filter,[string]$role='KhaoSat') {
    $n=@($headers).Count
    $sb=New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="utf-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    $frozen=2
    if($role -eq 'TongQuan') {$frozen=5}
    if($role -eq 'HuongDan') {$frozen=0}
    $topLeft=(ColLetter ($frozen+1))+'2'
    if($frozen -gt 0) {
        [void]$sb.Append('<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane xSplit="'+$frozen+'" ySplit="1" topLeftCell="'+$topLeft+'" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>')
    } else {
        [void]$sb.Append('<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    }
    [void]$sb.Append('<sheetFormatPr defaultRowHeight="30"/>')
    [void]$sb.Append('<cols>')
    for($i=1;$i -le $n;$i++) {
        $width=20
        $h=[string]$headers[$i-1]
        if($h -eq '__AUDIT_ROW_KEY__') {$width=0}
        if($h -eq 'STT') {$width=7}
        elseif($h -match 'HotFix|Cập nhật Mẫu Virus') {$width=25}
        elseif($h -match 'IP \(|MAC|SKU|TPM|Tầng|Loại máy|Guest|DHCP') {$width=17}
        elseif($h -eq 'Phần mềm Office') {$width=38}
        elseif($h -eq 'Mã hóa BitLocker' -or $h -eq 'Chính sách USB Storage') {$width=37}
        elseif($h -match 'Kích hoạt|BitLocker|PowerShell|SMBv1|Office|Kaspersky') {$width=26}
        elseif($h -match 'Hãng sản xuất|Model|Phần mềm|Antivirus|Screen Saver|Windows|Serial|Người phụ trách|Tên Cán bộ|Phòng ban') {$width=31}
        elseif($h -match 'Ghi chú|Lỗi|Chẩn đoán|Chi tiết|Policy|Firewall|Trạng thái|Hướng dẫn|NTP') {$width=35}
        if($role -eq 'HuongDan' -and $i -eq 1) {$width=30}
        if($role -eq 'HuongDan' -and $i -eq 2) {$width=95}
        if($role -eq 'ChiTiet' -and $width -gt 32) {$width=35}
        if(@('Chip TPM 2.0','SMBv1 (Server / Client)','Autorun / Autoplay Policy','NTP / W32Time',
             'Cập nhật HotFix gần nhất','Khóa màn hình / Screen Saver','Danh sách ứng dụng',
             'Antivirus đang dùng','Phiên bản Windows OS Build','Domain / Workgroup',
             'Kích hoạt Windows','Tường lửa Windows (Domain/Private/Public)','Tài khoản Guest') -contains $h) {$width=43}
        if($h -eq '__AUDIT_ROW_KEY__') {
            [void]$sb.Append('<col min="'+$i+'" max="'+$i+'" width="0" hidden="1" customWidth="1"/>')
        } else {
            [void]$sb.Append('<col min="'+$i+'" max="'+$i+'" width="'+$width+'" customWidth="1"/>')
        }
    }
    [void]$sb.Append('</cols><sheetData>')
    $all=@(,$headers)+@($rows)
    for($r=0;$r -lt $all.Count;$r++) {
        $cells=@($all[$r]); $rowIdx=$r+1
        # Keep height bounded even when a cell contains long security notes.
        $height=if($r -eq 0){64}else{36}
        if($role -eq 'HuongDan' -and $r -gt 0) {$height=52}
        if($r -gt 0) {
            foreach($cellItem in $cells) {
                $line=[string]$cellItem
                if($line.Length -gt 135 -or $line -match '\r?\n' -or $line -match 'LAN: .+ \| Wi-Fi:') {$height=if($role -eq 'HuongDan'){72}else{60}; break}
            }
        }
        if($r -gt 0){
            foreach($colName in @('Mã hóa BitLocker','Chính sách USB Storage')) {
                for($k=0;$k -lt $n;$k++) {
                    if([string]$headers[$k] -ne $colName -or $k -ge $cells.Count){continue}
                    $pieces=@(([string]$cells[$k]) -split ' \| ')
                    if($pieces.Count -ge 5){$height=[math]::Max($height,112)}
                    elseif($pieces.Count -ge 3){$height=[math]::Max($height,80)}
                }
            }
        }
        if($r -gt 0) {
            foreach($hName in @('Chip TPM 2.0','SMBv1 (Server / Client)','Autorun / Autoplay Policy',
                'NTP / W32Time','Cập nhật HotFix gần nhất','Khóa màn hình / Screen Saver',
                'Danh sách ứng dụng','Antivirus đang dùng','Phiên bản Windows OS Build','Domain / Workgroup',
                'Kích hoạt Windows','Tường lửa Windows (Domain/Private/Public)','Tài khoản Guest')) {
                for($k=0;$k -lt $n;$k++) {
                    if([string]$headers[$k] -ne $hName -or $k -ge $cells.Count){continue}
                    $v=[string]$cells[$k]
                    if(-not $v){continue}
                    $parts=@($v -split ' \| ')
                    $needed=[math]::Min(200, [math]::Max(36,($parts.Count*19+12)))
                    if($needed -gt $height){$height=$needed}
                }
            }
        }
        [void]$sb.Append('<row r="'+$rowIdx+'" ht="'+$height+'" customHeight="1">')
        for($c=0;$c -lt $n;$c++) {
            $raw=''
            if($c -lt $cells.Count -and $null -ne $cells[$c]) {$raw=[string]$cells[$c]}
            $h=[string]$headers[$c]
            # Presentation only; raw data in CSV and the original inventory remains unchanged.
            if($r -gt 0 -and $h -match 'HotFix' -and $raw -notmatch '^\s*- ') {
                $raw=$raw -replace '\s+/\s+([12][0-9]{3}-[0-9]{2}-[0-9]{2})',([string][char]10+'$1')
            }
            if($r -gt 0 -and $h -match 'Antivirus đang dùng|Tường lửa Windows|DNS Server đang dùng|Phần mềm Office|Địa chỉ MAC \(LAN / Wi-Fi\)|^Mã hóa BitLocker$|^Chính sách USB Storage$|^Chip TPM 2\.0$|^Phiên bản Windows OS Build$|^SMBv1 \(Server / Client\)$|^Autorun / Autoplay Policy$|^NTP / W32Time$|^Cập nhật HotFix gần nhất$|^Khóa màn hình / Screen Saver$|^Danh sách ứng dụng$|^Domain / Workgroup$|^Kích hoạt Windows$|^Tài khoản Guest$') {
                $raw=$raw.Replace(' | ',([string][char]10))
            }
            $v=XmlText $raw
            $address=(ColLetter ($c+1))+$rowIdx
            # 0 plain; 1 header; 2 zebra white; 3 zebra ice; 4 warning; 5 critical; 6 positive; 7 number.
            $style=2
            if(($r % 2) -eq 0) {$style=3}
            if($r -eq 0) {$style=1}
            elseif($h -eq 'STT') {$style=7}
            elseif($raw -match '^(CAN_XAC_MINH|CAN_DOI_CHIEU|CHUA_|KHONG_XAC_DINH|LOI_|WMI_|THIEU_QUYEN|QUET_QUA_THOI_GIAN)' -or $raw -match '(?i)Trạng thái: (CAN_XAC_MINH|CAN_DOI_CHIEU|CHUA_QUET|KHONG_XAC_DINH|LOI_|THIEU_QUYEN)') {$style=4}
            elseif($raw -match '^(CAN_RA_SOAT|GUEST_ENABLED|PROTECTION OFF - ENCRYPTED|CHUA_KICH_HOAT|OFF - FULLY DECRYPTED)' -or $raw -match '(?i)Trạng thái: (CAN_RA_SOAT|GUEST_ENABLED|PROTECTION OFF - ENCRYPTED|CHUA_KICH_HOAT)') {$style=5}
            elseif($raw -match '^(DA_KICH_HOAT|ON - FULLY ENCRYPTED|TPM_2\.0|GUEST_DISABLED|DAY_DU|DA_XUAT_DANH_SACH)' -or $raw -match '(?i)Trạng thái: (DA_KICH_HOAT|GUEST_DISABLED)') {$style=6}
            if($r -gt 0 -and $h -eq 'STT' -and $raw -match '^[0-9]+$') {
                [void]$sb.Append('<c r="'+$address+'" s="'+$style+'" t="n"><v>'+$raw+'</v></c>')
            } else {
                [void]$sb.Append('<c r="'+$address+'" s="'+$style+'" t="inlineStr"><is><t xml:space="preserve">'+$v+'</t></is></c>')
            }
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')
    if($filter -and $n -gt 0) {
        # Include the hidden row-key in the sort/filter range: Excel must move
        # its ID with the row when the user sorts by a visible column.
        [void]$sb.Append('<autoFilter ref="A1:'+(ColLetter $n)+$all.Count+'"/>')
    }
    [void]$sb.Append('<pageMargins left="0.25" right="0.25" top="0.45" bottom="0.45" header="0.2" footer="0.2"/>')
    [void]$sb.Append('</worksheet>')
    return $sb.ToString()
}

function AddReportProperty([object]$obj,[string]$key,[object]$value) {
    $s=''
    if($null -ne $value) {$s=[string]$value}
    $obj | Add-Member -MemberType NoteProperty -Name $key -Value $s
}
function Raw([object]$rec,[string]$key) {
    $p=$rec.PSObject.Properties[$key]
    if($p -and $null -ne $p.Value) {return ([string]$p.Value)}
    return ''
}
# Consolidate several read-only evidence fields into a single cell.
# Keep the underlying scan data and diagnostic log unchanged.
function BitLockerSummary([object]$item) {
    $state=Raw $item 'Mã hóa BitLocker (ổ Windows)'
    if(-not $state){return ''}
    if($state -eq 'OFF / SUSPENDED / PARTIAL') {
        $state='CAN_XAC_MINH - KET_QUA_V5_CU'
    }
    $pieces=@($state)
    $fields=@(
        @('BitLocker Conversion','Mã hóa'),
        @('BitLocker Protection','Bảo vệ'),
        @('BitLocker Encrypted Percent','Tỷ lệ mã hóa'),
        @('BitLocker Source','Nguồn')
    )
    foreach($entry in $fields){
        $val=Raw $item $entry[0]
        if(-not $val){continue}
        if($entry[0] -eq 'BitLocker Encrypted Percent' -and $val -notmatch '%$'){$val+='%'}
        $pieces+=([string]$entry[1]+': '+$val)
    }
    return ($pieces -join ' | ')
}
function UsbStorageSummary([object]$item) {
    $state=Raw $item 'USB Storage Policy'
    if(-not $state){return ''}
    $pieces=@($state)
    foreach($entry in @(
        @('USBSTOR Start','USBSTOR'),
        @('USB Deny_All Policy','DenyAll'),
        @('USB Deny_Read Policy','DenyRead'),
        @('USB Deny_Write Policy','DenyWrite'),
        @('USB WriteProtect','WriteProtect')
    )) {
        $value=Raw $item $entry[0]
        if($value){$pieces+=([string]$entry[1]+': '+$value)}
    }
    return ($pieces -join ' | ')
}
# Report-only grouping: do not alter original scan snapshots.
# ASCII list marker keeps PS2/Windows 7 compatible; XLSX renders each item on its own line.
function GroupAuditFields([object]$item,[string]$state,[object[]]$fields) {
    $parts=@()
    if($state -and $state.Trim()) {$parts+=('- Trạng thái: '+$state.Trim())}
    foreach($entry in $fields) {
        $source=[string]$entry[0]
        $label=[string]$entry[1]
        $val=Raw $item $source
        if(-not $val -or -not $val.Trim()) {continue}
        $val=[regex]::Replace($val,'[\r\n]+','; ').Replace(' | ','; ').Trim()
        $parts+=('- '+$label+': '+$val)
    }
    return ($parts -join ' | ')
}

function ReportHeaderOrder {
    # Columns A:Q follow the shared audit screenshot; R onward add survey context.
    return @(
        'STT','Mã Máy PC','Tên Cán bộ Sử dụng','Phòng ban / Bộ phận',
        'Hãng sản xuất & Model','Số Serial Number (Win32 BIOS)',
        'Địa chỉ IP (đang dùng)','Địa chỉ MAC (LAN / Wi-Fi)',
        'Phiên bản Windows OS Build','Quyền Admin Local Audit',
        'Rà soát Phần mềm Crack / Lạ','Trạng thái Antivirus Real-time',
        'Ngày Cập nhật Mẫu Virus','Chip TPM 2.0','Mã hóa BitLocker',
        'Phần mềm Office','Phần mềm Kaspersky',
        'Tầng','Loại máy','Antivirus đang dùng',
        'Công cụ điều khiển từ xa',
        'Kích hoạt Windows','Lý do cần xác minh',
        'Mã SKU','BaseBoard Product','Serial máy - Service Tag',
        'Mã tài sản SMBIOS','Phiên bản BIOS',
        'Thời gian khảo sát',
        'Tường lửa Windows (Domain/Private/Public)','Cập nhật HotFix gần nhất',
        'Chính sách USB Storage','Tài khoản Guest','DNS Server đang dùng','Khóa màn hình / Screen Saver',
        'Domain / Workgroup','IP DHCP hoặc Static','SMBv1 (Server / Client)',
        'Autorun / Autoplay Policy','NTP / W32Time','Danh sách ứng dụng'
    )
}
function ToSummary([object]$item,[int]$seq) {
    $o=New-Object PSObject
    $make=(Raw $item 'Hãng máy'); $model=(Raw $item 'Model máy')
    $both=(($make+' '+$model).Trim())
    $serial=(Raw $item 'Serial BIOS')
    if(-not $serial) {$serial=Raw $item 'Serial máy - Service Tag'}
    # Display separately identified physical LAN/Wi-Fi in ONE Excel column.
    # Do not infer other/virtual adapters to be a physical LAN or Wi-Fi card.
    $lanMac=Raw $item 'MAC LAN'
    $wifiMac=Raw $item 'MAC Wi-Fi'
    $macParts=@()
    if($lanMac) {$macParts+=('LAN: '+$lanMac)}
    if($wifiMac) {$macParts+=('Wi-Fi: '+$wifiMac)}
    $mac=($macParts -join ' | ')
    if(-not $mac) {$mac='CHUA_LAY_DUOC_MAC_LAN_WIFI'}
    $ip=Raw $item 'Địa chỉ IPv4'
    $win=Raw $item 'Windows OS - Version/Build'
    if(-not $win) {$win=((Raw $item 'Windows (loại/edition)')+' (Build '+(Raw $item 'Windows build')+')').Trim()}
    # Historical scans stored the same SMBIOS field as 'Model mainboard'.
    # Reuse it when rebuilding; never overwrite original snapshots.
    $baseBoardProduct=Raw $item 'BaseBoard Product'
    if(-not $baseBoardProduct) {$baseBoardProduct=Raw $item 'Model mainboard'}
    $assetTag=Raw $item 'Mã tài sản SMBIOS'
    if(-not $assetTag) {$assetTag='CHUA_CO_DU_LIEU_ASSET_TAG'}
    $office=Raw $item 'Phần mềm Office'
    if(-not $office) { $office=(Raw $item 'Sản phẩm Office'); if(-not $office){$office='CAN_XAC_MINH'} }
    # For legacy snapshots preserve Word/Excel evidence inside the only Office column.
    if($office -notmatch 'Word:|Excel:') {
        $wordSeen=Raw $item 'Word đã cài'
        $excelSeen=Raw $item 'Excel đã cài'
        $wordLicense=Raw $item 'Word kích hoạt'
        $excelLicense=Raw $item 'Excel kích hoạt'
        if($wordSeen) {
            $office+=' | Word: '+$wordSeen
            if($wordSeen -eq 'CO' -and $wordLicense) { $office+=' ['+$wordLicense+']' }
        }
        if($excelSeen) {
            $office+=' | Excel: '+$excelSeen
            if($excelSeen -eq 'CO' -and $excelLicense) { $office+=' ['+$excelLicense+']' }
        }
    }
    $windowsGroup=GroupAuditFields $item $win @(
        @('Windows version','Version'),@('Windows Release','Release')
    )
    $avGroup=GroupAuditFields $item (Raw $item 'Antivirus đang dùng') @(
        @('Antivirus - Số sản phẩm phát hiện','Số sản phẩm'),
        @('Antivirus - Phiên bản từng sản phẩm','Phiên bản')
    )
    $tpmGroup=GroupAuditFields $item (Raw $item 'Chip TPM 2.0') @(
        @('Phiên bản TPM','Phiên bản'),@('TPM Source','Nguồn')
    )
    $smbGroup=GroupAuditFields $item (Raw $item 'SMBv1 trạng thái cấu hình') @(
        @('SMBv1 Server','Server'),@('SMBv1 Client','Client')
    )
    $autoGroup=GroupAuditFields $item (Raw $item 'Autorun / Autoplay Policy') @(
        @('Autorun Registry HKLM','HKLM'),@('Autorun Registry HKCU','HKCU'),
        @('AutoPlay UI (HKCU)','AutoPlay UI'),@('AutoPlay - Tài khoản kiểm tra','Tài khoản'),
        @('Autorun - Chính sách chi tiết','Chính sách chi tiết')
    )
    $ntpGroup=GroupAuditFields $item (Raw $item 'NTP W32Time') @(
        @('NTP W32Time Service','Dịch vụ'),@('NTP Client Type','Client type'),
        @('NTP Server Config','Server config'),@('NTP Source','Nguồn đồng bộ'),
        @('NTP Last Successful Sync','Lần đồng bộ thành công'),@('NTP Notes','Ghi chú')
    )
    $hotfixGroup=GroupAuditFields $item (Raw $item 'Windows HotFix gần nhất') @(
        @('HotFix KB','KB'),@('HotFix InstalledOn','Ngày cài'),
        @('HotFix 3 gần nhất','3 bản gần nhất'),@('HotFix Count','Số bản ghi'),
        @('HotFix Source','Nguồn')
    )
    $screenGroup=GroupAuditFields $item (Raw $item 'Screen Saver Lock Status') @(
        @('ScreenSaveActive','Bật Screen Saver'),@('ScreenSaverIsSecure','Yêu cầu đăng nhập'),
        @('ScreenSaveTimeOut','Chờ (giây)'),
        @('Machine Inactivity Limit Seconds','Giới hạn khóa máy (giây)'),
        @('ScreenLock Registry User','Tài khoản được kiểm tra')
    )
    $appGroup=GroupAuditFields $item (Raw $item 'Trạng thái danh sách ứng dụng') @(
        @('Số ứng dụng tìm thấy','Số ứng dụng'),@('File danh sách ứng dụng','Tệp CSV')
    )
    $windowsLicenseState=Raw $item 'Windows kích hoạt'
    if(-not $windowsLicenseState) {$windowsLicenseState='CHUA_QUET_BAN_MOI'}
    $winLicenseGroup=GroupAuditFields $item $windowsLicenseState @(
        @('Windows License Source','Nguồn kiểm tra'),
        @('Windows License Note','Chi tiết')
    )
    $firewallGroup=GroupAuditFields $item (Raw $item 'Windows Firewall Status') @(
        @('Firewall Current Profiles','Profiles hiện tại'),
        @('Firewall Source','Nguồn kiểm tra')
    )
    $guestGroup=GroupAuditFields $item (Raw $item 'Guest Account Status') @(
        @('Guest Account Name','Tài khoản'),
        @('Guest Account Enabled','Enabled')
    )
    $domainType=Raw $item 'Domain / Workgroup'
    $domainName=Raw $item 'Tên Domain / Workgroup'
    if(-not $domainType) {$domainType='CAN_XAC_MINH'}
    if(-not $domainName) {$domainName='CAN_XAC_MINH'}
    # One report column; both WMI values remain in DuLieuGoc.
    $domainSummary='- Loai: '+$domainType+' | - Ten: '+$domainName
    $values=@(
        $seq,(Raw $item 'Tên máy'),(Raw $item 'Người phụ trách máy'),(Raw $item 'Phòng ban'),
        $both,$serial,$ip,$mac,$windowsGroup,
        (Raw $item 'Quyền Admin Local Audit'),(Raw $item 'Rà soát phần mềm Crack / Lạ'),
        (Raw $item 'Trạng thái Antivirus Real-time'),(Raw $item 'Ngày cập nhật mẫu virus'),
        $tpmGroup,(BitLockerSummary $item),
        $office,(Raw $item 'Phần mềm Kaspersky'),
        (Raw $item 'Tầng'),(Raw $item 'Loại máy'),$avGroup,
        (Raw $item 'Công cụ điều khiển từ xa'),
        $winLicenseGroup,(Raw $item 'Giải thích mục cần xác minh'),
        (Raw $item 'Product Number - SKU'),$baseBoardProduct,(Raw $item 'Serial máy - Service Tag'),
        $assetTag,(Raw $item 'Phiên bản BIOS'),
        (Raw $item 'Thời gian'),
        $firewallGroup,$hotfixGroup,
        (UsbStorageSummary $item),$guestGroup,
        (Raw $item 'DNS Client Servers'),$screenGroup,
        $domainSummary,(Raw $item 'IP DHCP hoặc Static'),
        $smbGroup,$autoGroup,$ntpGroup,
        $appGroup
    )
    $headers=ReportHeaderOrder
    for($i=0;$i -lt $headers.Count;$i++) {AddReportProperty $o ([string]$headers[$i]) $values[$i]}
    # Historical rows never get marked clean automatically when new checks were not available.
    if(-not (Raw $item 'Quyền Admin Local Audit')) {$o.'Quyền Admin Local Audit'='CHUA_QUET_BAN_MOI'}
    if(-not (Raw $item 'Rà soát phần mềm Crack / Lạ')) {$o.'Rà soát Phần mềm Crack / Lạ'='CHUA_QUET_BAN_MOI'}
    if(-not (Raw $item 'Trạng thái Antivirus Real-time')) {$o.'Trạng thái Antivirus Real-time'='CHUA_QUET_BAN_MOI'}
    if(-not (Raw $item 'Ngày cập nhật mẫu virus')) {$o.'Ngày Cập nhật Mẫu Virus'='CHUA_QUET_BAN_MOI'}
    if(-not (Raw $item 'Chip TPM 2.0')) {
        if($tpmGroup){$o.'Chip TPM 2.0'='- Trạng thái: CHUA_QUET_BAN_MOI | '+$tpmGroup}
        else{$o.'Chip TPM 2.0'='- Trạng thái: CHUA_QUET_BAN_MOI'}
    }
    if(-not (Raw $item 'Mã hóa BitLocker (ổ Windows)')) {$o.'Mã hóa BitLocker'='CHUA_QUET_BAN_MOI'}
    # Legacy ambiguous state is already labelled by BitLockerSummary.
    if(-not (Raw $item 'Phần mềm Kaspersky')) {$o.'Phần mềm Kaspersky'='CHUA_QUET_BAN_MOI'}
    $oldChecks=@(
       @('Windows Firewall Status','Tường lửa Windows (Domain/Private/Public)'),
       @('Windows HotFix gần nhất','Cập nhật HotFix gần nhất'),
       @('USB Storage Policy','Chính sách USB Storage'),
       @('Guest Account Status','Tài khoản Guest'),
       @('DNS Client Servers','DNS Server đang dùng'),
       @('Screen Saver Lock Status','Khóa màn hình / Screen Saver')
    )
    foreach($pair in $oldChecks) {
       if(-not (Raw $item $pair[0])) {
           if(@('Cập nhật HotFix gần nhất','Khóa màn hình / Screen Saver','Tài khoản Guest','Tường lửa Windows (Domain/Private/Public)') -contains [string]$pair[1]) {
               $existing=[string]$o.PSObject.Properties[$pair[1]].Value
               if($existing){$o.PSObject.Properties[$pair[1]].Value='- Trạng thái: CHUA_QUET_BAN_MOI | '+$existing}
               else {$o.PSObject.Properties[$pair[1]].Value='- Trạng thái: CHUA_QUET_BAN_MOI'}
           } else { $o.PSObject.Properties[$pair[1]].Value='CHUA_QUET_BAN_MOI' }
       }
    }

    # Old V5 snapshots do not contain newly introduced PA05 checks. Do not infer a clean status.
    foreach($pair in @(
       @('Domain / Workgroup','Domain / Workgroup'),
       @('IP DHCP hoặc Static','IP DHCP hoặc Static'),
       @('SMBv1 trạng thái cấu hình','SMBv1 (Server / Client)'),
       @('Autorun / Autoplay Policy','Autorun / Autoplay Policy'),
       @('NTP W32Time','NTP / W32Time'),
       @('Trạng thái danh sách ứng dụng','Danh sách ứng dụng')
    )) {
       if(-not (Raw $item $pair[0])) {
           if(@('SMBv1 (Server / Client)','Autorun / Autoplay Policy','NTP / W32Time','Danh sách ứng dụng') -contains [string]$pair[1]) {
               $existing=[string]$o.PSObject.Properties[$pair[1]].Value
               if($existing){$o.PSObject.Properties[$pair[1]].Value='- Trạng thái: CHUA_QUET_BAN_MOI | '+$existing}
               else {$o.PSObject.Properties[$pair[1]].Value='- Trạng thái: CHUA_QUET_BAN_MOI'}
           } else { $o.PSObject.Properties[$pair[1]].Value='CHUA_QUET_BAN_MOI' }
       }
    }
return $o
}

# Fields already represented by the compact summary, or explicitly hidden by the user.
# Raw device snapshots in DuLieuGoc are not deleted or rewritten.
function Exclude-ReportField([string]$name) {
    if($name -match '(?i)EOS/EOL|EOL/EOSL|^Ngày EOS|^Ngày EOL|^Nguồn EOS') { return $true }
    if(@(
        '__AUDIT_ROW_KEY__','Chất lượng dữ liệu','Lỗi truy vấn','Lỗi kiểm tra cấu hình bổ sung',
        'Kiểu thực thi Step 5','Step 5 LICENSE (giây)',
        'Step 5 SECURITY (giây)','Step 5 POLICIES (giây)',
        'Dấu hiệu nghi vấn (chi tiết)',
        'BitLocker - Trạng thái mã hóa','BitLocker - Bảo vệ khóa',
        'BitLocker - Tỷ lệ mã hóa (%)','BitLocker - Phương pháp kiểm tra',
        'BitLocker Conversion','BitLocker Protection','BitLocker Encrypted Percent','BitLocker Source',
        'Phiên bản PowerShell','Windows PowerShell',
        'Rà soát IP/MAC toàn mạng','NAC / Port Security',
        'USBSTOR Start','USB Deny_All Policy','USB Deny_Read Policy',
        'USB Deny_Write Policy','USB WriteProtect',
        'BitLocker Diagnostic','TPM Diagnostic','BitLocker - Chẩn đoán lỗi','TPM - Chẩn đoán lỗi',
        'Mục chất lượng dữ liệu','Trạng thái EOS/EOL phần cứng',
        'Firewall Domain','Firewall Private','Firewall Public','Windows Firewall Status',
        'Firewall Current Profiles','Firewall Source',
        'Phần mềm Office','Sản phẩm Office','Office License Method','Office License Note',
        'Word đã cài','Word kích hoạt','Excel đã cài','Excel kích hoạt',
        'Antivirus - Real-time từng sản phẩm','Antivirus - Cập nhật từng sản phẩm',
        'Antivirus - Dịch vụ từng sản phẩm','Antivirus - Giấy phép từng sản phẩm',
        'Antivirus - Nguồn kiểm tra chi tiết','Antivirus - Lần quét gần nhất từng sản phẩm',
        'Antivirus - Engine / Signature từng sản phẩm',
        'Tên máy','Hãng máy','Model máy','Serial BIOS','MAC LAN','MAC Wi-Fi','MAC khác','Địa chỉ IPv4',
        'Product Number - SKU','BaseBoard Product','Model mainboard',
        'Mã tài sản SMBIOS','Phiên bản BIOS','Serial máy - Service Tag','Thời gian',
        'Windows OS - Version/Build','Windows (loại/edition)','Windows build',
        'Quyền Admin Local Audit','Rà soát phần mềm Crack / Lạ',
        'Trạng thái Antivirus Real-time','Ngày cập nhật mẫu virus',
        'Chip TPM 2.0','Mã hóa BitLocker (ổ Windows)',
        'Phần mềm Kaspersky','Antivirus đang dùng','Dấu hiệu nghi vấn chi tiết',
        'Công cụ điều khiển từ xa','Ghi chú Antivirus','Windows kích hoạt',
        'Tầng','Phòng ban','Người phụ trách máy','Loại máy',
        'Windows HotFix gần nhất','USB Storage Policy','Guest Account Status',
        'Guest Account Name','Guest Account Enabled',
        'Windows License Source','Windows License Note','Giải thích mục cần xác minh',
        'DNS Client Servers','Screen Saver Lock Status','Domain / Workgroup',
        'Tên Domain / Workgroup','IP DHCP hoặc Static','SMBv1 trạng thái cấu hình',
        'Autorun / Autoplay Policy','NTP W32Time','Trạng thái danh sách ứng dụng',
        'Số ứng dụng tìm thấy','File danh sách ứng dụng','Rà soát IP/MAC toàn mạng',
        'NAC Port Security','Antivirus - Số sản phẩm phát hiện',
        'Antivirus - Phiên bản từng sản phẩm',
        'Windows version',
        'Windows Release',
        'Tiến trình đang chạy Admin',
        'Ghi chú SMBv1',
        'Ghi chú Autorun',
        'Autorun - Chính sách chi tiết',
        'Phiên bản TPM',
        'TPM Source',
        'SMBv1 Server',
        'SMBv1 Client',
        'Autorun Registry HKLM',
        'Autorun Registry HKCU',
        'AutoPlay UI (HKCU)',
        'AutoPlay - Tài khoản kiểm tra',
        'NTP W32Time Service',
        'NTP Client Type',
        'NTP Server Config',
        'NTP Source',
        'NTP Last Successful Sync',
        'NTP Notes',
        'HotFix KB',
        'HotFix InstalledOn',
        'HotFix 3 gần nhất',
        'HotFix Count',
        'HotFix Source',
        'ScreenSaveActive',
        'ScreenSaverIsSecure',
        'ScreenSaveTimeOut',
        'Machine Inactivity Limit Seconds',
        'ScreenLock Registry User'
    ) -contains $name) { return $true }
    return $false
}
function BuildXlsx([object[]]$records,[object[]]$details,[string]$dest,[string[]]$recordKeys) {
    if($records.Count -eq 0) {throw 'Khong co du lieu.'}
    # ONE XLSX SHEET ONLY. Every visible cell now comes from ONE merged
    # report row, keyed by the stable snapshot ID. Do not append fields from
    # $details[$idx]: Merge-AuditManualEdits may add rows or change order,
    # which previously made UUID/CPU/RAM/network/evidence columns blank/wrong.
    # Technical evidence was copied into each report row BEFORE manual edits.
    $headers=@($records[0].PSObject.Properties | ForEach-Object {$_.Name})
    $allHeaders=@($headers)+@('__AUDIT_ROW_KEY__')
    if($recordKeys.Count -ne $records.Count) {throw 'Audit record ID count mismatch - Excel not changed.'}
    $body=@()
    for($idx=0; $idx -lt $records.Count; $idx++) {
        $rec=$records[$idx]
        $values=@()
        foreach($h in $headers) { $values+=(Raw $rec $h) }
        $values+=([string]$recordKeys[$idx])
        $body+=,$values
    }
    Add-Type -AssemblyName WindowsBase -ErrorAction Stop
    Assert-AuditWorkbookWritable $dest
    $tmp=$dest+'.creating'
    if(Test-Path -LiteralPath $tmp) {Remove-Item -LiteralPath $tmp -Force}
    $p=[System.IO.Packaging.Package]::Open($tmp,[System.IO.FileMode]::Create,[System.IO.FileAccess]::ReadWrite)
    try {
        $base='http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
        $uriBook=New-PartUri '/xl/workbook.xml'
        $book=$p.CreatePart($uriBook,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',[System.IO.Packaging.CompressionOption]::Normal)
        $p.CreateRelationship($uriBook,[System.IO.Packaging.TargetMode]::Internal,($base+'officeDocument')) | Out-Null
        # Inventory.xlsx has exactly ONE worksheet, named KhaoSat.
        $titles=@('KhaoSat')
        $sheetData=@((MakeSheet $allHeaders $body $true 'KhaoSat'))
        $sheetsXml=''
        for($i=0;$i -lt $titles.Count;$i++) {
            $id=$i+1
            $relative='worksheets/sheet'+$id+'.xml'
            $uri=New-PartUri ('/xl/'+$relative)
            $part=$p.CreatePart($uri,'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',[System.IO.Packaging.CompressionOption]::Normal)
            $ref=$book.CreateRelationship((New-PartUri $relative),[System.IO.Packaging.TargetMode]::Internal,($base+'worksheet'))
            $sheetsXml+='<sheet name="'+$titles[$i]+'" sheetId="'+$id+'" r:id="'+$ref.Id+'"/>'
            WritePart $part ([string]$sheetData[$i])
        }
        $uriStyles=New-PartUri '/xl/styles.xml'
        $styles=$p.CreatePart($uriStyles,'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',[System.IO.Packaging.CompressionOption]::Normal)
        $book.CreateRelationship((New-PartUri 'styles.xml'),[System.IO.Packaging.TargetMode]::Internal,($base+'styles')) | Out-Null
        WritePart $book ('<?xml version="1.0" encoding="utf-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+$sheetsXml+'</sheets></workbook>')
        # Calibri is installed on Windows 7 onwards; wrap text and use soft banded rows.
        $styleXml='<?xml version="1.0" encoding="utf-8" standalone="yes"?>'+
          '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'+
          '<fonts count="5">'+
          '<font><sz val="11"/><name val="Calibri"/><color rgb="FF233047"/></font>'+
          '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'+
          '<font><b/><sz val="11"/><color rgb="FF945C0B"/><name val="Calibri"/></font>'+
          '<font><b/><sz val="11"/><color rgb="FFAB3232"/><name val="Calibri"/></font>'+
          '<font><b/><sz val="11"/><color rgb="FF176648"/><name val="Calibri"/></font>'+
          '</fonts>'+
          '<fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'+
          '<fill><patternFill patternType="solid"><fgColor rgb="FF264E7B"/><bgColor indexed="64"/></patternFill></fill>'+
          '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>'+
          '<fill><patternFill patternType="solid"><fgColor rgb="FFF0F5FB"/><bgColor indexed="64"/></patternFill></fill>'+
          '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF1D9"/><bgColor indexed="64"/></patternFill></fill>'+
          '<fill><patternFill patternType="solid"><fgColor rgb="FFFFE9E8"/><bgColor indexed="64"/></patternFill></fill></fills>'+
          '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'+
          '<border><left/><right/><top/><bottom style="hair"><color rgb="FFE0E6EE"/></bottom><diagonal/></border></borders>'+
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'+
          '<cellXfs count="8">'+
          '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'+
          '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'+
          '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>'+
          '<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>'+
          '<xf numFmtId="0" fontId="2" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>'+
          '<xf numFmtId="0" fontId="3" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>'+
          '<xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>'+
          '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'+
          '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
        WritePart $styles $styleXml
    } finally { $p.Close() }
    Commit-AuditWorkbook -tempPath $tmp -targetPath $dest
    $script:lastReportPath=$dest
}
function Write-Reports([string]$rawFolder,[string]$outFolder) {
    $script:lastReportPath=$null
    $sourceFiles=@(Get-ChildItem -LiteralPath $rawFolder -Filter '*.csv' |
       Where-Object { $_.Name -like 'May_*.csv' -or $_.Name -like 'Device_*.csv' } | Sort-Object Name)
    $items=@()
    foreach($f in $sourceFiles) {
        try {
            $fileRows=@(Read-Utf8Csv -FilePath $f.FullName)
            for($j=0;$j -lt $fileRows.Count;$j++) {
                $x=$fileRows[$j]
                # A snapshot filename is immutable; survives changing names/serials in Excel.
                $key='SCAN_'+$f.Name+'_'+$j
                $x | Add-Member -MemberType NoteProperty -Name '__AUDIT_ROW_KEY__' -Value $key -Force
                $items+=,$x
            }
        } catch { ReportMessage ('Bo qua file hong: '+$f.Name+' - '+$_.Exception.Message) 'WARN' }
    }
    if($items.Count -eq 0) {throw 'Khong tim thay du lieu goc trong KhaoSat\DuLieuGoc.'}
    $items=@($items | Sort-Object -Property 'Thời gian')
    $front=@('Thời gian','Tầng','Phòng ban','Người phụ trách máy','Loại máy','Tên máy')
    $dheaders=@($front)
    # Stable column even for old backups that predate this dedicated field.
    if($dheaders -notcontains 'BaseBoard Product') {$dheaders+='BaseBoard Product'}
    foreach($h in @('Windows Firewall Status','Windows HotFix gần nhất','USB Storage Policy','Guest Account Status','DNS Client Servers','Screen Saver Lock Status')) {if($dheaders -notcontains $h){$dheaders+=$h}}
    foreach($record in $items) {
        foreach($prop in $record.PSObject.Properties) {
            # Bỏ mọi cột '(nhập sau)' từ dữ liệu gốc cũ khi dựng lại báo cáo.
            if($prop.Name -match '\(nhập sau\)$') {continue}
            # Do not drop fields needed by ToSummary; BuildXlsx filters duplicate columns later.
            if($prop.Name -match '(?i)EOS/EOL|EOL/EOSL|^Ngày EOS|^Ngày EOL|^Nguồn EOS') {continue}
            if(@('Lỗi truy vấn','Lỗi kiểm tra cấu hình bổ sung','BitLocker Diagnostic','TPM Diagnostic',
                  'BitLocker - Chẩn đoán lỗi','TPM - Chẩn đoán lỗi') -contains [string]$prop.Name) {continue}
            if($dheaders -notcontains $prop.Name) { $dheaders+=$prop.Name }
        }
    }
    $details=@()
    foreach($item in $items) {
        $d=New-Object PSObject
        foreach($h in $dheaders) {
            $v=Raw $item $h
            if($h -eq 'BaseBoard Product' -and -not $v) {$v=Raw $item 'Model mainboard'}
            AddReportProperty $d $h $v
        }
        $details+=$d
    }
    $summary=@()
    for($i=0;$i -lt $details.Count;$i++) {
        $one=ToSummary $details[$i] ($i+1)
        # Copy ALL raw-only visible evidence onto this SAME report row.
        # This guarantees the values survive Office auto/manual choices and
        # historical Excel rows; no positional join is done during XLSX output.
        foreach($h in $dheaders) {
            if(Exclude-ReportField ([string]$h)) {continue}
            if(-not $one.PSObject.Properties[$h]) {
                AddReportProperty $one ([string]$h) (Raw $details[$i] ([string]$h))
            }
        }
        $summary+=,$one
    }
    # Merge all existing Excel edits (including blanks) by hidden scan ID, then persisted backup.
    # Do this before exporting CSV or Excel; any ambiguous match aborts safely.
    $merged=Merge-AuditManualEdits -records $summary -details $details -outFolder $outFolder
    $summary=@($merged.Records)
    $recordKeys=[string[]]@($merged.Keys)
    # Detect evidence loss before writing: each nonempty raw value must reach the
    # report unless the operator has explicitly supplied another nonblank value.
    # Blank old Excel cells must never suppress collected WMI/network evidence.
    $evidenceFields=@('UUID','Hãng mainboard','Serial mainboard','CPU',
        'RAM (GB)','Ổ lưu trữ (model - GB)','Nguồn dữ liệu Antivirus',
        'Ngày mẫu virus Defender (tham khảo)','DNS per Adapter','Gateway',
        'Mục cần đối chiếu','Trường còn thiếu')
    $sourceIndex=@{}
    foreach($d in $details) {
        $scanKey=Raw $d '__AUDIT_ROW_KEY__'
        if($scanKey){$sourceIndex[$scanKey]=$d}
    }
    for($i=0;$i -lt $summary.Count;$i++) {
        $scanKey=[string]$recordKeys[$i]
        if(-not $sourceIndex.ContainsKey($scanKey)){continue} # manual-only rows
        $source=$sourceIndex[$scanKey]
        foreach($h in $evidenceFields) {
            $expected=Raw $source $h
            $actual=Raw $summary[$i] $h
            if($expected -and -not $actual) {
                # Legacy report may contain a generated blank. Recover the
                # actual immutable scan value rather than losing the field.
                if(-not $summary[$i].PSObject.Properties[$h]) {
                    AddReportProperty $summary[$i] $h $expected
                } else {
                    $summary[$i].PSObject.Properties[$h].Value=$expected
                }
                ReportMessage ('Phuc hoi du lieu goc tai '+$h+' (ma luot quet '+$scanKey+')') 'WARN'
            }
        }
    }
    Write-Utf8Csv -Records $summary -FilePath (Join-Path $outFolder 'Inventory.csv')
    # Inventory_Detail.csv keeps only technical fields not already shown in the main report.
    $detailForExport=@()
    foreach($sourceItem in $details) {
        $compactItem=New-Object PSObject
        foreach($prop in $sourceItem.PSObject.Properties) {
            if(-not (Exclude-ReportField ([string]$prop.Name))) {
                AddReportProperty $compactItem ([string]$prop.Name) $prop.Value
            }
        }
        if($compactItem.PSObject.Properties.Count -gt 0) { $detailForExport+=,$compactItem }
    }
    if($detailForExport.Count -gt 0) {
        Write-Utf8Csv -Records $detailForExport -FilePath (Join-Path $outFolder 'Inventory_Detail.csv')
    }
    $review=@($summary | Where-Object {
        $_.'Mục cần đối chiếu' -ne '' -or
        $_.'Rà soát Phần mềm Crack / Lạ' -like 'CAN_RA_SOAT*' -or
        $_.'Rà soát Phần mềm Crack / Lạ' -eq 'CHUA_QUET_BAN_MOI' -or
        $_.'Tường lửa Windows (Domain/Private/Public)' -match 'CHUA_QUET_BAN_MOI' -or
        $_.'Cập nhật HotFix gần nhất' -match 'CHUA_QUET_BAN_MOI' -or
        $_.'Chính sách USB Storage' -eq 'CHUA_QUET_BAN_MOI' -or
        $_.'Chính sách USB Storage' -match 'CAN_XAC_MINH|USBSTOR_MANUAL|KHONG_CO_USBSTOR' -or
        $_.'Tài khoản Guest' -match 'CHUA_QUET_BAN_MOI' -or
        $_.'DNS Server đang dùng' -eq 'CHUA_QUET_BAN_MOI' -or
        $_.'Khóa màn hình / Screen Saver' -match 'CHUA_QUET_BAN_MOI' -or
        $_.'SMBv1 (Server / Client)' -match 'CHUA_QUET_BAN_MOI' -or
        $_.'SMBv1 (Server / Client)' -match 'CAN_XAC_MINH|CAN_DOI_CHIEU|KHONG_CAU_HINH|ENABLED_CONFIG|START_|KHONG_TIM_THAY' -or
        $_.'Autorun / Autoplay Policy' -match 'CHUA_QUET_BAN_MOI' -or
        $_.'Autorun / Autoplay Policy' -match 'CAN_XAC_MINH|NEEDS_REVIEW|KHONG_TIM_THAY|NO_GPO_FOUND' -or
        $_.'NTP / W32Time' -match 'CHUA_QUET_BAN_MOI' -or
        $_.'Danh sách ứng dụng' -notmatch 'DA_XUAT_DANH_SACH_REGISTRY' -or
        $_.'Tài khoản Guest' -match 'GUEST_ENABLED' -or
        $_.'Trạng thái Antivirus Real-time' -like 'CAN_XAC_MINH*' -or
        $_.'Trạng thái Antivirus Real-time' -eq 'CHUA_QUET_BAN_MOI' -or
        $_.'Mã hóa BitLocker' -like 'CAN_XAC_MINH*' -or
        $_.'Mã hóa BitLocker' -like 'THIEU_QUYEN_ADMIN*' -or
        $_.'Mã hóa BitLocker' -like 'PROTECTION OFF - ENCRYPTED*' -or
        $_.'Mã hóa BitLocker' -like 'LOI_*' -or
        $_.'Mã hóa BitLocker' -like 'WMI_CLASS_HOAC_NAMESPACE_KHONG_CO*' -or
        $_.'Chip TPM 2.0' -match 'THIEU_QUYEN_ADMIN' -or
        $_.'Chip TPM 2.0' -match 'WMI_' -or
        $_.'Mã hóa BitLocker' -like 'KHONG_TIM_THAY*'
    })
    $reviewPath=Join-Path $outFolder 'NeedsReview.csv'
    if($review.Count -gt 0){Write-Utf8Csv -Records $review -FilePath $reviewPath}
    elseif(Test-Path -LiteralPath $reviewPath){Remove-Item -LiteralPath $reviewPath -Force}
    try {
        BuildXlsx -records $summary -details $details -dest (Join-Path $outFolder 'Inventory.xlsx') -recordKeys $recordKeys
        Save-AuditManualEdits -records $summary -keys $recordKeys -outFolder $outFolder
        ReportMessage ('Excel: '+$script:lastReportPath+' | '+$summary.Count+' luot quet | can kiem tra: '+$review.Count) 'OK'
    } catch {ReportMessage ('Loi tao Excel: '+$_.Exception.Message+'. CSV van duoc luu. Dong Inventory.xlsx neu dang mo, sau do chay Rebuild.') 'WARN'}
}
