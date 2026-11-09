# Preserve manually edited Excel cell VALUES on repeated scans or Rebuild.
# Windows PowerShell 2.0 / Windows 7+; uses built-in WindowsBase and XmlDocument only.
# No Excel COM, ImportExcel module, external executables or Windows configuration changes.
# The only worksheet is KhaoSat. __AUDIT_ROW_KEY__ is an internal HIDDEN column.
function AuditCellValue($cell,$ns,$strings) {
    if($null -eq $cell) {return ''}
    $typ=[string]$cell.GetAttribute('t')
    if($typ -eq 'inlineStr') {
        $parts=@($cell.SelectNodes('./x:is//x:t',$ns))
        $out=''
        foreach($p in $parts) {$out+=$p.InnerText}
        return $out
    }
    $v=$cell.SelectSingleNode('./x:v',$ns)
    if($null -eq $v) {
        # Formula without cached value cannot be safely restored by value-only merger.
        $formula=$cell.SelectSingleNode('./x:f',$ns)
        if($null -ne $formula) {return '='+$formula.InnerText}
        return ''
    }
    $txt=[string]$v.InnerText
    if($typ -eq 's') {
        $index=0
        if([int]::TryParse($txt,[ref]$index) -and $index -ge 0 -and $index -lt $strings.Count) {return [string]$strings[$index]}
    }
    if($typ -eq 'b') { if($txt -eq '1'){return 'TRUE'} else{return 'FALSE'} }
    return $txt
}
function Read-AuditWorkbook([string]$path) {
    Add-Type -AssemblyName WindowsBase -ErrorAction Stop
    $pkg=[System.IO.Packaging.Package]::Open($path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read)
    try {
        # Fail safely if a different workbook/sheet was accidentally placed at this path.
        $wbUri=New-PartUri '/xl/workbook.xml'
        if(-not $pkg.PartExists($wbUri)) {throw 'Khong phai Excel XLSX hop le.'}
        $wbStream=$pkg.GetPart($wbUri).GetStream([System.IO.FileMode]::Open,[System.IO.FileAccess]::Read)
        try {$wbXml=New-Object System.Xml.XmlDocument; $wbXml.Load($wbStream)} finally {$wbStream.Close()}
        $wbNs=New-Object System.Xml.XmlNamespaceManager($wbXml.NameTable)
        $wbNs.AddNamespace('x','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $sheetNames=@($wbXml.SelectNodes('//x:sheets/x:sheet',$wbNs))
        if($sheetNames.Count -ne 1 -or [string]$sheetNames[0].GetAttribute('name') -ne 'KhaoSat') {
            throw 'Inventory.xlsx phai co duy nhat mot sheet ten KhaoSat. Khong ghi de.'
        }
        # This tool creates the single worksheet at this fixed internal part path.
        $uri=New-PartUri '/xl/worksheets/sheet1.xml'
        if(-not $pkg.PartExists($uri)) {throw 'Khong tim thay KhaoSat sheet1.xml. Khong ghi de file.'}
        $part=$pkg.GetPart($uri)
        $stream=$part.GetStream([System.IO.FileMode]::Open,[System.IO.FileAccess]::Read)
        try {
            $xml=New-Object System.Xml.XmlDocument
            $xml.Load($stream)
        } finally {$stream.Close()}
        $ns=New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace('x','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $strings=@()
        $stringsUri=New-PartUri '/xl/sharedStrings.xml'
        if($pkg.PartExists($stringsUri)) {
            $ssStream=$pkg.GetPart($stringsUri).GetStream([System.IO.FileMode]::Open,[System.IO.FileAccess]::Read)
            try {$ssXml=New-Object System.Xml.XmlDocument; $ssXml.Load($ssStream)} finally {$ssStream.Close()}
            $ssNs=New-Object System.Xml.XmlNamespaceManager($ssXml.NameTable)
            $ssNs.AddNamespace('x','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            foreach($si in $ssXml.SelectNodes('//x:sst/x:si',$ssNs)) {
                $word=''
                foreach($textNode in $si.SelectNodes('.//x:t',$ssNs)) {$word+=$textNode.InnerText}
                $strings+=,$word
            }
        }
        $result=@(); $headers=@(); $headRow=$null
        $rowNodes=@($xml.SelectNodes('//x:sheetData/x:row',$ns))
        foreach($rowNode in $rowNodes) {
            $values=@{}
            foreach($cell in $rowNode.SelectNodes('./x:c',$ns)) {
                $ref=[string]$cell.GetAttribute('r')
                if($ref -notmatch '^([A-Z]+)[0-9]+$'){continue}
                $letters=$matches[1]; $col=0
                foreach($ch in $letters.ToCharArray()) {$col=$col*26+([int][char]$ch-64)}
                $values[$col]=[string](AuditCellValue $cell $ns $strings)
            }
            if($null -eq $headRow) {
                $headRow=$rowNode
                $max=0; foreach($num in $values.Keys){if($num -gt $max){$max=$num}}
                for($n=1;$n -le $max;$n++) {if($values.ContainsKey($n)) {$headers+=([string]$values[$n])} else {$headers+=''}}
                continue
            }
            $obj=New-Object PSObject
            for($n=1;$n -le $headers.Count;$n++) {
                $header=[string]$headers[$n-1]
                if(-not $header){continue}
                $value=''; if($values.ContainsKey($n)){$value=[string]$values[$n]}
                # XLSX display uses wrapped lines for several columns, but the
                # PS2 CSV reader expects ONE physical line per record.
                if($value -match '[\r\n]') {
                    $separator=' | '
                    if($header -match 'HotFix') {$separator=' / '}
                    $value=[regex]::Replace($value,'[\r\n]+',$separator)
                }
                $obj | Add-Member -MemberType NoteProperty -Name $header -Value $value -Force
            }
            $nonblank=$false
            foreach($field in $obj.PSObject.Properties) {if($field.Name -ne '__AUDIT_ROW_KEY__' -and [string]$field.Value -ne '') {$nonblank=$true;break}}
            if($nonblank) {$result+=,$obj}
        }
        return $result
    } finally {$pkg.Close()}
}
function Get-AuditMatch([object]$old,[object[]]$records,[string[]]$keys,[int]$rowIndex) {
    $key=Raw $old '__AUDIT_ROW_KEY__'
    if($key) {
        for($i=0;$i -lt $keys.Count;$i++){if($keys[$i] -eq $key){return $i}}
        return -1
    }
    # Compatibility for a PRE-UPGRADE workbook with no hidden IDs.
    # First match snapshot time + PC name + BIOS serial; refuse ambiguous matches.
    $found=@()
    for($i=0;$i -lt $records.Count;$i++) {
        $matchFields=0; $valid=$true
        foreach($h in @('Thời gian khảo sát','Mã Máy PC','Số Serial Number (Win32 BIOS)')) {
            $a=Raw $old $h; $b=Raw $records[$i] $h
            if($a -and $b) {if($a -ne $b){$valid=$false;break};$matchFields++}
        }
        if($valid -and $matchFields -ge 2) {$found+=,$i}
    }
    if($found.Count -eq 1) {return [int]$found[0]}
    # Legacy only: conservative position fallback, requiring two independent
    # matching identifying fields. Never match solely on a repeated PC name.
    if($rowIndex -lt $records.Count -and $rowIndex -ge 0) {
        $cand=$records[$rowIndex]
        $score=0
        foreach($h in @('Thời gian khảo sát','Mã Máy PC','Số Serial Number (Win32 BIOS)')) {
            $a=Raw $old $h; $b=Raw $cand $h
            if($a -and $b -and $a -eq $b) {$score++}
        }
        if($score -ge 2) {return $rowIndex}
    }
    return -1
}
# Preserve manual edits when previously separate technical columns become
# bullet lines inside one report cell. Only transfer values that differ from
# the immutable original scan data, to avoid copying obsolete generated values.
function Get-GroupedReportTarget([string]$oldName) {
    $map=@{
        'Số ứng dụng tìm thấy'=@('Danh sách ứng dụng','Số ứng dụng');
        'File danh sách ứng dụng'=@('Danh sách ứng dụng','Tệp CSV');
        'Trạng thái danh sách ứng dụng'=@('Danh sách ứng dụng','Trạng thái');
        'Antivirus - Số sản phẩm phát hiện'=@('Antivirus đang dùng','Số sản phẩm');
        'Antivirus - Phiên bản từng sản phẩm'=@('Antivirus đang dùng','Phiên bản');
        'Tên Domain / Workgroup'=@('Domain / Workgroup','Ten');
        'Windows License Source'=@('Kích hoạt Windows','Nguồn kiểm tra');
        'Windows License Note'=@('Kích hoạt Windows','Chi tiết');
        'Guest Account Name'=@('Tài khoản Guest','Tài khoản');
        'Guest Account Enabled'=@('Tài khoản Guest','Enabled');
        'Firewall Current Profiles'=@('Tường lửa Windows (Domain/Private/Public)','Profiles hiện tại');
        'Firewall Source'=@('Tường lửa Windows (Domain/Private/Public)','Nguồn kiểm tra');
        'Windows version'=@('Phiên bản Windows OS Build','Version');
        'Windows Release'=@('Phiên bản Windows OS Build','Release');
        'Phiên bản TPM'=@('Chip TPM 2.0','Phiên bản');
        'TPM Source'=@('Chip TPM 2.0','Nguồn');
        'SMBv1 Server'=@('SMBv1 (Server / Client)','Server');
        'SMBv1 Client'=@('SMBv1 (Server / Client)','Client');
        'Autorun Registry HKLM'=@('Autorun / Autoplay Policy','HKLM');
        'Autorun Registry HKCU'=@('Autorun / Autoplay Policy','HKCU');
        'AutoPlay UI (HKCU)'=@('Autorun / Autoplay Policy','AutoPlay UI');
        'AutoPlay - Tài khoản kiểm tra'=@('Autorun / Autoplay Policy','Tài khoản');
        'Autorun - Chính sách chi tiết'=@('Autorun / Autoplay Policy','Chính sách chi tiết');
        'NTP W32Time Service'=@('NTP / W32Time','Dịch vụ');
        'NTP Client Type'=@('NTP / W32Time','Client type');
        'NTP Server Config'=@('NTP / W32Time','Server config');
        'NTP Source'=@('NTP / W32Time','Nguồn đồng bộ');
        'NTP Last Successful Sync'=@('NTP / W32Time','Lần đồng bộ thành công');
        'NTP Notes'=@('NTP / W32Time','Ghi chú');
        'HotFix KB'=@('Cập nhật HotFix gần nhất','KB');
        'HotFix InstalledOn'=@('Cập nhật HotFix gần nhất','Ngày cài');
        'HotFix 3 gần nhất'=@('Cập nhật HotFix gần nhất','3 bản gần nhất');
        'HotFix Count'=@('Cập nhật HotFix gần nhất','Số bản ghi');
        'HotFix Source'=@('Cập nhật HotFix gần nhất','Nguồn');
        'ScreenSaveActive'=@('Khóa màn hình / Screen Saver','Bật Screen Saver');
        'ScreenSaverIsSecure'=@('Khóa màn hình / Screen Saver','Yêu cầu đăng nhập');
        'ScreenSaveTimeOut'=@('Khóa màn hình / Screen Saver','Chờ (giây)');
        'Machine Inactivity Limit Seconds'=@('Khóa màn hình / Screen Saver','Giới hạn khóa máy (giây)');
        'ScreenLock Registry User'=@('Khóa màn hình / Screen Saver','Tài khoản được kiểm tra')
    }
    if($map.ContainsKey($oldName)) {return ,$map[$oldName]}
    return $null
}
function Update-GroupedReportLine([object]$record,[string]$header,[string]$label,[string]$newValue) {
    if(-not $record.PSObject.Properties[$header]) {return}
    $source=Raw $record $header
    $chunks=@()
    if($source) {$chunks=@($source -split ' \| ')}
    # A user-edited legacy value may contain list separators or newlines.
    # Keep it inside ONE grouped item rather than creating accidental extra lines.
    $newValue=[regex]::Replace($newValue,'[\r\n]+','; ').Replace(' | ','; ')
    $prefix='- '+$label+':'
    $found=$false
    for($k=0;$k -lt $chunks.Count;$k++) {
        if(([string]$chunks[$k]).StartsWith($prefix)) {
            $chunks[$k]=$prefix+' '+$newValue
            $found=$true
            break
        }
    }
    if(-not $found) {$chunks+=($prefix+' '+$newValue)}
    $record.PSObject.Properties[$header].Value=($chunks -join ' | ')
}
function Apply-GroupedOldColumnEdits([object]$old,[object]$detail,[object]$record) {
    foreach($prop in $old.PSObject.Properties) {
        $oldName=[string]$prop.Name
        $target=Get-GroupedReportTarget $oldName
        if($null -eq $target) {continue}
        $oldValue=[string]$prop.Value
        $original=Raw $detail $oldName
        if($oldValue -eq $original) {continue}
        # This can be an intentional user edit, including clearing the cell.
        # Never overwrite the original snapshot to preserve the audit trail.
        Update-GroupedReportLine $record ([string]$target[0]) ([string]$target[1]) $oldValue
    }
}
function Is-OldAutoSummary([string]$header,[string]$oldValue,[object]$detail) {
    $oldSources=@{
        'Phiên bản Windows OS Build'='Windows OS - Version/Build';
        'Antivirus đang dùng'='Antivirus đang dùng';
        'Chip TPM 2.0'='Chip TPM 2.0';
        'SMBv1 (Server / Client)'='SMBv1 trạng thái cấu hình';
        'Autorun / Autoplay Policy'='Autorun / Autoplay Policy';
        'NTP / W32Time'='NTP W32Time';
        'Cập nhật HotFix gần nhất'='Windows HotFix gần nhất';
        'Khóa màn hình / Screen Saver'='Screen Saver Lock Status';
        'Kích hoạt Windows'='Windows kích hoạt';
        'Tường lửa Windows (Domain/Private/Public)'='Windows Firewall Status';
        'Tài khoản Guest'='Guest Account Status'
    }
    if(-not $oldSources.ContainsKey($header)) {return $false}
    $source=Raw $detail ([string]$oldSources[$header])
    if(-not $source) {$source='CHUA_QUET_BAN_MOI'}
    if($oldValue -eq $source -or $oldValue -eq ('- Trạng thái: '+$source)) {return $true}
    return $false
}

function Merge-AuditManualEdits([object[]]$records,[object[]]$details,[string]$outFolder) {
    $keys=@()
    for($i=0;$i -lt $records.Count;$i++) {
        $key=Raw $details[$i] '__AUDIT_ROW_KEY__'
        if(-not $key) {throw 'Thieu ID luot quet. Dung xuat de tranh ghep nham du lieu.'}
        $keys+=,$key
    }
    $history=Join-Path $outFolder 'ManualEdits.csv'
    $excel=Join-Path $outFolder 'Inventory.xlsx'
    $previous=@()
    if(Test-Path -LiteralPath $history) {
        try {$previous=@(Read-Utf8Csv -FilePath $history)}
        catch {throw ('Khong doc duoc ManualEdits.csv. Dung de tranh mat du lieu: '+$_.Exception.Message)}
    }
    # Excel takes priority over the last saved manual-edit backup.
    if(Test-Path -LiteralPath $excel) {
        try {$previous+=@(Read-AuditWorkbook $excel)}
        catch {throw ('Khong the doc Excel hien tai. KHONG GHI DE: '+$_.Exception.Message)}
    }
    $headers=@($records[0].PSObject.Properties | ForEach-Object {$_.Name})
    foreach($old in $previous) {
        foreach($p in $old.PSObject.Properties) {
            $h=[string]$p.Name
            if($h -eq '__AUDIT_ROW_KEY__' -or $h -eq 'STT' -or -not $h.Trim()) {continue}
            if(Exclude-ReportField $h) {continue}
            if($headers -notcontains $h) {$headers+=,$h}
        }
    }
    # Every generated row has the same schema, including any manually added columns.
    foreach($r in $records) {
        foreach($h in $headers) {if(-not $r.PSObject.Properties[$h]) {AddReportProperty $r $h ''}}
    }
    $iOld=0
    foreach($old in $previous) {
        $key=Raw $old '__AUDIT_ROW_KEY__'
        $i=-1
        if($key) {
            for($j=0;$j -lt $keys.Count;$j++){if($keys[$j] -eq $key){$i=$j;break}}
        } else { $i=Get-AuditMatch $old $records $keys $iOld }
        if($i -ge 0) {
            foreach($h in $headers) {
                if($h -eq 'STT') {continue}
                $p=$old.PSObject.Properties[$h]
                if($null -ne $p) {
                    # Raw evidence was written into historical Excel as extra
                    # columns. Some older generated rows have EMPTY extra cells
                    # even though DuLieuGoc contains valid readings. Empty
                    # historical cells are not proof of an intentional edit.
                    # Restore the original raw value in that case; preserve
                    # non-empty manually edited values as before.
                    if(@('UUID','Hãng mainboard','Serial mainboard','CPU',
                        'RAM (GB)','Ổ lưu trữ (model - GB)','Nguồn dữ liệu Antivirus',
                        'Ngày mẫu virus Defender (tham khảo)','DNS per Adapter',
                        'Gateway','Mục cần đối chiếu','Trường còn thiếu') -contains $h) {
                        if(-not ([string]$p.Value).Trim() -and (Raw $details[$i] $h)) {continue}
                    }
                    # Older workbooks stored the domain type alone. Do not let
                    # that generated value replace the new grouped report cell.
                    if($h -eq 'Domain / Workgroup') {
                        $previousType=Raw $details[$i] 'Domain / Workgroup'
                        $oldText=[string]$p.Value
                        if($oldText -eq $previousType) {continue}
                        if($oldText -notmatch '^\s*-\s*Loai:') {
                            Update-GroupedReportLine $records[$i] $h 'Loai' $oldText
                            continue
                        }
                    }
                    # Existing Excel in older formats held only a one-line status
                    # for these three fields. Rebuild the full grouped cell unless
                    # the operator actually edited that one-line status.
                    if(@('Kích hoạt Windows','Tường lửa Windows (Domain/Private/Public)','Tài khoản Guest') -contains $h) {
                        $statusSourceName=''
                        if($h -eq 'Kích hoạt Windows') {$statusSourceName='Windows kích hoạt'}
                        elseif($h -eq 'Tường lửa Windows (Domain/Private/Public)') {$statusSourceName='Windows Firewall Status'}
                        else {$statusSourceName='Guest Account Status'}
                        $oldText=[string]$p.Value
                        $sourceStatus=Raw $details[$i] $statusSourceName
                        if($oldText -eq $sourceStatus -or $oldText -eq ('- Trạng thái: '+$sourceStatus)) {continue}
                        if($oldText -notmatch '^\s*-\s*Trạng thái:') {
                            Update-GroupedReportLine $records[$i] $h 'Trạng thái' $oldText
                            continue
                        }
                    }
                    $legacySource=''
                    if($h -eq 'Mã hóa BitLocker') {
                        $legacySource=Raw $details[$i] 'Mã hóa BitLocker (ổ Windows)'
                        if($legacySource -eq 'OFF / SUSPENDED / PARTIAL'){$legacySource='CAN_XAC_MINH - KET_QUA_V5_CU'}
                    } elseif($h -eq 'Chính sách USB Storage') {
                        $legacySource=Raw $details[$i] 'USB Storage Policy'
                    }
                    # Old generated Excel values match the original status exactly;
                    # recompute those two columns so historical scans gain details.
                    # A different user-entered value is preserved as a manual edit.
                    if($legacySource -and ([string]$p.Value) -eq $legacySource){continue}
                    if(Is-OldAutoSummary $h ([string]$p.Value) $details[$i]){continue}
                    $records[$i].PSObject.Properties[$h].Value=[string]$p.Value
                }
            }
            Apply-GroupedOldColumnEdits $old $details[$i] $records[$i]
        } elseif($key -like 'MANUAL_*' -or -not $key) {
            # Preserve a row that the user manually added in Excel, including on future rebuilds.
            # Unmatched legacy rows are retained rather than silently discarded.
            $manualKey=$key
            if(-not $manualKey) {$manualKey='MANUAL_'+[guid]::NewGuid().ToString('N')}
            $new=New-Object PSObject
            foreach($h in $headers) {AddReportProperty $new $h (Raw $old $h)}
            $records+=,$new; $keys+=,$manualKey
        } else {
            throw ('Excel co ma luot quet khong con trong DuLieuGoc: '+$key+'. Dung xuat de tranh mat du lieu.')
        }
        $iOld++
    }
    for($j=0;$j -lt $records.Count;$j++){$records[$j].PSObject.Properties['STT'].Value=[string]($j+1)}
    # Repair stable property order when additional columns were found in an edited file.
    return (New-Object PSObject -Property @{Records=@($records);Keys=[string[]]@($keys)})
}
function Save-AuditManualEdits([object[]]$records,[string[]]$keys,[string]$outFolder) {
    if($records.Count -ne $keys.Count) {throw 'Cannot save manual edits: row count mismatch.'}
    $save=@()
    for($i=0;$i -lt $records.Count;$i++) {
        $copy=New-Object PSObject
        foreach($prop in $records[$i].PSObject.Properties) {AddReportProperty $copy $prop.Name $prop.Value}
        AddReportProperty $copy '__AUDIT_ROW_KEY__' $keys[$i]
        $save+=,$copy
    }
    $path=Join-Path $outFolder 'ManualEdits.csv'
    $staging=$path+'.tmp'
    Write-Utf8Csv -Records $save -FilePath $staging
    if(Test-Path -LiteralPath $path) {
        # Backup previous manual changes before replacing them.
        $dir=Join-Path $outFolder 'Backup'
        if(-not (Test-Path -LiteralPath $dir)){[void](New-Item -ItemType Directory -Path $dir -Force)}
        $backup=Join-Path $dir ('ManualEdits_'+(Get-Date -Format 'yyyyMMdd_HHmmss_fff')+'.csv')
        Copy-Item -LiteralPath $path -Destination $backup -ErrorAction Stop
        $oldBackups=@(Get-ChildItem -LiteralPath $dir -Filter 'ManualEdits_*.csv' | Sort-Object -Property Name -Descending)
        if($oldBackups.Count -gt 10) {for($j=10;$j -lt $oldBackups.Count;$j++){try {$oldBackups[$j].Delete()}catch{}}}
    }
    [System.IO.File]::Copy($staging,$path,$true)
    [System.IO.File]::Delete($staging)
}
function Assert-AuditWorkbookWritable([string]$path) {
    if(-not [System.IO.File]::Exists($path)){return}
    try {
        $fs=[System.IO.File]::Open($path,[System.IO.FileMode]::Open,[System.IO.FileAccess]::ReadWrite,[System.IO.FileShare]::None)
        $fs.Close()
    } catch {throw 'Inventory.xlsx dang mo/bi khoa. Hay dong Excel roi chay Rebuild. Khong ghi de.'}
}
function Commit-AuditWorkbook([string]$tempPath,[string]$targetPath) {
    Assert-AuditWorkbookWritable $targetPath
    $dir=Join-Path (Split-Path -Parent $targetPath) 'Backup'
    if(-not (Test-Path -LiteralPath $dir)){[void](New-Item -ItemType Directory -Path $dir -Force)}
    $backup=Join-Path $dir ('Inventory_'+(Get-Date -Format 'yyyyMMdd_HHmmss_fff')+'.xlsx')
    if([System.IO.File]::Exists($targetPath)) {
        # File.Replace is atomic on supporting filesystems; fallback is a backed-up copy.
        try {[System.IO.File]::Replace($tempPath,$targetPath,$backup)}
        catch {
            Copy-Item -LiteralPath $targetPath -Destination $backup -ErrorAction Stop
            try {[System.IO.File]::Copy($tempPath,$targetPath,$true)}
            catch {
                Copy-Item -LiteralPath $backup -Destination $targetPath -Force -ErrorAction SilentlyContinue
                throw
            }
            [System.IO.File]::Delete($tempPath)
        }
    } else {[System.IO.File]::Move($tempPath,$targetPath)}
    # Cap Excel backup count to 10 to avoid filling a USB during multi-floor surveys.
    $past=@(Get-ChildItem -LiteralPath $dir -Filter 'Inventory_*.xlsx' | Sort-Object -Property Name -Descending)
    if($past.Count -gt 10) {for($i=10;$i -lt $past.Count;$i++){try {$past[$i].Delete()}catch{}}}
}
