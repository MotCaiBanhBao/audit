# ReviewReasons.ps1 - explain why each item needs verification, based on
# the observed values. This is documentation, not an automated compliance verdict.
# Windows PowerShell 2.0+; no external calls or modification of machine state.
function Audit-Short([object]$value,[int]$max=320) {
    $s=([string]$value).Trim() -replace '[\r\n\t]+',' '
    if(-not $s){return '(khong co ket qua)'}
    if($s.Length -gt $max){return ($s.Substring(0,$max)+'... (xem DuLieuGoc/LenhGoc)')}
    return $s
}
function Get-AuditReviewReasons([string[]]$reasons,[string[]]$missing,[string[]]$errors,
    [object]$lic,[object]$security,[object]$extra,[object]$compliance,[object]$row) {
    $out=@()
    foreach($reason in @($reasons)) {
        if(-not $reason){continue}
        $detail=''
        switch -Regex ([string]$reason) {
            '^Thieu thong tin nhan dang$' {
                $detail='Truong chua doc duoc hoac SMBIOS/nguoi khao sat bo trong: '+(Audit-Short (@($missing) -join ', ') 460)+'. Kiem tra nhan may, BIOS va thong tin nguoi su dung.';break
            }
            '^Loi truy van$' {
                $detail='Co loi khi lay du lieu; '+(Audit-Short (@($errors | Select-Object -First 4) -join '; ') 520)+'. Xem KhaoSat\\Log va KhaoSat\\LenhGoc de biet lenh nao loi/timeout.';break
            }
            '^Tien trinh kiem tra song song loi$' {
                $detail='Mot worker khong hoan tat hoac timeout; ket qua cac nhom khac co the van hop le. Xem Error/TIMOUT trong log va LenhGoc.';break
            }
            '^Phan mem nghi van' {
                $detail='Trang thai quet: '+(Audit-Short $security.Suspicious)+'. Neu quet chua du, kiem tra danh sach ung dung xuat kem; neu phat hien can doi chieu ten/nguon cai dat.';break
            }
            '^Trang thai antivirus' {
                $detail='Real-time: '+(Audit-Short $security.AVRealTime)+'; Nguon: '+(Audit-Short $security.AVSource 180)+'. Neu CAN_XAC_MINH_TRONG_UNG_DUNG, phai xem giao dien/phien ban AV thuc te; trang thai khong phai bang chung virus.';break
            }
            '^Ngay cap nhat antivirus$' {
                $detail='Moc cap nhat nhan duoc: '+(Audit-Short $security.AVUpdated)+'. Cong cu khong doc duoc ngay cap nhat dang tin cay; xem lich su signature trong ung dung AV.';break
            }
            '^TPM ' {
                $detail='TPM doc duoc: '+(Audit-Short $security.TPM)+'; phien ban: '+(Audit-Short $security.TPMSpec)+'; nguon: '+(Audit-Short $security.TPMSource 120)+'. Co the khong ho tro TPM 2.0 hoac truy van khong du; doi chieu tpm.msc/BIOS.';break
            }
            '^BitLocker ' {
                $detail='Trang thai doc duoc: '+(Audit-Short $security.BitLocker)+'; nguon: '+(Audit-Short $security.BitLockerSource 150)+'. Neu OFF thi la trang thai bao ve hien tai, khong nhat thiet loi truy van; doi chieu chinh sach don vi.';break
            }
            '^Quyen Local Admin$' {
                $detail='Gia tri: '+(Audit-Short $security.Admin)+'. Chua du bang chung ve nhom quyen local; xem Local Users and Groups/Administrators.';break
            }
            '^Ban quyen Windows$' {
                $method=Audit-Short $lic.WindowsMethod 140
                $note=Audit-Short $lic.WindowsNote 560
                $detail='Ket qua: '+(Audit-Short $lic.WindowsState)+'; nguon: '+$method+'; chi tiet: '+$note+'. Neu CAN_XAC_MINH: doi chieu Settings > Activation hoac slmgr /xpr, sau do nhap tay trong tool.';break
            }
            '^Ban quyen Office$' {
                $detail='Office: '+(Audit-Short $lic.OfficeState)+'; nguon: '+(Audit-Short $lic.OfficeMethod 150)+'; chi tiet: '+(Audit-Short $lic.OfficeNote 440)+'. Doi chieu trang thai trong File > Account cua Office.';break
            }
            '^SMBv1' {
                $detail='Server: '+(Audit-Short $compliance.SMB1Server 170)+'; Client: '+(Audit-Short $compliance.SMB1Client 170)+'. Can doi chieu cau hinh ca 2 thanh phan; ENABLING khong phai loi truy van.';break
            }
            '^Autorun' {
                $detail='Gia tri: '+(Audit-Short $compliance.AutoRun 200)+'; chi tiet: '+(Audit-Short $compliance.AutoRunNote 260)+'. Khong xac nhan day du chinh sach tren ca HKLM/HKCU/GPO.';break
            }
            '^NTP ' {
                $detail='Nguon: '+(Audit-Short $compliance.NTPSource 170)+'; lan dong bo: '+(Audit-Short $compliance.NTPLastSync 170)+'. Khong co du thong tin xac nhan dong bo thoi gian.';break
            }
            '^Domain ' {
                $detail='Trang thai: '+(Audit-Short $compliance.DomainStatus)+'; ten: '+(Audit-Short $compliance.DomainName)+'. Can doi chieu domain/workgroup thuc te do truy van chua xac dinh.';break
            }
            '^IP DHCP' {
                $detail='Trang thai: '+(Audit-Short $compliance.IPMode)+'. Chua doc duoc DHCP Enabled tu card mang hien hoat.';break
            }
            '^Windows Firewall$' {
                $detail='Ket qua: '+(Audit-Short $extra.Firewall)+'. Neu OFF: can danh gia theo chinh sach bao mat; neu CAN_XAC_MINH: chua doc duoc profile day du.';break
            }
            '^Windows HotFix$' {
                $detail='Ket qua: '+(Audit-Short $extra.Hotfix)+'; nguon: '+(Audit-Short $extra.HotfixSource 140)+'. Khong tim thay hotfix qua API khong dong nghia may khong duoc cap nhat.';break
            }
            '^USB Storage' {
                $detail='Ket qua: '+(Audit-Short $extra.USB)+'. Chua doc duoc day du policy USBSTOR/GPO; can kiem tra theo yeu cau don vi.';break
            }
            '^Guest Account$' {
                $detail='Gia tri: '+(Audit-Short $extra.Guest)+'. Guest bat la trang thai can ra soat; CAN_XAC_MINH la khong du thong tin tai khoan.';break
            }
            '^DNS Server$' {
                $detail='Gia tri: '+(Audit-Short $extra.DNS)+'. Khong du thong tin DNS cua card mang hoat dong (hoac card chua cau hinh DNS).';break
            }
            '^Screen Lock' {
                $detail='Gia tri: '+(Audit-Short $extra.ScreenLock 370)+'. Chua chung minh duoc khoa man hinh theo policy may/nguoi dung, can doi chieu GPO/Windows Settings.';break
            }
            '^Screen Saver khac tai khoan' {
                $detail='HKCU duoc doc trong ngu canh tai khoan dang chay tool (co the la admin), khong nhat thiet tai khoan nguoi dung may.';break
            }
            default {$detail='Trang thai can doi chieu voi thuc te/cau hinh co quan va du lieu lenh goc.'}
        }
        $out+=($reason+': '+$detail)
    }
    # Additional explanations for inconclusive AV vendor license/status fields.
    $kas=[string]$row.'Phần mềm Kaspersky'
    if($kas -match 'CAN_XAC_MINH_LICENSE'){
        $out+=('Kaspersky: '+(Audit-Short $kas 200)+'. Phat hien ung dung, nhung tool khong xac nhan duoc han/bang phep; kiem tra tab License/Subscription trong Kaspersky.')
    }
    $avLic=[string]$row.'Antivirus - Giấy phép từng sản phẩm'
    if($avLic -match 'CAN_XAC_MINH'){
        $out+=('Ban quyen antivirus: '+(Audit-Short $avLic 350)+'. API nha san xuat khong tra du ket qua; kiem tra trang thai ngay trong ung dung antivirus.')
    }
    return ($out -join ' | ')
}
