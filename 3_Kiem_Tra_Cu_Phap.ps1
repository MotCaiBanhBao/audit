# Optional syntax check. Does not collect data or modify activation.
$root=Split-Path -Parent $MyInvocation.MyCommand.Path
$failed=0
foreach($file in @(Get-ChildItem -LiteralPath $root -Filter '*.ps1')){
    $err=@()
    try{
        if($PSVersionTable.PSVersion.Major -ge 3){
            $tok=$null
            [void][System.Management.Automation.Language.Parser]::ParseFile($file.FullName,[ref]$tok,[ref]$err)
        }else{
            $code=[System.IO.File]::ReadAllText($file.FullName)
            [void][System.Management.Automation.PSParser]::Tokenize($code,[ref]$err)
        }
        if(@($err).Count -gt 0){
            $failed++
            Write-Host ('[LOI] '+$file.Name) -ForegroundColor Red
            foreach($one in $err){Write-Host ('  '+$one.Message) -ForegroundColor Red}
        }else{Write-Host ('[OK] '+$file.Name) -ForegroundColor Green}
    }catch{
        $failed++
        Write-Host ('[LOI] '+$file.Name+': '+$_.Exception.Message) -ForegroundColor Red
    }
}
if($failed -gt 0){Write-Host ('Co '+$failed+' file can kiem tra.') -ForegroundColor Red;exit 1}
Write-Host 'Khong phat hien loi cu phap qua bo phan tich PowerShell.' -ForegroundColor Green
exit 0
