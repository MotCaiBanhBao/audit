# Registry evidence reader: Windows PowerShell 2.0+ (Windows 7-11).
# READ-ONLY. A missing value is different from 0, disabled, or access denied.
function Read-RegistryEvidence([string]$Path,[string]$Name) {
    $result=New-Object PSObject -Property @{
        Path=$Path; Name=$Name; State='UNKNOWN'; Value=$null; Error=''
    }
    try {
        if(-not (Test-Path -Path $Path -ErrorAction Stop)) {
            $result.State='KEY_MISSING'
            return $result
        }
        $item=Get-ItemProperty -Path $Path -ErrorAction Stop
        if(Get-Command Write-RawEvidence -ErrorAction SilentlyContinue){
            Write-RawEvidence ('Registry_'+$Path+'_'+$Name) ('Get-ItemProperty -Path '+$Path) $item
        }
        $member=$item.PSObject.Properties[$Name]
        if($null -eq $member) {
            $result.State='VALUE_MISSING'
            return $result
        }
        $result.Value=$member.Value
        $result.State='FOUND'
    }catch{
        $result.State='QUERY_ERROR'
        $result.Error=[string]$_.Exception.Message
        if(Get-Command SurveyLog -ErrorAction SilentlyContinue){
            SurveyLog ('Registry ['+$Path+' / '+$Name+'] '+$result.Error) 'WARN'
        }
    }
    return $result
}
