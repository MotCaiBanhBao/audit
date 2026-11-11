# RawEvidence.ps1 - records complete returned objects/text BEFORE report filtering.
# Compatible with Windows PowerShell 2.0+, read-only audit operations.
# Raw evidence can contain usernames, IPs, installed product IDs and other
# sensitive device metadata. Do not share this folder publicly.
function Get-RawEvidenceFolder {
    $where=[string]$env:AUDIT_RAW_OUTPUT_DIR
    if(-not $where){return ''}
    if(-not (Test-Path -LiteralPath $where)){
        try {[void](New-Item -ItemType Directory -Path $where -Force -ErrorAction Stop)}catch{return ''}
    }
    return $where
}
function New-RawEvidenceName([string]$name) {
    $safe=[regex]::Replace($name,'[^a-zA-Z0-9_.-]','_')
    if($safe.Length -gt 90){$safe=$safe.Substring(0,90)}
    if(-not $safe){$safe='Query'}
    return $safe+'_'+(Get-Date -Format 'HHmmss_fff')+'_'+[guid]::NewGuid().ToString('N').Substring(0,6)
}
function Write-RawEvidence([string]$name,[string]$command,[object]$value,[string]$errorText='') {
    $folder=Get-RawEvidenceFolder
    if(-not $folder){return}
    try {
        $base=Join-Path $folder (New-RawEvidenceName $name)
        $meta=@('Command: '+$command,'Time: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
            'Computer: '+$env:COMPUTERNAME,'Status: '+$(if($errorText){'ERROR'}else{'OK'}),
            'Error: '+$errorText)
        [System.IO.File]::WriteAllLines($base+'.meta.txt',[string[]]$meta,(New-Object System.Text.UTF8Encoding -ArgumentList $false))
        if(-not $errorText) {
            # CLIXML preserves types and fields, including array members not used
            # by the filtered inventory. Text is human-readable Format-List *.
            $objects=@($value)
            try {
                $objects | Export-Clixml -Path ($base+'.clixml') -ErrorAction Stop
            }catch{
                [System.IO.File]::AppendAllText($base+'.meta.txt',('CLIXML_ERROR: '+$_.Exception.Message+[Environment]::NewLine))
            }
            try {
                $txt=($objects | Format-List * | Out-String -Width 4096)
                [System.IO.File]::WriteAllText($base+'.txt',[string]$txt,(New-Object System.Text.UTF8Encoding -ArgumentList $false))
            }catch{
                [System.IO.File]::AppendAllText($base+'.meta.txt',('TEXT_ERROR: '+$_.Exception.Message+[Environment]::NewLine))
            }
        }
    } catch { }
}
function Write-RawCommandText([string]$name,[string]$command,[string]$stdout,[string]$stderr,[string]$status='OK') {
    $folder=Get-RawEvidenceFolder
    if(-not $folder){return}
    try {
        $base=Join-Path $folder (New-RawEvidenceName $name)
        $utf8=New-Object System.Text.UTF8Encoding -ArgumentList $false
        [System.IO.File]::WriteAllText($base+'.stdout.txt',[string]$stdout,$utf8)
        [System.IO.File]::WriteAllText($base+'.stderr.txt',[string]$stderr,$utf8)
        [System.IO.File]::WriteAllLines($base+'.meta.txt',[string[]]@(
            'Command: '+$command,'Time: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
            'Computer: '+$env:COMPUTERNAME,'Status: '+$status),$utf8)
    } catch { }
}

# One human-readable file for every command result captured in this survey.
# Detailed CLIXML files remain available alongside the aggregate for fidelity.
function Export-RawCommandTranscript([string]$Root,[string]$OutputPath) {
    if(-not (Test-Path -LiteralPath $Root)){return $false}
    $writer=$null
    try {
        $utf8=New-Object System.Text.UTF8Encoding -ArgumentList $false
        $writer=New-Object System.IO.StreamWriter -ArgumentList @($OutputPath,$false,$utf8)
        $writer.WriteLine('DEVICE AUDIT TOOL - RAW COMMAND OUTPUT (FULL RETURN VALUES)')
        $writer.WriteLine('Computer: '+$env:COMPUTERNAME)
        $writer.WriteLine('Exported: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
        $writer.WriteLine('Source directory: '+$Root)
        $writer.WriteLine('NOTE: File contains sensitive device inventory; share only with authorized staff.')
        $entries=@(Get-ChildItem -LiteralPath $Root -Recurse -Filter '*.meta.txt' -ErrorAction SilentlyContinue | Sort-Object FullName)
        $writer.WriteLine('Captured commands: '+$entries.Count)
        $writer.WriteLine('Each CLIXML sibling preserves typed command objects (when available).')
        foreach($meta in $entries){
            $base=$meta.FullName.Substring(0,$meta.FullName.Length-('.meta.txt').Length)
            $relative=$meta.FullName.Substring($Root.Length).TrimStart('\')
            $writer.WriteLine('')
            $writer.WriteLine(('=' * 95))
            $writer.WriteLine('COMMAND EVIDENCE: '+$relative)
            $writer.WriteLine(('=' * 95))
            foreach($suffix in @('.meta.txt','.txt','.stdout.txt','.stderr.txt')){
                $source=$base+$suffix
                if(-not (Test-Path -LiteralPath $source -PathType Leaf)){continue}
                $writer.WriteLine('--- '+$suffix+' ---')
                $reader=$null
                try {
                    $reader=New-Object System.IO.StreamReader -ArgumentList @($source,$utf8,$true)
                    $line=$reader.ReadLine()
                    while($null -ne $line){$writer.WriteLine($line);$line=$reader.ReadLine()}
                } catch {$writer.WriteLine('READ_ERROR: '+$_.Exception.Message)}
                finally {if($reader){$reader.Dispose()}}
            }
        }
        $writer.Flush()
        return $true
    } catch {return $false}
    finally {if($writer){$writer.Dispose()}}
}
