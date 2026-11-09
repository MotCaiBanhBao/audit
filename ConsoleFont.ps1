# Optional visual improvement for classic Windows console (Win7-11).
# Changes font size for CURRENT console ONLY. Does not write to Registry.
# Windows Terminal/ConPTY may ignore this legacy API. No effect on scan/export.
function Set-AuditConsoleFont([int]$pixelHeight) {
    if($pixelHeight -lt 14 -or $pixelHeight -gt 30) { return $false }
    try {
        if(-not ('AuditConsoleFontNative' -as [type])) {
            $src=@'
using System;
using System.Runtime.InteropServices;
[StructLayout(LayoutKind.Sequential)]
public struct AuditCoord { public short X; public short Y; }
[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
public struct AuditConsoleFontInfo {
    public uint cbSize;
    public uint nFont;
    public AuditCoord dwFontSize;
    public int FontFamily;
    public int FontWeight;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst=32)]
    public string FaceName;
}
public static class AuditConsoleFontNative {
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern IntPtr GetStdHandle(int nStdHandle);
    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool GetCurrentConsoleFontEx(IntPtr hConsoleOutput, bool bMaximumWindow, ref AuditConsoleFontInfo info);
    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool SetCurrentConsoleFontEx(IntPtr hConsoleOutput, bool bMaximumWindow, ref AuditConsoleFontInfo info);
}
'@
            Add-Type -TypeDefinition $src -ErrorAction Stop | Out-Null
        }
        $handle=[AuditConsoleFontNative]::GetStdHandle(-11)
        if($handle -eq [IntPtr]::Zero -or $handle.ToInt64() -eq -1) {return $false}
        $info=New-Object AuditConsoleFontInfo
        $info.cbSize=[uint32][Runtime.InteropServices.Marshal]::SizeOf([type][AuditConsoleFontInfo])
        if(-not [AuditConsoleFontNative]::GetCurrentConsoleFontEx($handle,$false,[ref]$info)) {return $false}
        $info.dwFontSize.Y=[int16]$pixelHeight
        return [AuditConsoleFontNative]::SetCurrentConsoleFontEx($handle,$false,[ref]$info)
    } catch {return $false}
}
