Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class BaselineGdi {
 [StructLayout(LayoutKind.Sequential,CharSet=CharSet.Unicode)] public struct TM {public int height,ascent,descent,internalLeading,externalLeading,ave,max,weight,overhang,aspectX,aspectY;public char first,last,def,brk;public byte italic,under,strike,pitch,charset;}
 [DllImport("gdi32.dll",CharSet=CharSet.Unicode)] public static extern IntPtr CreateDC(string driver,string device,string output,IntPtr init);
 [DllImport("gdi32.dll",CharSet=CharSet.Unicode)] public static extern int GetTextFace(IntPtr dc,int count,System.Text.StringBuilder face);
 [DllImport("gdi32.dll")] public static extern uint GetFontData(IntPtr dc,uint table,uint offset,byte[] buffer,uint count);
 [DllImport("gdi32.dll")] public static extern bool DeleteDC(IntPtr dc);
 [DllImport("gdi32.dll")] public static extern int GetDeviceCaps(IntPtr dc,int index);
 [DllImport("gdi32.dll",CharSet=CharSet.Unicode)] public static extern IntPtr CreateFont(int h,int w,int e,int o,int weight,uint italic,uint underline,uint strike,uint charset,uint output,uint clip,uint quality,uint pitch,string face);
 [DllImport("gdi32.dll")] public static extern IntPtr SelectObject(IntPtr dc,IntPtr obj);
 [DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr obj);
 [DllImport("gdi32.dll",CharSet=CharSet.Unicode)] public static extern bool GetTextMetrics(IntPtr dc,out TM result);
}
"@
$dc=[BaselineGdi]::CreateDC('WINSPOOL','Microsoft Print to PDF',$null,[IntPtr]::Zero)
if($dc -eq [IntPtr]::Zero){throw 'No printer metrics DC'}
try {
 $dpi=[BaselineGdi]::GetDeviceCaps($dc,90);$rows=@()
 foreach($size in @(10,20,30)){
  $height=[Math]::Round($size*$dpi/72)
  $font=[BaselineGdi]::CreateFont(-$height,0,0,0,400,0,0,0,1,0,0,0,0,'Arial');$prior=[BaselineGdi]::SelectObject($dc,$font)
  try { $m=New-Object BaselineGdi+TM;[void][BaselineGdi]::GetTextMetrics($dc,[ref]$m);$face=New-Object Text.StringBuilder 256;[void][BaselineGdi]::GetTextFace($dc,256,$face)
  $length=[BaselineGdi]::GetFontData($dc,0,0,$null,0);$fontHash=$null
  if($length -ne [uint32]::MaxValue -and $length -lt 16777216){$bytes=New-Object byte[] $length;[void][BaselineGdi]::GetFontData($dc,0,0,$bytes,$length);$hasher=[Security.Cryptography.SHA256]::Create();try{$fontHash=([BitConverter]::ToString($hasher.ComputeHash($bytes))).Replace('-','').ToLowerInvariant()}finally{$hasher.Dispose()}}
  $rows+=@{size=$size;dpi=$dpi;selectedHeight=$height;metrics=$m;face=$face.ToString();fontSha256=$fontHash} }
  finally{[void][BaselineGdi]::SelectObject($dc,$prior);[void][BaselineGdi]::DeleteObject($font)}
 }
 $rows|ConvertTo-Json -Depth 5|Set-Content -Encoding UTF8 .local/word-baselines/gdi-metrics.json
}finally{[void][BaselineGdi]::DeleteDC($dc)}
