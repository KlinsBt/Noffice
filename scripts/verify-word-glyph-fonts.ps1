Add-Type @'
using System; using System.Runtime.InteropServices;
public class GlyphGdi {
 [DllImport("gdi32.dll")] public static extern IntPtr CreateCompatibleDC(IntPtr h);
 [DllImport("gdi32.dll",CharSet=CharSet.Unicode)] public static extern IntPtr CreateFontW(int h,int w,int e,int o,int weight,uint italic,uint underline,uint strike,uint charset,uint output,uint clip,uint quality,uint pitch,string face);
 [DllImport("gdi32.dll")] public static extern IntPtr SelectObject(IntPtr h,IntPtr o);
 [DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr o);
 [DllImport("gdi32.dll")] public static extern bool DeleteDC(IntPtr h);
 [DllImport("gdi32.dll")] public static extern bool GetTextMetricsW(IntPtr h,IntPtr data);
 public static int[] Measure(int size) {
  var dc=CreateCompatibleDC(IntPtr.Zero);var font=CreateFontW(-size,0,0,0,400,0,0,0,1,0,0,0,0,"Arial");var old=SelectObject(dc,font);var memory=Marshal.AllocHGlobal(80);
  try {if(!GetTextMetricsW(dc,memory))throw new Exception("No native font metrics");var values=new int[5];Marshal.Copy(memory,values,0,5);return values;}
  finally {Marshal.FreeHGlobal(memory);SelectObject(dc,old);DeleteObject(font);DeleteDC(dc);}
 }
}
'@
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root '.local/word-glyph-fonts'
[void][IO.Directory]::CreateDirectory($directory)
$app=$null;$doc=$null;$idle=$false
try {
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Requires an idle owned Word instance.'};$idle=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.Options.UpdateLinksAtOpen
 $app.Visible=$false;$app.AutomationSecurity=3;$app.DisplayAlerts=0;$app.Options.UpdateLinksAtOpen=$false
 $native=(Get-Content docs/parity/baseline.json -Raw|ConvertFrom-Json).applications|Where-Object executable -eq 'WINWORD.EXE'
 $exe=(Get-FileHash (Join-Path ([string]$app.Path) 'WINWORD.EXE')).Hash.ToLowerInvariant()
 if($exe -ne $native.sha256){throw 'Word baseline changed.'}
 @{security=$security;alerts=$alerts;links=$links}|ConvertTo-Json|Set-Content (Join-Path $directory 'preferences.json') -Encoding UTF8
 $source=[string](Join-Path $root 'tests/fixtures/word-glyph-fonts.docx');$saved=[string](Join-Path $directory 'saved.docx');$pdf=[string](Join-Path $directory 'native.pdf')
 Write-Output 'Opening source'
 $doc=$app.Documents.Open($source,$false,$false,$false)
 Write-Output 'Saving native copy'
 $doc.SaveAs2([ref]$saved,[ref]12);$doc.Close([ref]0);$doc=$null
 Write-Output 'Reopening native copy'
 $doc=$app.Documents.Open($saved,$false,$true,$false)
 Write-Output 'Rendering native copy'
 $doc.Repaginate();$doc.ExportAsFixedFormat($pdf,17)
 $gdi=@();foreach($size in @(10,15,20,25,30,35,40)){$pixels=[int][Math]::Round($size*600/72);$m=[GlyphGdi]::Measure($pixels);$gdi+=@{size=$size;dpi=600;pixels=$pixels;height=$m[0];ascent=$m[1];descent=$m[2];internal=$m[3];external=$m[4]}}
 @{executableSha256=$exe;sourceSha256=(Get-FileHash $source).Hash.ToLowerInvariant();scriptSha256=(Get-FileHash $PSCommandPath).Hash.ToLowerInvariant();pdfSha256=(Get-FileHash $pdf).Hash.ToLowerInvariant();savedSha256=(Get-FileHash $saved).Hash.ToLowerInvariant();fontSha256=(Get-FileHash (Join-Path $env:WINDIR 'Fonts/arial.ttf')).Hash.ToLowerInvariant();printer=[string]$app.ActivePrinter;gdi600=$gdi}|ConvertTo-Json -Depth 6|Set-Content (Join-Path $directory 'report.json') -Encoding UTF8
} finally {
 if($doc){$doc.Close([ref]0)}
 if($app){if($idle){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
