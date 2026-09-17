param([ValidateSet('variable','regular')][string]$Font='variable')
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$root=Join-Path $workspace '.local/word-default-fonts'
$path=[IO.Path]::GetFullPath((Join-Path $root "Inter-$Font.ttf"))
if(-not $path.StartsWith($workspace+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Font path leaves workspace'}
function Hash($file){(Get-FileHash -LiteralPath $file).Hash.ToLowerInvariant()}
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class NofficeTemporaryFont {
 [DllImport("gdi32.dll",EntryPoint="AddFontResourceExW",CharSet=CharSet.Unicode,SetLastError=true)]
 public static extern int Add(string path,uint flags,IntPtr reserved);
 [DllImport("gdi32.dll",EntryPoint="RemoveFontResourceExW",CharSet=CharSet.Unicode,SetLastError=true)]
 public static extern bool Remove(string path,uint flags,IntPtr reserved);
}
'@
$app=$null;$doc=$null;$idle=$false;$added=$false
try{
 $count=[NofficeTemporaryFont]::Add($path,0,[IntPtr]::Zero)
 if($count -le 0){throw 'Could not temporarily load the authored font'};$added=$true
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Requires idle owned Word'};$idle=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.Options.UpdateLinksAtOpen
 $app.Visible=$false;$app.AutomationSecurity=3;$app.DisplayAlerts=0;$app.Options.UpdateLinksAtOpen=$false
 $exe=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 $baseline=(Get-Content (Join-Path $workspace 'docs/parity/baseline.json') -Raw|ConvertFrom-Json).applications|Where-Object executable -eq 'WINWORD.EXE'
 if($exe -ne $baseline.sha256){throw 'Native Word baseline changed'}
 $names=@($app.FontNames|ForEach-Object {[string]$_})
 if('Inter' -notin $names){throw 'Word did not enumerate Inter'}

 $folder=Join-Path $root 'calibration';[IO.Directory]::CreateDirectory($folder)|Out-Null
 $rows=@()
 foreach($size in @(8,10,12,14,18,24)){
  foreach($rule in @('single','onehalf','default','double','minimum','exact')){
   $doc=$app.Documents.Add();$doc.Activate();$doc.ActiveWindow.View.ReadingLayout=$false;$doc.ActiveWindow.View.Type=3
   $setup=$doc.PageSetup;$setup.PageWidth=595.3;$setup.PageHeight=841.9
   $setup.TopMargin=72;$setup.BottomMargin=72;$setup.LeftMargin=72;$setup.RightMargin=72
   $doc.Content.Text=(1..12|ForEach-Object {'sample'+([string]$_).PadLeft(2,'0')}) -join [char]11
   $doc.Content.Font.Name='Inter';$doc.Content.Font.Size=$size;$doc.Content.Font.Kerning=0
   $format=$doc.Content.ParagraphFormat;$format.SpaceBefore=0;$format.SpaceAfter=0
   $format.SpaceBeforeAuto=0;$format.SpaceAfterAuto=0;$format.KeepTogether=0;$format.KeepWithNext=0;$format.WidowControl=-1
   switch($rule){
    'single' {$format.LineSpacingRule=0}
    'onehalf' {$format.LineSpacingRule=1}
    'double' {$format.LineSpacingRule=2}
    'default' {$format.LineSpacingRule=5;$format.LineSpacing=19.8}
    'minimum' {$format.LineSpacingRule=3;$format.LineSpacing=18}
    'exact' {$format.LineSpacingRule=4;$format.LineSpacing=18}
   }
   $name="$size-$rule";$source=[string](Join-Path $folder ($name+'.docx'))
   $doc.SaveAs2([ref]$source,[ref]12);$doc.Close([ref]0)
   $doc=$app.Documents.Open($source,$false,$true,$false)
   $pdf=[string](Join-Path $folder ($name+'.pdf'));$doc.ExportAsFixedFormat($pdf,17)
   $rows+=@{name=$name;size=$size;rule=$rule;sourceHash=Hash $source;pdfHash=Hash $pdf;nativeLine=[double]$doc.Content.ParagraphFormat.LineSpacing}
   $doc.Close([ref]0);$doc=$null
   Write-Output "Captured Inter $name"
  }
 }
 @{cases=$rows;fontHash=Hash $path;scriptHash=Hash $PSCommandPath;executableHash=$exe}|ConvertTo-Json -Depth 10|Set-Content (Join-Path $folder 'native-report.json') -Encoding UTF8
}finally{
 if($doc){$doc.Close([ref]0)}
 if($app){if($idle){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 if($added){if(-not [NofficeTemporaryFont]::Remove($path,0,[IntPtr]::Zero)){Write-Error 'Temporary font removal failed'}}
}
