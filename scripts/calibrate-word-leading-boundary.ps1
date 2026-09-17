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

 $folder=Join-Path $root 'leading-boundary';[IO.Directory]::CreateDirectory($folder)|Out-Null
 $rows=@()
 foreach($name in @('plain-six','widow-off','widow-on','keep-paragraph','keep-next')){
  $doc=$app.Documents.Add();$doc.Activate();$doc.ActiveWindow.View.ReadingLayout=$false;$doc.ActiveWindow.View.Type=3
  $setup=$doc.PageSetup;$setup.PageWidth=400;$setup.PageHeight=200
  $setup.TopMargin=20;$setup.BottomMargin=20;$setup.LeftMargin=20;$setup.RightMargin=20
  $count=if($name -in @('widow-off','widow-on')){7}else{6}
  $parts=@(1..$count|ForEach-Object {'item'+([string]$_).PadLeft(2,'0')})
  $text=$parts -join [char]11
  if($name -eq 'keep-paragraph'){$text=($parts[0..3] -join [char]11)+[char]13+($parts[4..5] -join [char]11)}
  if($name -eq 'keep-next'){$text=($parts[0..4] -join [char]11)+[char]13+$parts[5]}
  $doc.Content.Text=$text;$doc.Content.Font.Name='Inter';$doc.Content.Font.Size=12;$doc.Content.Font.Kerning=0
  $format=$doc.Content.ParagraphFormat;$format.SpaceBefore=0;$format.SpaceAfter=0
  $format.SpaceBeforeAuto=0;$format.SpaceAfterAuto=0;$format.KeepTogether=0;$format.KeepWithNext=0;$format.WidowControl=-1;$format.LineSpacingRule=2
  if($name -eq 'widow-off'){$format.WidowControl=0}
  if($name -eq 'keep-paragraph'){$doc.Paragraphs.Item(2).Format.KeepTogether=-1}
  if($name -eq 'keep-next'){$doc.Paragraphs.Item(1).Format.KeepWithNext=-1}
  $source=[string](Join-Path $folder ($name+'.docx'));$doc.SaveAs2([ref]$source,[ref]12);$doc.Close([ref]0)
  $doc=$app.Documents.Open($source,$false,$true,$false);$doc.ActiveWindow.View.ReadingLayout=$false;$doc.ActiveWindow.View.Type=3;$doc.Repaginate();$pdf=[string](Join-Path $folder ($name+'.pdf'));$doc.ExportAsFixedFormat($pdf,17)
  $characters=@()
  for($i=0;$i -lt [int]$doc.Content.End-1;$i++){$c=$doc.Range($i,$i+1);$characters+=@{text=[string]$c.Text;page=[int]$c.Information(3)}}
  $rows+=@{name=$name;sourceHash=Hash $source;pdfHash=Hash $pdf;pages=[int]$doc.ComputeStatistics(2);characters=$characters}
  $doc.Close([ref]0);$doc=$null
  Write-Output ('Captured automatic-leading boundary '+$name)
 }
 @{cases=$rows;fontHash=Hash $path;scriptHash=Hash $PSCommandPath;executableHash=$exe}|ConvertTo-Json -Depth 15|Set-Content (Join-Path $folder 'native-report.json') -Encoding UTF8
}finally{
 if($doc){$doc.Close([ref]0)}
 if($app){if($idle){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 if($added){if(-not [NofficeTemporaryFont]::Remove($path,0,[IntPtr]::Zero)){Write-Error 'Temporary font removal failed'}}
}
