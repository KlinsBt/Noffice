param([ValidateSet('variable','regular')][string]$Font='variable',[switch]$Long,[switch]$Boundary)
$ErrorActionPreference='Stop'
if($Long -and $Boundary){throw 'Choose one native matrix'}
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

 $matrix=if($Long){'word-inter-long-metrics'}elseif($Boundary){'word-leading-boundary'}else{'word-inter-metrics'}
 $oraclePath=Join-Path $workspace ('tests/fixtures/native-'+$matrix+'.json')
 $oracle=Get-Content $oraclePath -Raw|ConvertFrom-Json
 if((Hash $path) -ne $oracle.fontHash){throw 'Bundled font differs from the independent native oracle'}
 $rows=@()
 foreach($case in $oracle.cases){
  $folder=Join-Path $workspace ('.local/'+$matrix+'/'+$case.name)
  $browser=Get-Content (Join-Path $folder 'browser-report.json') -Raw|ConvertFrom-Json
  $source=[string](Join-Path $workspace ('tests/fixtures/'+$matrix+'/'+$case.name+'.docx'))
  $actual=[string](Join-Path $folder 'restored.docx')
  if((Hash $source) -ne $case.sourceHash -or (Hash $actual) -ne $browser.outputs.docx){throw 'Stale native/browser input'}
  $states=@{}
  foreach($stage in @('expected','actual')){
   $input=if($stage -eq 'expected'){$source}else{$actual}
   $doc=$app.Documents.Open($input,$false,$true,$false)
   $doc.ActiveWindow.View.ReadingLayout=$false;$doc.ActiveWindow.View.Type=3;$doc.Repaginate()
   $paragraphs=@()
   foreach($p in $doc.Paragraphs){
    $r=$p.Range;$mark=$doc.Range([int]$r.End-1,[int]$r.End);$pf=$r.ParagraphFormat
    $paragraphs+=@{text=[string]$r.Text;font=[string]$r.Font.Name;size=[double]$r.Font.Size;markFont=[string]$mark.Font.Name;markSize=[double]$mark.Font.Size;lineRule=[int]$pf.LineSpacingRule;line=[double]$pf.LineSpacing;before=[double]$pf.SpaceBefore;after=[double]$pf.SpaceAfter;keepLines=[int]$pf.KeepTogether;keepNext=[int]$pf.KeepWithNext;widowControl=[int]$pf.WidowControl}
   }
   $pdf=[string](Join-Path $folder ($stage+'.pdf'));$doc.ExportAsFixedFormat($pdf,17)
   $states[$stage]=@{paragraphs=$paragraphs;pages=[int]$doc.ComputeStatistics(2);pdfHash=Hash $pdf;sourceHash=Hash $input}
   $doc.Close([ref]0);$doc=$null
  }
  $rows+=@{name=$case.name;states=$states;browserHash=Hash (Join-Path $folder 'browser-report.json')}
  Write-Output ('Verified native rendering inputs for Inter '+$case.name)
 }
 @{captured=$true;cases=$rows;fontHash=Hash $path;scriptHash=Hash $PSCommandPath;executableHash=$exe;oracleHash=Hash $oraclePath}|ConvertTo-Json -Depth 20|Set-Content (Join-Path $workspace ('.local/'+$matrix+'/native-report.json')) -Encoding UTF8
}finally{
 if($doc){$doc.Close([ref]0)}
 if($app){if($idle){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 if($added){if(-not [NofficeTemporaryFont]::Remove($path,0,[IntPtr]::Zero)){Write-Error 'Temporary font removal failed'}}
}
