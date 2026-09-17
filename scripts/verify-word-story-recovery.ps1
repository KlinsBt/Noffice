$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$root=Join-Path $workspace '.local/word-story-options'
function Hash($path){(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant()}
$scriptHash=Hash $PSCommandPath
function Snapshot($document){
 $document.Activate();$document.ActiveWindow.View.ReadingLayout=$false;$document.ActiveWindow.View.Type=3
 $document.Repaginate();$sections=@()
 foreach($section in $document.Sections){
  $stories=@()
  foreach($kind in @('Headers','Footers')){foreach($slot in @(1,2,3)){
   $story=$section.$kind.Item($slot);$paragraphs=@();$characters=@()
   if($story.Exists){
    foreach($p in $story.Range.Paragraphs){
     $r=$p.Range;$f=$r.ParagraphFormat
     $paragraphs+=@{text=[string]$r.Text;lineRule=[int]$f.LineSpacingRule;line=[double]$f.LineSpacing;before=[double]$f.SpaceBefore;after=[double]$f.SpaceAfter}
    }
    foreach($c in $story.Range.Characters){$characters+=@{text=[string]$c.Text;font=[string]$c.Font.Name;size=[double]$c.Font.Size;bold=[int]$c.Font.Bold;italic=[int]$c.Font.Italic;underline=[int]$c.Font.Underline}}
   }
   $stories+=@{kind=$kind;slot=$slot;exists=[bool]$story.Exists;linked=[bool]$story.LinkToPrevious;paragraphs=$paragraphs;characters=$characters}
  }}
  $p=$section.PageSetup
  $sections+=@{stories=$stories;width=[double]$p.PageWidth;height=[double]$p.PageHeight;top=[double]$p.TopMargin;bottom=[double]$p.BottomMargin;header=[double]$p.HeaderDistance;footer=[double]$p.FooterDistance;first=[bool]$p.DifferentFirstPageHeaderFooter;even=[bool]$p.OddAndEvenPagesHeaderFooter}
 }
 return @{pages=[int]$document.ComputeStatistics(2);body=[string]$document.Content.Text;sections=$sections}
}
$app=$null;$doc=$null;$owned=$false;$rows=@();$hashes=@()
try{
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Requires an idle owned Word instance'};$owned=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.Options.UpdateLinksAtOpen
 $app.Visible=$false;$app.AutomationSecurity=3;$app.DisplayAlerts=0;$app.Options.UpdateLinksAtOpen=$false
 $baseline=(Get-Content (Join-Path $workspace 'docs/parity/baseline.json') -Raw|ConvertFrom-Json).applications|Where-Object executable -eq 'WINWORD.EXE'
 $exe=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 if($exe -ne $baseline.sha256){throw 'Pinned Word build changed'}
 $optionReceipt=Join-Path $root 'native-export-report.json'
 $source=Join-Path $workspace 'tests/fixtures/word-section-stories/first-even.docx'
 foreach($name in @('late-abort','other-tab')){
  $folder="recovery/$name";$caseRoot=Join-Path $root $folder
  $browserPath=Join-Path $caseRoot 'browser-report.json';$browser=Get-Content $browserPath -Raw|ConvertFrom-Json
  $actualPath=[string](Join-Path $caseRoot 'unsaved.docx')
  if((Hash $actualPath) -ne $browser.exportHash -or (Hash $source) -ne $browser.sourceHash){throw 'Recovery artifact binding changed'}
  $doc=$app.Documents.Open($actualPath,$false,$true,$false)
  $state=Snapshot $doc;$pdf=[string](Join-Path $caseRoot 'native-unsaved.pdf')
  $doc.ExportAsFixedFormat($pdf,17);$doc.Close([ref]0);$doc=$null
  $rows+=@{name=$name;folder=$folder;actual=$state}
  foreach($path in @($actualPath,$pdf,$browserPath)){$hashes+=@{path=$path.Substring($root.Length+1).Replace('\','/');sha256=Hash $path}}
  Write-Output "Captured actual unsaved recovery export $name"
 }
 if((Hash $PSCommandPath) -ne $scriptHash){throw 'Recovery verifier changed during capture'}
 @{captured=$true;scriptHash=$scriptHash;executableHash=$exe;optionReceiptHash=Hash $optionReceipt;rows=$rows;hashes=$hashes}|ConvertTo-Json -Depth 20|Set-Content (Join-Path $root 'recovery-native-report.json') -Encoding UTF8
}finally{
 if($doc){$doc.Close([ref]0)}
 if($app){if($owned){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
