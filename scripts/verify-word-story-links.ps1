$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$root=Join-Path $workspace '.local/word-story-links'
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
$app=$null;$doc=$null;$owned=$false;$rows=@();$hashes=@();$returns=@()
try{
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Requires an idle owned Word instance'};$owned=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.Options.UpdateLinksAtOpen
 $app.Visible=$false;$app.AutomationSecurity=3;$app.DisplayAlerts=0;$app.Options.UpdateLinksAtOpen=$false
 $baseline=(Get-Content (Join-Path $workspace 'docs/parity/baseline.json') -Raw|ConvertFrom-Json).applications|Where-Object executable -eq 'WINWORD.EXE'
 $exe=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 if($exe -ne $baseline.sha256){throw 'Pinned Word build changed'}
 $reference=Get-Content (Join-Path $workspace 'tests/fixtures/native-word-story-links.json') -Raw|ConvertFrom-Json
 foreach($entry in $reference.cases.PSObject.Properties){
  $name=$entry.Name;$sample=$entry.Value;$folder="browser/$name";$caseRoot=Join-Path $root $folder
  $source=Join-Path $workspace ('tests/fixtures/'+$sample.sourceFile)
  $browserPath=Join-Path $caseRoot 'browser-report.json';$browser=Get-Content $browserPath -Raw|ConvertFrom-Json
  if((Hash $source) -ne $browser.hashes.source -or (Hash $source) -ne $sample.sourceSha256){throw 'Source binding changed'}
  $kind=if($sample.action.kind -eq 'header'){'Headers'}else{'Footers'}
  $slot=switch($sample.action.slot){'default'{1};'first'{2};'even'{3}}
  foreach($stage in @('undo','linked','reloaded','edited')){
   $actualPath=[string](Join-Path $caseRoot "$stage.docx")
   if((Hash $actualPath) -ne $browser.hashes."$stage.docx"){throw 'Actual export binding changed'}
   $expectedPath=[string](Join-Path $caseRoot "expected-$stage.docx")
   Copy-Item -LiteralPath $source -Destination $expectedPath -Force
   $doc=$app.Documents.Open($expectedPath,$false,$false,$false)
   $doc.Activate();$doc.ActiveWindow.View.ReadingLayout=$false;$doc.ActiveWindow.View.Type=3
   if($stage -ne 'undo'){
    $story=$doc.Sections.Item(2).$kind.Item($slot);$story.LinkToPrevious=[bool]$sample.action.linked
    if($stage -eq 'edited'){
     $range=$story.Range.Duplicate;$range.SetRange([int]$range.Start,[int]$range.End-1);$range.Text='Updated '+$sample.action.kind
    }
   }
   $doc.SaveAs2([ref]$expectedPath,[ref]12);$doc.Close([ref]0);$doc=$null
   $doc=$app.Documents.Open($expectedPath,$false,$true,$false)
   $wanted=Snapshot $doc;$expectedPdf=[string](Join-Path $caseRoot "expected-$stage.pdf")
   $doc.ExportAsFixedFormat($expectedPdf,17);$doc.Close([ref]0);$doc=$null
   $doc=$app.Documents.Open($actualPath,$false,$true,$false)
   $actual=Snapshot $doc;$actualPdf=[string](Join-Path $caseRoot "native-$stage.pdf")
   $doc.ExportAsFixedFormat($actualPdf,17);$doc.Close([ref]0);$doc=$null
   $rows+=@{name=$name;folder=$folder;stage=$stage;expected=$wanted;actual=$actual}
   foreach($path in @($actualPath,$expectedPath,$expectedPdf,$actualPdf)){$hashes+=@{path=$path.Substring($root.Length+1).Replace('\','/');sha256=Hash $path}}
   Write-Output "Captured $name $stage"
  }
  foreach($file in @('linked.pdf','edited.pdf')){
   $path=Join-Path $caseRoot $file
   if((Hash $path) -ne $browser.hashes.$file){throw 'Browser PDF binding changed'}
   $hashes+=@{path="$folder/$file";sha256=Hash $path}
  }
  $returnPath=[string](Join-Path $caseRoot 'native-return.docx')
  Copy-Item -LiteralPath (Join-Path $caseRoot 'edited.docx') -Destination $returnPath -Force
  $doc=$app.Documents.Open($returnPath,$false,$false,$false)
  $doc.Activate();$doc.ActiveWindow.View.ReadingLayout=$false;$doc.ActiveWindow.View.Type=3
  $range=$doc.Sections.Item(2).$kind.Item($slot).Range.Duplicate
  $end=[int]$range.End-1;$range.SetRange($end,$end);$range.Text=' native'
  $doc.Save();$doc.Close([ref]0);$doc=$null
  $doc=$app.Documents.Open($returnPath,$false,$true,$false)
  $state=Snapshot $doc;$returnPdf=[string](Join-Path $caseRoot 'native-return.pdf')
  $doc.ExportAsFixedFormat($returnPdf,17);$doc.Close([ref]0);$doc=$null
  $returns+=@{name=$name;folder=$folder;native=$state}
  foreach($path in @($returnPath,$returnPdf,$browserPath)){$hashes+=@{path=$path.Substring($root.Length+1).Replace('\','/');sha256=Hash $path}}
  if((Hash $source) -ne $browser.hashes.source){throw 'Original source changed'}
 }
 if((Hash $PSCommandPath) -ne $scriptHash){throw 'Native verifier changed during capture'}
 @{captured=$true;scriptHash=$scriptHash;executableHash=$exe;rows=$rows;hashes=$hashes;returns=$returns}|ConvertTo-Json -Depth 20|Set-Content (Join-Path $root 'native-export-report.json') -Encoding UTF8
}finally{
 if($doc){$doc.Close([ref]0)}
 if($app){if($owned){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
