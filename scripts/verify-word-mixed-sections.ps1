$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/word-mixed'))
$baseline=Get-Content -LiteralPath (Join-Path $PSScriptRoot '../docs/parity/baseline.json') -Raw|ConvertFrom-Json
function Hash($path){return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()}
function Snapshot($document) {
 $sections=@()
 for($i=1;$i -le $document.Sections.Count;$i++){
  $s=$document.Sections.Item($i);$p=$s.PageSetup;$paragraphs=@()
  foreach($para in $s.Range.Paragraphs){$r=$para.Range;$paragraphs+=@{text=[string]$r.Text;font=[string]$r.Font.Name;size=[double]$r.Font.Size}}
  $sections+=@{text=[string]$s.Range.Text;width=[math]::Round($p.PageWidth*20);height=[math]::Round($p.PageHeight*20);left=[math]::Round($p.LeftMargin*20);right=[math]::Round($p.RightMargin*20);top=[math]::Round($p.TopMargin*20);bottom=[math]::Round($p.BottomMargin*20);paragraphs=$paragraphs}
 }
 return ,$sections
}
$browser=Get-Content -LiteralPath (Join-Path $root 'browser-report.json') -Raw|ConvertFrom-Json
if(!$browser.workflowPassed){throw 'Browser workflow has not passed'}
$source=[string](Join-Path $root 'source.docx')
if((Hash $source) -ne $browser.sourceHash){throw 'Stale browser source'}
$observed=@();$hashes=@();$app=$null;$doc=$null;$links=$null;$fonts=@()
foreach($name in @('arial.ttf','arialbd.ttf','arialbi.ttf','ariali.ttf')){$fonts+=@{name=$name;sha256=Hash (Join-Path $env:WINDIR "Fonts/$name")}}
foreach($operation in @('split','join')){
 $idle=$false
 try {
  $app=New-Object -ComObject Word.Application
  if($app.Documents.Count -ne 0){throw 'Requires an idle owned Word instance'};$idle=$true
  $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts
  $app.Visible=$false;$app.DisplayAlerts=0;$app.AutomationSecurity=3
  $links=$app.Options.UpdateLinksAtOpen;$app.Options.UpdateLinksAtOpen=$false
  $exeHash=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
  if($exeHash -ne ($baseline.applications|Where-Object executable -eq 'WINWORD.EXE').sha256){throw 'Word baseline changed'}
  $output=[string](Join-Path $root "browser-$operation.docx")
  if((Hash $output) -ne $browser.($operation+'Hash')){throw 'Stale browser export'}
  $editable=[string](Join-Path $root "native-input-$operation.docx")
  Copy-Item -LiteralPath $source -Destination $editable -Force
  $doc=$app.Documents.Open($editable,$false,$false,$false)
  $doc.Activate();$doc.ActiveWindow.View.ReadingLayout=$false;$doc.ActiveWindow.View.Type=3
  $initial=Snapshot $doc
  if($initial.Count -ne 2 -or $initial[0].text -cne "AlphaBeta$([char]12)" -or $initial[1].text -cne "Gamma`r"){throw 'Unexpected native fixture'}
  # Compare initial browser page rectangles and paragraph origins with explicit native margins.
  for($i=0;$i -lt 2;$i++){
   $rect=$browser.initial.rects[$i];$block=$browser.initial.blocks[$i];$native=$initial[$i]
   if([math]::Abs($rect.width-$native.width/15) -gt 0.1 -or [math]::Abs($rect.height-$native.height/15) -gt 0.1 -or [math]::Abs($block.x-$rect.x-$native.left/15) -gt 0.1 -or [math]::Abs($block.y-$rect.y-$native.top/15) -gt 0.1){throw 'Browser surface/native section geometry mismatch'}
  }
  $end=[int]$doc.Content.End-1;$doc.Range($end,$end).InsertAfter(' edited')
  $doc.Range(5,5).InsertParagraph()
  if($operation -eq 'join'){$end=[int]$doc.Sections.Item(1).Range.End;[void]$doc.Range($end-1,$end).Delete()}
  $copy=[string](Join-Path $root "native-$operation.docx");$doc.SaveAs2([ref]$copy,[ref]12)
  $doc.Close([ref]0);$doc=$null
  # Saved file comparison: reopen both paths before measuring/rendering.
  $doc=$app.Documents.Open($copy,$false,$true,$false);$doc.Repaginate();$expected=Snapshot $doc
  $pdf=[string](Join-Path $root "native-$operation.pdf");$doc.ExportAsFixedFormat($pdf,17)
  $doc.Close([ref]0);$doc=$null
  $doc=$app.Documents.Open($output,$false,$true,$false);$doc.Repaginate();$actual=Snapshot $doc
  $pdf=[string](Join-Path $root "browser-native-$operation.pdf");$doc.ExportAsFixedFormat($pdf,17)
  $doc.Close([ref]0);$doc=$null
  if(($expected|ConvertTo-Json -Depth 10 -Compress) -cne ($actual|ConvertTo-Json -Depth 10 -Compress)){throw 'Export/native semantic mismatch'}
  $observed+=@{operation=$operation;initial=$initial;expected=$expected;actual=$actual}
  $version=[string]$app.Version;$build=[string]$app.Build;$printer=[string]$app.ActivePrinter;$language=[int]$app.LanguageSettings.LanguageID(2)
 } finally {
  try{if($doc){$doc.Close([ref]0);$doc=$null}}finally{if($app){try{if($idle){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}}}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null}}}
 }
}
foreach($name in @('source.docx','browser-report.json','browser-split.docx','browser-join.docx','native-split.docx','native-join.docx','native-split.pdf','native-join.pdf','browser-native-split.pdf','browser-native-join.pdf','split-print.pdf')){$hashes+=@{path=$name;sha256=Hash (Join-Path $root $name)}}
@{passed=$true;version=$version;build=$build;executableHash=$exeHash;fonts=$fonts;printer=$printer;uiLanguage=$language;scriptHash=Hash $PSCommandPath;browser=$browser.browser;observed=$observed;hashes=$hashes}|ConvertTo-Json -Depth 15|Set-Content -Encoding UTF8 (Join-Path $root 'native-report.json')
Write-Output 'Browser mixed-section rectangles and both live exports match installed Word geometry and paragraph semantics.'
