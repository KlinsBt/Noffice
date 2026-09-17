# Unequal-column acceptance keeps its independent source/print matrix separate
# from the earlier section-flow receipt set while this dependency is extended.
param([string]$Case='narrow-first-hard',[switch]$NegativeControl)
$Family='unequal'
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$root=Join-Path $workspace ".local/word-$Family-acceptance/$Case"
$prefix='word-unequal-flow'
$fixture=[string](Join-Path $workspace "tests/fixtures/$prefix-$Case.docx")
$referencePath=Join-Path $workspace "tests/fixtures/native-$prefix.json"
$reference=(Get-Content $referencePath -Raw|ConvertFrom-Json).cases.$Case
if(!$reference){throw 'Unknown flow fixture'}
$baseline=(Get-Content (Join-Path $workspace 'docs/parity/baseline.json') -Raw|ConvertFrom-Json).applications|Where-Object executable -eq 'WINWORD.EXE'
function Hash($path){(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant()}
function Snapshot($document){
 $document.Repaginate();$paragraphs=@()
 foreach($p in $document.Paragraphs){
  $r=$p.Range;$characters=@()
  for($i=[int]$r.Start;$i -lt [int]$r.End-1;$i++){
   $c=$document.Range($i,$i+1)
   $characters+=@{text=[string]$c.Text;page=[int]$c.Information(3);x=[double]$c.Information(5);y=[double]$c.Information(6);font=[string]$c.Font.Name;size=[double]$c.Font.Size}
  }
  $caret=$document.Range([int]$r.End-1,[int]$r.End-1)
  $paragraphs+=@{text=[string]$r.Text;characters=$characters;caret=@{page=[int]$caret.Information(3);x=[double]$caret.Information(5);y=[double]$caret.Information(6)}}
 }
 return @{paragraphs=$paragraphs}
}
function AssertNativeSnapshot($expected,$actual){
 if($expected.paragraphs.Count -ne $actual.paragraphs.Count){throw 'Paragraph count differs'}
 for($p=0;$p -lt $expected.paragraphs.Count;$p++){
  $a=$expected.paragraphs[$p];$b=$actual.paragraphs[$p]
  if($a.text -cne $b.text -or $a.characters.Count -ne $b.characters.Count){throw "Paragraph text differs: $p"}
  if($a.caret -and ($a.caret.page -ne $b.caret.page -or [Math]::Abs($a.caret.x-$b.caret.x) -gt .001 -or [Math]::Abs($a.caret.y-$b.caret.y) -gt .001)){throw "Native caret differs: $p; expected $($a.caret|ConvertTo-Json -Compress), actual $($b.caret|ConvertTo-Json -Compress)"}
  for($i=0;$i -lt $a.characters.Count;$i++){
   $x=$a.characters[$i];$y=$b.characters[$i]
   if($x.text -cne $y.text -or $x.page -ne $y.page -or [Math]::Abs($x.x-$y.x) -gt .001 -or [Math]::Abs($x.y-$y.y) -gt .001){throw "Native character differs: $p/$i"}
   if($x.font -and ($x.font -cne $y.font -or $x.size -ne $y.size)){throw "Native character font differs: $p/$i"}
  }
 }
}
# Physical DOM column origins are compared to native PDF glyph origins by the
# Python gate. COM screen coordinates have a distinct printer quantization.
function Browser($expected,$positions,$lines,$carets){
 if($positions.Count -ne $expected.paragraphs.Count -or $lines.Count -ne $expected.paragraphs.Count){throw 'Browser paragraph count differs'}
 for($p=0;$p -lt $positions.Count;$p++){
  $chars=$expected.paragraphs[$p].characters;$line=0
  if($positions[$p].Count -ne $chars.Count){throw "Browser text length differs: $p"}
  for($i=0;$i -lt $chars.Count;$i++){
   $c=$chars[$i];$b=$positions[$p][$i]
   if($b.text -cne ([string]$c.text).Replace([string][char]11,"`n") -or $b.page -ne $c.page){throw "Browser character/page differs: $p/$i"}
   $softLine=$Case.EndsWith('-wrapped') -and $i -gt 0 -and ($c.page -ne $chars[$i-1].page -or [Math]::Abs($c.y-$chars[$i-1].y) -gt .01 -or $c.x -lt $chars[$i-1].x-.01)
   if($c.text -cne [string][char]12 -and $c.text -cne [string][char]14 -and ($i -eq 0 -or $softLine -or $chars[$i-1].text -cin @([string][char]11,[string][char]12,[string][char]14))){
    if($line -ge $lines[$p].Count -or $lines[$p][$line].page -ne $c.page){throw "Browser line/column differs: $p/$line"}
    $line++
   }
  }
  if($line -ne $lines[$p].Count){throw "Browser line count differs: $p"}
  if($Family -eq 'inline' -and ($carets.Count -ne $positions.Count -or $carets[$p].page -ne $expected.paragraphs[$p].caret.page -or [Math]::Abs($carets[$p].x-$expected.paragraphs[$p].caret.x) -gt .15)){throw "Browser caret differs: $p"}
 }
}
$app=$null;$doc=$null;$idle=$false
try{
 if((Hash $fixture) -ne $reference.sourceSha256){throw 'Source binding changed'}
 $browser=Get-Content (Join-Path $root 'browser-report.json') -Raw|ConvertFrom-Json
 if(!$browser.passed -or $browser.sourceHash -ne (Hash $fixture)){throw 'Stale browser evidence'}
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Requires an idle owned Word instance'};$idle=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.Options.UpdateLinksAtOpen
 $app.Visible=$false;$app.AutomationSecurity=3;$app.DisplayAlerts=0;$app.Options.UpdateLinksAtOpen=$false
 $exe=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE');$font=Hash (Join-Path $env:WINDIR 'Fonts/arial.ttf')
 $referenceFont=$reference.native.fontSha256
 if($exe -ne $baseline.sha256 -or $font -ne $referenceFont){throw 'Native baseline changed'}
 $doc=$app.Documents.Open($fixture,$false,$false,$false)
 $source=Snapshot $doc
 if($NegativeControl){$source.paragraphs[0].characters[0].x+=1}
 AssertNativeSnapshot $reference.native $source;Browser $source $browser.positions $browser.lines $browser.caret
 $doc.ExportAsFixedFormat([string](Join-Path $root 'source-native.pdf'),17)
 $paragraphIndex=1
 $end=[int]$doc.Paragraphs.Item($paragraphIndex).Range.End-1
 $addition=([string][char]11)+'added'
 $doc.Range($end,$end).InsertAfter($addition)
 $copy=[string](Join-Path $root 'expected-edited.docx');$doc.SaveAs2([ref]$copy,[ref]12)
 $doc.Close([ref]0);$doc=$app.Documents.Open($copy,$false,$true,$false)
 $edited=Snapshot $doc
 $edited | ConvertTo-Json -Depth 16 | Set-Content (Join-Path $root 'expected-edited-snapshot.json') -Encoding UTF8
 Browser $edited $browser.editedPositions $browser.editedLines $browser.editedCaret
 $doc.ExportAsFixedFormat([string](Join-Path $root 'expected-edited.pdf'),17)
 $doc.Close([ref]0);$doc=$null
 # The browser restores text after exporting and reopening its edited DOCX.
 # Mirror that save session in Word: its first Save can normalize source runs
 # even on pages where the appended text did not change the visible content.
 $doc=$app.Documents.Open($copy,$false,$false,$false)
 $end=[int]$doc.Paragraphs.Item($paragraphIndex).Range.End-1
 $doc.Range($end-$addition.Length,$end).Delete()|Out-Null
 $restoredPath=[string](Join-Path $root 'expected-restored.docx');$doc.SaveAs2([ref]$restoredPath,[ref]12)
 $doc.Close([ref]0);$doc=$app.Documents.Open($restoredPath,$false,$true,$false)
 $restored=Snapshot $doc;AssertNativeSnapshot $source $restored
 $doc.ExportAsFixedFormat([string](Join-Path $root 'expected-restored.pdf'),17)
 $doc.Close([ref]0);$doc=$null
 foreach($stage in @('source','edited')){
  $path=[string](Join-Path $root $(if($stage -eq 'source'){'browser.docx'}else{'browser-edited.docx'}))
  $binding=if($stage -eq 'source'){$browser.exportHash}else{$browser.editedHash}
  if((Hash $path) -ne $binding){throw 'Export binding changed'}
  $doc=$app.Documents.Open($path,$false,$true,$false)
  $expected=if($stage -eq 'source'){$restored}else{$edited};AssertNativeSnapshot $expected (Snapshot $doc)
  $doc.ExportAsFixedFormat([string](Join-Path $root $(if($stage -eq 'source'){'browser-native.pdf'}else{'browser-edited.pdf'})),17)
  $doc.Close([ref]0);$doc=$null
 }
 $hashes=@();foreach($name in @('browser.docx','browser-edited.docx','expected-edited.docx','expected-restored.docx','expected-restored.pdf','browser-native.pdf','source-native.pdf','expected-edited.pdf','browser-edited.pdf','browser-print.pdf','browser-edited-print.pdf','browser-report.json')){
  $hashes+=@{path=$name;sha256=Hash (Join-Path $root $name)}
 }
 if($browser.pdfDownloadHash){
  if((Hash (Join-Path $root 'download-source.pdf')) -ne $browser.pdfDownloadHash -or (Hash (Join-Path $root 'download-edited.pdf')) -ne $browser.editedPdfDownloadHash){throw 'PDF download binding changed'}
  foreach($name in @('download-source.pdf','download-edited.pdf')){$hashes+=@{path=$name;sha256=Hash (Join-Path $root $name)}}
 }
 @{passed=$true;case=$Case;source=$source;edited=$edited;restored=$restored;hashes=$hashes;sourceHash=Hash $fixture;referenceHash=Hash $referencePath;scriptHash=Hash $PSCommandPath;executableHash=$exe;fontHash=$font;printer=[string]$app.ActivePrinter}|ConvertTo-Json -Depth 18|Set-Content (Join-Path $root 'native-report.json') -Encoding UTF8
 Write-Output "Native unequal-column flow verified: $Case, original and independently edited actual exports."
}finally{
 if($doc){$doc.Close([ref]0)}
 if($app){if($idle){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
