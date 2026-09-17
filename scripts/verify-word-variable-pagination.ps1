param([ValidateSet('mixed','single','keep')][string]$Case='mixed')
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$root=Join-Path $workspace ".local/word-variable-$Case"
[void][IO.Directory]::CreateDirectory($root)
$fixture=[string](Join-Path $workspace "tests/fixtures/word-variable-$Case.docx")
$baseline=Get-Content (Join-Path $workspace 'docs/parity/baseline.json') -Raw|ConvertFrom-Json
$reference=(Get-Content (Join-Path $workspace 'tests/fixtures/native-word-variable-pagination.json') -Raw|ConvertFrom-Json).cases.$Case
function Hash($path){
 $stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
 $digest=[Security.Cryptography.SHA256]::Create()
 try{return ([BitConverter]::ToString($digest.ComputeHash($stream))).Replace('-','').ToLowerInvariant()}
 finally{$stream.Dispose();$digest.Dispose()}
}
function Snapshot($document) {
 $document.ActiveWindow.View.ReadingLayout=$false;$document.ActiveWindow.View.Type=3
 $document.ActiveWindow.View.Zoom.Percentage=100
 $document.Repaginate();$paragraphs=@()
 foreach($p in $document.Paragraphs){
  $lines=@();$last='';$r=$p.Range
  for($i=[int]$r.Start;$i -lt [int]$r.End-1;$i++){
   $c=$document.Range($i,$i+1);$page=[int]$c.Information(3);$y=[double]$c.Information(6)
   $key="$page/$y"
   if($key -cne $last){$lines+=@{from=$i-[int]$r.Start;to=$i-[int]$r.Start+1;page=$page;y=$y;text=[string]$c.Text};$last=$key}
   else{$line=$lines[$lines.Count-1];$line.to=$i-[int]$r.Start+1;$line.text+=[string]$c.Text}
  }
  $end=[int]$r.End-1;$caret=$document.Range($end,$end)
  $paragraphs+=@{caret=@{page=[int]$caret.Information(3);y=[double]$caret.Information(6)};text=[string]$r.Text;keep=[int]$p.KeepTogether;widow=[int]$p.WidowControl;lines=$lines}
 }
 # ComputeStatistics can return the pre-repagination count on this build. Use the
 # observed physical page assignments; the PDF comparison independently checks it.
 return @{pages=[int](($paragraphs | ForEach-Object {$_.lines} | ForEach-Object {$_.page} | Measure-Object -Maximum).Maximum);paragraphs=$paragraphs}
}
function AssertPlacements($snapshot,$positions) {
  if($positions.Count -ne $snapshot.paragraphs.Count){throw 'Browser/native paragraph count differs'}
  for($p=0;$p -lt $snapshot.paragraphs.Count;$p++){
   $length=([string]$snapshot.paragraphs[$p].text).Length-1
   if($positions[$p].Count -ne $length){throw "Browser/native text length differs at paragraph $p"}
   foreach($line in $snapshot.paragraphs[$p].lines){
    for($i=[int]$line.from;$i -lt [int]$line.to;$i++){
     $character=$positions[$p][$i]
     $text=([string]$line.text).Substring($i-[int]$line.from,1).Replace([string][char]11,"`n")
     if($character.page -ne $line.page -or $character.text -cne $text){throw "Browser/native character placement differs at paragraph $p offset $i"}
    }
   }
  }
}

function CompareSnapshots($expected,$actual,[bool]$CompareCoordinates=$true){
 if($expected.pages -ne $actual.pages -or $expected.paragraphs.Count -ne $actual.paragraphs.Count){throw 'Native paragraph/page count changed'}
 for($i=0;$i -lt $expected.paragraphs.Count;$i++){
  $a=$expected.paragraphs[$i];$b=$actual.paragraphs[$i]
  if($a.text -cne $b.text -or $a.keep -ne $b.keep -or $a.widow -ne $b.widow -or $a.lines.Count -ne $b.lines.Count){throw "Native paragraph differs: $i"}
  for($j=0;$j -lt $a.lines.Count;$j++){
   $x=$a.lines[$j];$y=$b.lines[$j]
   if($x.text -cne $y.text -or $x.from -ne $y.from -or $x.to -ne $y.to -or $x.page -ne $y.page -or ($CompareCoordinates -and [Math]::Abs($x.y-$y.y) -gt .001)){throw "Native line differs: $i/$j"}
  }
 }
}
$app=$null;$doc=$null;$idle=$false
try{
 if((Hash $fixture) -ne $reference.sourceSha256){throw 'Fixture hash changed'}
 # Native editing uses a verified local copy, never the checked source. Shared
 # reads also permit comparison when another Word window has that source open.
 $nativeFixture=[string](Join-Path $root 'source-copy.docx')
 $stream=[IO.File]::Open($fixture,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
 try{$copyStream=[IO.File]::Create($nativeFixture);try{$stream.CopyTo($copyStream)}finally{$copyStream.Dispose()}}finally{$stream.Dispose()}
 if((Hash $nativeFixture) -ne $reference.sourceSha256){throw 'Source copy changed during read'}
 $browser=Get-Content (Join-Path $root 'browser-report.json') -Raw|ConvertFrom-Json
 if(!$browser.passed -or $browser.sourceHash -ne (Hash $fixture)){throw 'Stale browser evidence'}
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Requires an idle owned Word instance'};$idle=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.Options.UpdateLinksAtOpen
 $app.Visible=$false;$app.AutomationSecurity=3;$app.DisplayAlerts=0;$app.Options.UpdateLinksAtOpen=$false
 $exeHash=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 $fontHash=Hash (Join-Path $env:WINDIR 'Fonts/arial.ttf')
 if($fontHash -ne $reference.nativeReceipt.fontSha256){throw 'Native font baseline changed'}
 if($exeHash -ne ($baseline.applications|Where-Object executable -eq 'WINWORD.EXE').sha256){throw 'Word baseline changed'}
 $doc=$app.Documents.Open($nativeFixture,$false,$false,$false)
 $expected=Snapshot $doc
 # Historical COM caret Y values drift even when every native PDF pixel and
 # glyph stays identical. Retain that UI-oracle discrepancy explicitly. Source
 # rendering is instead pinned exactly below; fresh native/export coordinates
 # still use the unchanged .001pt comparison in every later call.
 CompareSnapshots $reference.snapshot $expected $false;AssertPlacements $expected $browser.positions
 $referenceCoordinateDiagnostics=@()
 for($p=0;$p -lt $expected.paragraphs.Count;$p++){
  for($line=0;$line -lt $expected.paragraphs[$p].lines.Count;$line++){
   $before=$reference.snapshot.paragraphs[$p].lines[$line];$current=$expected.paragraphs[$p].lines[$line]
   if([Math]::Abs($before.y-$current.y) -gt .001){$referenceCoordinateDiagnostics+=@{paragraph=$p;line=$line;referenceY=$before.y;currentY=$current.y}}
  }
 }
 $pdf=[string](Join-Path $root 'source-native.pdf');$doc.ExportAsFixedFormat($pdf,17)
 & python (Join-Path $PSScriptRoot 'compare-word-variable-reference.py') --case $Case
 if($LASTEXITCODE -ne 0){throw 'Pinned native source render differs; preserve the reference and investigate'}
 $end=[int]$doc.Content.End-1;$doc.Range($end,$end).InsertAfter('!')
 $copy=[string](Join-Path $root 'expected-edited.docx');$doc.SaveAs2([ref]$copy,[ref]12)
 $doc.Close([ref]0);$doc=$null
 $doc=$app.Documents.Open($copy,$false,$true,$false);$wantedEdit=Snapshot $doc
 $pdf=[string](Join-Path $root 'expected-edited.pdf');$doc.ExportAsFixedFormat($pdf,17)
 $doc.Close([ref]0);$doc=$null
 $actuals=@()
 foreach($stage in @('source','edited')){
  $name=if($stage -eq 'source'){'browser.docx'}else{'browser-edited.docx'}
  $exportPath=[string](Join-Path $root $name)
  $binding=if($stage -eq 'source'){$browser.exportHash}else{$browser.editedHash}
  if((Hash $exportPath) -ne $binding){throw 'Stale browser export'}
  $doc=$app.Documents.Open($exportPath,$false,$true,$false);$actual=Snapshot $doc
  $expectedStage=if($stage -eq 'source'){$expected}else{$wantedEdit}
  CompareSnapshots $expectedStage $actual
  $positions=if($stage -eq 'source'){$browser.positions}else{$browser.editedPositions}
  AssertPlacements $actual $positions
  $pdf=[string](Join-Path $root $(if($stage -eq 'source'){'browser-native.pdf'}else{'browser-edited.pdf'}));$doc.ExportAsFixedFormat($pdf,17)
  $doc.Close([ref]0);$doc=$null;$actuals+=@{stage=$stage;snapshot=$actual}
 }
 $keepSnapshot=$null;$extraFiles=@()
 if($browser.keepHash){
  $doc=$app.Documents.Open($nativeFixture,$false,$false,$false)
  $doc.Paragraphs.Item(2).KeepWithNext=-1;$doc.Paragraphs.Item(3).KeepTogether=-1
  $doc.Paragraphs.Item(4).PageBreakBefore=-1;$doc.Paragraphs.Item(9).WidowControl=-1;$doc.Paragraphs.Item(10).KeepWithNext=-1
  $copy=[string](Join-Path $root 'expected-keep.docx');$doc.SaveAs2([ref]$copy,[ref]12)
  $doc.Close([ref]0);$doc=$app.Documents.Open($copy,$false,$true,$false)
  $keepSnapshot=Snapshot $doc;AssertPlacements $keepSnapshot $browser.keepPositions
  $pdf=[string](Join-Path $root 'expected-keep.pdf');$doc.ExportAsFixedFormat($pdf,17)
  $doc.Close([ref]0);$doc=$null
  $exportPath=[string](Join-Path $root 'browser-keep.docx');if((Hash $exportPath) -ne $browser.keepHash){throw 'Stale keep export'}
  $doc=$app.Documents.Open($exportPath,$false,$true,$false);CompareSnapshots $keepSnapshot (Snapshot $doc)
  $pdf=[string](Join-Path $root 'browser-keep.pdf');$doc.ExportAsFixedFormat($pdf,17)
  $doc.Close([ref]0);$doc=$null
  $extraFiles=@('expected-keep.docx','expected-keep.pdf','browser-keep.docx','browser-keep.pdf','browser-keep-print.pdf')
 }
 $hashes=@();foreach($name in (@('browser.docx','browser-edited.docx','expected-edited.docx','source-native.pdf','expected-edited.pdf','browser-native.pdf','browser-edited.pdf','browser-print.pdf','browser-edited-print.pdf','browser-report.json','native-source-reference.json')+$extraFiles)){$hashes+=@{path=$name;sha256=Hash (Join-Path $root $name)}}
 @{passed=$true;scope='Actual export properties, exact native source render and same-session export coordinates; historical COM caret-coordinate discrepancy remains separate native UI evidence.';referenceCoordinateDiagnostics=$referenceCoordinateDiagnostics;historicalCoordinatesMatch=($referenceCoordinateDiagnostics.Count -eq 0);nativeUiVerified=$false;view=3;zoom=100;case=$Case;source=$expected;edited=$wantedEdit;keep=$keepSnapshot;actual=$actuals;hashes=$hashes;sourceHash=Hash $fixture;executableHash=$exeHash;fontHash=$fontHash;scriptHash=Hash $PSCommandPath;referenceHash=Hash (Join-Path $workspace 'tests/fixtures/native-word-variable-pagination.json');printer=[string]$app.ActivePrinter}|ConvertTo-Json -Depth 18|Set-Content (Join-Path $root 'native-report.json') -Encoding UTF8
 Write-Output "Native variable pagination verified: $Case, $($expected.pages) pages and both actual exports."
}finally{
 if($doc){$doc.Close([ref]0)}
 if($app){if($idle){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
