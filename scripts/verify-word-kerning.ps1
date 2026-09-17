param([switch]$NegativeControl)
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$root=Join-Path $workspace '.local/word-kerning-acceptance'
$source=[string](Join-Path $workspace 'tests/fixtures/word-kerning.docx')
$referencePath=Join-Path $workspace 'tests/fixtures/native-word-kerning.json'
function Hash($path){(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant()}
function Snapshot($doc){
 $doc.Repaginate();$paragraphs=@()
 foreach($p in $doc.Paragraphs){
  $r=$p.Range.Duplicate;$r.End=$r.End-1;$chars=@()
  for($i=[int]$r.Start;$i -lt [int]$r.End;$i++){
   $c=$doc.Range($i,$i+1)
   $chars+=@{text=[string]$c.Text;page=[int]$c.Information(3);x=[double]$c.Information(5);y=[double]$c.Information(6)}
  }
  $paragraphs+=@{text=[string]$r.Text;size=[double]$r.Font.Size;kerning=[double]$r.Font.Kerning;characters=$chars}
 }
 return ,$paragraphs
}
function AssertSame($expected,$actual){
 if($expected.Count -ne $actual.Count){throw 'Native paragraph count differs'}
 for($p=0;$p -lt $expected.Count;$p++){
  $a=$expected[$p];$b=$actual[$p]
  if($a.size -ne $b.size -or $a.kerning -ne $b.kerning -or $a.characters.Count -ne $b.characters.Count){throw "Native size/kerning/count differs: $p"}
  for($i=0;$i -lt $a.characters.Count;$i++){
   $x=$a.characters[$i];$y=$b.characters[$i]
   if($x.text -cne $y.text -or $x.page -ne $y.page -or [Math]::Abs($x.x-$y.x) -gt .001 -or [Math]::Abs($x.y-$y.y) -gt .001){throw "Native kerning character differs: $p/$i"}
  }
 }
}
$app=$null;$doc=$null;$owned=$false
try{
 $reference=Get-Content $referencePath -Raw|ConvertFrom-Json
 $browser=Get-Content (Join-Path $root 'browser-report.json') -Raw|ConvertFrom-Json
 if(!$browser.passed -or $browser.sourceHash -ne (Hash $source) -or $reference.sourceSha256 -ne (Hash $source)){throw 'Stale kerning source/browser evidence'}
 $app=New-Object -ComObject Word.Application;if($app.Documents.Count -ne 0){throw 'Requires idle owned Word'};$owned=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.Options.UpdateLinksAtOpen
 $app.Visible=$false;$app.AutomationSecurity=3;$app.DisplayAlerts=0;$app.Options.UpdateLinksAtOpen=$false
 $exe=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE');$font=Hash (Join-Path $env:WINDIR 'Fonts/arial.ttf')
 $baseline=(Get-Content (Join-Path $workspace 'docs/parity/baseline.json') -Raw|ConvertFrom-Json).applications|Where-Object executable -eq 'WINWORD.EXE'
 if($exe -ne $baseline.sha256 -or $font -ne $reference.native.fontHash){throw 'Native baseline changed'}
 $doc=$app.Documents.Open($source,$false,$false,$false);$initial=Snapshot $doc
 if($NegativeControl){$initial[2].kerning+=1}
 AssertSame $reference.native.reopened $initial
 $doc.Close([ref]0);$doc=$null
 $stages=@();$hashes=@()
 foreach($stage in $browser.stages){
  $doc=$app.Documents.Open($source,$false,$false,$false)
  if($stage.name -ne 'source'){$range=$doc.Paragraphs.Item(4).Range.Duplicate;$range.End=$range.End-1;$range.Font.Size=[double]$stage.size}
  $expected=[string](Join-Path $root "expected-$($stage.name).docx")
  $doc.SaveAs2([ref]$expected,[ref]12);$doc.Close([ref]0);$doc=$null
  $doc=$app.Documents.Open($expected,$false,$true,$false);$snapshot=Snapshot $doc
  $expectedPdf=[string](Join-Path $root "expected-$($stage.name).pdf");$doc.ExportAsFixedFormat($expectedPdf,17);$doc.Close([ref]0);$doc=$null
  $actual=[string](Join-Path $root "$($stage.name).docx")
  if((Hash $actual) -ne $stage.exportHash -or (Hash (Join-Path $root "$($stage.name).pdf")) -ne $stage.printHash){throw 'Stale actual exported file'}
  $doc=$app.Documents.Open($actual,$false,$true,$false);AssertSame $snapshot (Snapshot $doc)
  $actualPdf=[string](Join-Path $root "native-$($stage.name).pdf");$doc.ExportAsFixedFormat($actualPdf,17);$doc.Close([ref]0);$doc=$null
  $stages+=@{name=$stage.name;size=$stage.size;native=$snapshot}
  foreach($path in @($expected,$expectedPdf,$actual,$actualPdf,(Join-Path $root "$($stage.name).pdf"))){$hashes+=@{path=[IO.Path]::GetFileName($path);sha256=Hash $path}}
 }
 $zero=[string](Join-Path $root 'new-zero.docx');if((Hash $zero) -ne $browser.newZeroHash){throw 'Stale zero export'}
 $doc=$app.Documents.Open($zero,$false,$false,$false);$zeroInitial=Snapshot $doc
 $zeroInitial|ConvertTo-Json -Depth 8|Set-Content (Join-Path $root 'zero-snapshot.json') -Encoding UTF8
 if($zeroInitial.Count -ne 1 -or $zeroInitial[0].text -cne '11 AV To' -or $zeroInitial[0].kerning -ne 0 -or $zeroInitial[0].size -ne 10){throw 'Native zero kerning export differs'}
 $zeroSaved=[string](Join-Path $root 'native-zero-reopened.docx');$doc.SaveAs2([ref]$zeroSaved,[ref]12);$doc.Close([ref]0);$doc=$null
 $doc=$app.Documents.Open($zeroSaved,$false,$true,$false);AssertSame $zeroInitial (Snapshot $doc);$doc.Close([ref]0);$doc=$null
 if((Hash $source) -ne $reference.sourceSha256){throw 'Original changed'}
 foreach($path in @($zero,$zeroSaved,(Join-Path $root 'browser-report.json'))){$hashes+=@{path=[IO.Path]::GetFileName($path);sha256=Hash $path}}
 @{passed=$true;stages=$stages;zero=$zeroInitial;hashes=$hashes;sourceHash=Hash $source;referenceHash=Hash $referencePath;
  scriptHash=Hash $PSCommandPath;executableHash=$exe;fontHash=$font;printer=[string]$app.ActivePrinter
 }|ConvertTo-Json -Depth 16|Set-Content (Join-Path $root 'native-report.json') -Encoding UTF8
 Write-Output 'Native kerning: all twelve source contexts, three actual exports and new-zero save/reopen pass.'
}finally{
 if($doc){$doc.Close([ref]0)}
 if($app){if($owned){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
