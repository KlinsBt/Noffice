param([switch]$SourceOnly, [switch]$Widow, [switch]$Terminal, [switch]$Orphan)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot $(if($Orphan){'../.local/word-orphan'}elseif($Terminal){'../.local/word-terminal'}elseif($Widow){'../.local/word-widow'}else{'../.local/word-pagination'})))
$baseline=Get-Content -LiteralPath (Join-Path $PSScriptRoot '../docs/parity/baseline.json') -Raw|ConvertFrom-Json
function Hash($path){return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()}
function Snapshot($document) {
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
$app=$null;$doc=$null;$links=$null;$idle=$false
try {
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Requires an idle owned Word instance'};$idle=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts
 $app.Visible=$false;$app.DisplayAlerts=0;$app.AutomationSecurity=3
 $links=$app.Options.UpdateLinksAtOpen;$app.Options.UpdateLinksAtOpen=$false
 $exeHash=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 if($exeHash -ne ($baseline.applications|Where-Object executable -eq 'WINWORD.EXE').sha256){throw 'Word baseline changed'}
 $source=[string](Join-Path $root 'source.docx');$binding=Get-Content -LiteralPath (Join-Path $root 'source.json') -Raw|ConvertFrom-Json
 if((Hash $source) -ne $binding.sha256){throw 'Stale source'}
 $doc=$app.Documents.Open($source,$false,$false,$false);$expected=Snapshot $doc
 $pdf=[string](Join-Path $root 'source-native.pdf');$doc.ExportAsFixedFormat($pdf,17)
 if(!$SourceOnly -or $Widow -or $Terminal -or $Orphan){
  $doc.Activate();$doc.ActiveWindow.View.ReadingLayout=$false;$doc.ActiveWindow.View.Type=3
  if($Orphan){$doc.Paragraphs.Item(2).WidowControl=-1}
  elseif($Terminal){$end=[int]$doc.Paragraphs.Item(1).Range.End-1;$doc.Range($end,$end).InsertAfter('!')}
  elseif($Widow){$doc.Paragraphs.Item(1).WidowControl=-1}
  else{$end=[int]$doc.Content.End-1;$doc.Range($end,$end).InsertAfter('!')}
  $copy=[string](Join-Path $root 'expected-edited.docx');$doc.SaveAs2([ref]$copy,[ref]12)
 }
 $doc.Close([ref]0);$doc=$null
 $hashes=@(@{path='source.docx';sha256=Hash $source},@{path='source-native.pdf';sha256=Hash $pdf})
 $fonts=@();foreach($name in @('arial.ttf','arialbd.ttf','ariali.ttf','arialbi.ttf')){$fonts+=@{name=$name;sha256=Hash (Join-Path $env:WINDIR "Fonts/$name")}}
 $report=@{source=$expected;sourceHash=Hash $source;version=[string]$app.Version;build=[string]$app.Build;executableHash=$exeHash;fonts=$fonts;printer=[string]$app.ActivePrinter;scriptHash=Hash $PSCommandPath;hashes=$hashes}
 if($SourceOnly -and ($Widow -or $Terminal -or $Orphan)){
  $doc=$app.Documents.Open($copy,$false,$true,$false);$report.edited=Snapshot $doc
  $pdf=[string](Join-Path $root 'expected-edited.pdf');$doc.ExportAsFixedFormat($pdf,17)
  $doc.Close([ref]0);$doc=$null
  foreach($name in @('expected-edited.docx','expected-edited.pdf')){$report.hashes+=@{path=$name;sha256=Hash (Join-Path $root $name)}}
 }
 if(!$SourceOnly){
  $browser=Get-Content -LiteralPath (Join-Path $root 'browser-report.json') -Raw|ConvertFrom-Json
  if(!$browser.passed -or $browser.sourceHash -ne (Hash $source)){throw 'Stale browser run'}
  AssertPlacements $expected $browser.positions
  if($Terminal -and ($browser.caret.page -ne $expected.paragraphs[0].caret.page -or !$browser.caret.collapsed -or $browser.caret.height -le 0)){throw 'Empty-line caret differs from native Word'}
  $export=[string](Join-Path $root 'browser.docx')
  if((Hash $export) -ne $browser.exportHash){throw 'Stale export'}
  $doc=$app.Documents.Open($export,$false,$true,$false);$actual=Snapshot $doc
  $pdf=[string](Join-Path $root 'browser-native.pdf');$doc.ExportAsFixedFormat($pdf,17)
  $doc.Close([ref]0);$doc=$null
  if(($actual|ConvertTo-Json -Depth 12 -Compress) -cne ($expected|ConvertTo-Json -Depth 12 -Compress)){throw 'Exported file differs from restored source'}
  $report.actual=$actual;$report.browser=$browser
  $doc=$app.Documents.Open($copy,$false,$true,$false);$wantedEdit=Snapshot $doc
  $pdf=[string](Join-Path $root 'expected-edited.pdf');$doc.ExportAsFixedFormat($pdf,17)
  $doc.Close([ref]0);$doc=$null
  $edited=[string](Join-Path $root 'browser-edited.docx')
  if((Hash $edited) -ne $browser.editedHash){throw 'Stale edited export'}
  $doc=$app.Documents.Open($edited,$false,$true,$false);$actualEdit=Snapshot $doc
  $pdf=[string](Join-Path $root 'browser-edited.pdf');$doc.ExportAsFixedFormat($pdf,17)
  $doc.Close([ref]0);$doc=$null
  if(($actualEdit|ConvertTo-Json -Depth 12 -Compress) -cne ($wantedEdit|ConvertTo-Json -Depth 12 -Compress)){throw 'Edited browser export differs from native edit'}
  if($Widow -or $Terminal -or $Orphan){
   AssertPlacements $wantedEdit $browser.editedPositions
   $report.hashes+=@{path='browser-edited-print.pdf';sha256=Hash (Join-Path $root 'browser-edited-print.pdf')}
  }
  $report.edited=$actualEdit
  foreach($name in @('expected-edited.docx','expected-edited.pdf','browser-edited.docx','browser-edited.pdf')){$report.hashes+=@{path=$name;sha256=Hash (Join-Path $root $name)}}
  foreach($name in @('browser.docx','browser-native.pdf','browser-print.pdf','browser-report.json')){$report.hashes+=@{path=$name;sha256=Hash (Join-Path $root $name)}}
 }
 $report|ConvertTo-Json -Depth 16|Set-Content -Encoding UTF8 (Join-Path $root $(if($SourceOnly){'source-native.json'}else{'native-report.json'}))
 Write-Output "Native pagination: $($expected.pages) pages, $($expected.paragraphs.Count) semantic paragraphs."
}finally{
 try{if($doc){$doc.Close([ref]0)}}finally{if($app){try{if($idle){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}}}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}}
}
