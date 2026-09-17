param([switch]$Create, [switch]$Compare, [switch]$Probe, [switch]$UniformLeading, [switch]$Reference, [switch]$WrappedLeading)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/word-mixed-lines'))
if($Probe){
 $root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/word-baselines'))
}
if($UniformLeading){
 $root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/word-uniform-leading'))
}
if($WrappedLeading){
 $root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/word-wrapped-leading'))
}
function Hash($path){(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()}
function Snapshot($document){
 $document.Repaginate();$paragraphs=@()
 foreach($p in $document.Paragraphs){
  $r=$p.Range;$chars=@()
  for($i=[int]$r.Start;$i -lt [int]$r.End;$i++){
   $c=$document.Range($i,$i+1)
   $chars+=@{text=[string]$c.Text;size=[double]$c.Font.Size;family=[string]$c.Font.Name;x=[double]$c.Information(5);y=[double]$c.Information(6);page=[int]$c.Information(3)}
  }
  $paragraphs+=@{text=[string]$r.Text;rule=[int]$p.LineSpacingRule;spacing=[double]$p.LineSpacing;chars=$chars}
 }
 return $paragraphs
}
$app=$null;$doc=$null;$idle=$false;$alerts=$null;$security=$null;$links=$null
try{
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Native test requires an idle owned Word instance'}
 $idle=$true;$alerts=$app.DisplayAlerts;$security=$app.AutomationSecurity;$links=$app.Options.UpdateLinksAtOpen
 $app.Visible=$false;$app.DisplayAlerts=0;$app.AutomationSecurity=3;$app.Options.UpdateLinksAtOpen=$false
 $baseline=Get-Content -LiteralPath (Join-Path $PSScriptRoot '../docs/parity/baseline.json') -Raw|ConvertFrom-Json
 $exeHash=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 if($exeHash -ne ($baseline.applications|Where-Object executable -eq 'WINWORD.EXE').sha256){throw 'Word baseline changed'}
 $source=[string](Join-Path $root 'source.docx')
 if($Create){
  $doc=$app.Documents.Open([string](Join-Path $root 'authored.docx'),$false,$false,$false)
  $doc.SaveAs2([ref]$source,[ref]12);$doc.Close([ref]0);$doc=$null
 }
 $doc=$app.Documents.Open($source,$false,$true,$false)
 $initial=Snapshot $doc
 $doc.ExportAsFixedFormat([string](Join-Path $root 'source-native.pdf'),17)
 if($Probe){$doc.ExportAsFixedFormat([string](Join-Path $root 'source-native.xps'),18)}
 $doc.Close([ref]0);$doc=$null
 $fonts=@();foreach($name in @('arial.ttf','arialbd.ttf','ariali.ttf','arialbi.ttf')){$fonts+=@{name=$name;sha256=Hash (Join-Path $env:WINDIR "Fonts/$name")}}
 $report=@{sourceSha256=Hash $source;scriptSha256=Hash $PSCommandPath;executableSha256=$exeHash;fonts=$fonts;version=[string]$app.Version;build=[string]$app.Build;printer=[string]$app.ActivePrinter;source=$initial;stages=@();passed=$true;artifacts=@()}
 if($Compare -or $Reference){
  if($Compare){
  $browser=Get-Content -LiteralPath (Join-Path $root 'browser-report.json') -Raw|ConvertFrom-Json
  if(!$browser.passed -or $browser.sourceSha256 -ne (Hash $source)){throw 'Stale browser source receipt'}
  $from=0;$to=6
  if($Probe){$from=12;$to=15}
  if($UniformLeading){$from=9;$to=12}
  if($WrappedLeading){$from=1;$to=1} # Wrapped lines are checked individually by the PDF verifier.
  for($i=$from;$i -lt $to;$i++){
   $advance=[double]$initial[$i+1].chars[0].y-[double]$initial[$i].chars[0].y
   $height=if($Probe -or $UniformLeading){[double]$browser.sourceHeights[$i-$from]}else{[double]$browser.heights[$i]}
   if([Math]::Abs($height*0.75-$advance) -gt 0.15){throw "Browser height differs at paragraph $i : $advance"}
  }
  }
  $stages=if($UniformLeading -or $WrappedLeading){@('double','minimum','single')}elseif($Probe){@('size','empty','typing')}else{@('size','minimum')}
  foreach($stage in $stages){
   $doc=$app.Documents.Open($source,$false,$false,$false)
   if($UniformLeading -or $WrappedLeading){
    $indices=if($WrappedLeading){@(2,3)}else{@(10)}
    foreach($index in $indices){
     $p=$doc.Paragraphs.Item($index)
     if($stage -eq 'minimum'){$p.LineSpacingRule=3;$p.LineSpacing=$(if($WrappedLeading){18}else{50})}
     elseif($stage -eq 'double'){$p.LineSpacingRule=2}
     else{$p.LineSpacingRule=0}
    }
   }else{
   $paragraphIndex=if($Probe){13}else{2}
   $start=[int]$doc.Paragraphs.Item($paragraphIndex).Range.Start
   if($Probe){$doc.Range($start+3,$start+8).Font.Size=20}else{$doc.Range($start+6,$start+11).Font.Size=20}
   if($stage -in @('empty','typing')){
    [void]$doc.Range($start,[int]$doc.Paragraphs.Item($paragraphIndex).Range.End-1).Delete()
    if($stage -eq 'typing'){$doc.Activate();$doc.Range($start,$start).Select();$app.Selection.TypeText('New')}
   }
   if($stage -eq 'minimum'){$doc.Paragraphs.Item(2).LineSpacingRule=3;$doc.Paragraphs.Item(2).LineSpacing=50}
   }
   $expectedPath=[string](Join-Path $root "expected-$stage.docx")
   $doc.SaveAs2([ref]$expectedPath,[ref]12);$doc.Close([ref]0);$doc=$null
   $doc=$app.Documents.Open($expectedPath,$false,$true,$false);$expected=Snapshot $doc
   $doc.ExportAsFixedFormat([string](Join-Path $root "expected-$stage.pdf"),17);$doc.Close([ref]0);$doc=$null
   if($Compare){
   $export=[string](Join-Path $root "browser-$stage.docx")
   if((Hash $export) -ne $browser.exports.$stage){throw 'Stale browser export'}
   $doc=$app.Documents.Open($export,$false,$true,$false);$actual=Snapshot $doc
   $doc.ExportAsFixedFormat([string](Join-Path $root "browser-$stage.pdf"),17);$doc.Close([ref]0);$doc=$null
   if(($actual|ConvertTo-Json -Depth 12 -Compress) -cne ($expected|ConvertTo-Json -Depth 12 -Compress)){throw "Native exported file differs: $stage"}
   }
   $report.stages+=@{stage=$stage;expected=$expected;exportSha256=$(if($Compare){Hash $export}else{$null});expectedSha256=Hash $expectedPath}
  }
 }
 foreach($f in Get-ChildItem -LiteralPath $root -File){
  if($f.Extension -in @('.docx','.pdf','.xps') -or $f.Name -eq 'browser-report.json'){$report.artifacts+=@{path=$f.Name;sha256=Hash $f.FullName}}
 }
 $report|ConvertTo-Json -Depth 16|Set-Content -Encoding UTF8 (Join-Path $root $(if($Compare){'native-report.json'}elseif($Reference){'reference-report.json'}else{'source-report.json'}))
 Write-Output "Native mixed lines: $($initial.Count) paragraphs; $($report.stages.Count) compared exports."
}finally{
 try{if($doc){$doc.Close([ref]0)}}finally{if($app){try{if($idle){$app.Options.UpdateLinksAtOpen=$links;$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Quit([ref]0)}}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}}
}

