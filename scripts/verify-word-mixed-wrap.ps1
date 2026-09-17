param([switch]$Create, [switch]$Compare, [switch]$Reference)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/word-mixed-wrap'))
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
 $doc.Close([ref]0);$doc=$null
 $fonts=@();foreach($name in @('arial.ttf','arialbd.ttf','ariali.ttf','arialbi.ttf')){$fonts+=@{name=$name;sha256=Hash (Join-Path $env:WINDIR "Fonts/$name")}}
 $report=@{sourceSha256=Hash $source;scriptSha256=Hash $PSCommandPath;executableSha256=$exeHash;fonts=$fonts;version=[string]$app.Version;build=[string]$app.Build;printer=[string]$app.ActivePrinter;source=$initial;stages=@();passed=$true;artifacts=@()}
 if($Compare -or $Reference){
  if($Compare){
  $browser=Get-Content -LiteralPath (Join-Path $root 'browser-report.json') -Raw|ConvertFrom-Json
  if(!$browser.passed -or $browser.sourceSha256 -ne (Hash $source)){throw 'Stale browser source receipt'}
  }
  $stages=@('double','minimum','single','size','typing')
  foreach($stage in $stages){
   $doc=$app.Documents.Open($source,$false,$false,$false)
   if($stage -in @('size','typing')){
    $r=$doc.Paragraphs.Item(2).Range;$start=[int]$r.Start
    if($stage -eq 'size'){
     $offset=([string]$r.Text).IndexOf('alpha5')
     if($offset -lt 0){throw 'Missing authored size target'}
     $doc.Range($start+$offset,$start+$offset+6).Font.Size=20
    }else{
     $doc.Activate();$doc.Range($start+5,$start+5).Select();$app.Selection.TypeText('0')
    }
   }else{
    foreach($index in @(2,3)){
     $p=$doc.Paragraphs.Item($index)
     if($stage -eq 'minimum'){$p.LineSpacingRule=3;$p.LineSpacing=50}
     elseif($stage -eq 'double'){$p.LineSpacingRule=2}
     else{$p.LineSpacingRule=0}
    }
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
 if($Compare){
  $reedit=@{}
  foreach($kind in @('expected','browser')){
   $input=[string](Join-Path $root "$kind-typing.docx")
   $doc=$app.Documents.Open($input,$false,$false,$false)
   $r=$doc.Paragraphs.Item(2).Range;$offset=([string]$r.Text).IndexOf('alpha3')
   if($offset -lt 0){throw 'Missing native re-edit target'}
   $start=[int]$r.Start+$offset;$doc.Range($start,$start+6).Font.Size=15
   $output=[string](Join-Path $root "native-reedit-$kind.docx")
   $doc.SaveAs2([ref]$output,[ref]12);$doc.Close([ref]0);$doc=$null
   $doc=$app.Documents.Open($output,$false,$true,$false)
   $reedit[$kind]=Snapshot $doc
   $doc.ExportAsFixedFormat([string](Join-Path $root "native-reedit-$kind.pdf"),17)
   $doc.Close([ref]0);$doc=$null
  }
  if(($reedit.expected|ConvertTo-Json -Depth 12 -Compress) -cne ($reedit.browser|ConvertTo-Json -Depth 12 -Compress)){throw 'Native re-edit/reopen differs'}
  $report.nativeReedit=@{passed=$true;expected=$reedit.expected;actual=$reedit.browser;inputExpectedSha256=Hash (Join-Path $root 'expected-typing.docx');inputBrowserSha256=Hash (Join-Path $root 'browser-typing.docx')}
 }
 foreach($f in Get-ChildItem -LiteralPath $root -File){
  if($f.Extension -in @('.docx','.pdf','.xps') -or $f.Name -eq 'browser-report.json'){$report.artifacts+=@{path=$f.Name;sha256=Hash $f.FullName}}
 }
 $report|ConvertTo-Json -Depth 16|Set-Content -Encoding UTF8 (Join-Path $root $(if($Compare){'native-report.json'}elseif($Reference){'reference-report.json'}else{'source-report.json'}))
 Write-Output "Native mixed wrapping: $($initial.Count) paragraphs; $($report.stages.Count) compared exports."
}finally{
 try{if($doc){$doc.Close([ref]0)}}finally{if($app){try{if($idle){$app.Options.UpdateLinksAtOpen=$links;$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Quit([ref]0)}}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}}
}

