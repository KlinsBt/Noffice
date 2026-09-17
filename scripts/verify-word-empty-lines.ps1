param([switch]$Create,[switch]$Compare)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/word-empty-lines'))
function Hash($p){(Get-FileHash -LiteralPath $p).Hash.ToLowerInvariant()}
function Snapshot($doc){
 $doc.ActiveWindow.View.Type=3
 $doc.ActiveWindow.View.Zoom.Percentage=100
 $doc.Repaginate();$result=@()
 foreach($p in $doc.Paragraphs){
  $chars=@();$r=$p.Range
  for($i=[int]$r.Start;$i -lt [int]$r.End;$i++){
   $c=$doc.Range($i,$i+1)
   $chars+=@{text=[string]$c.Text;size=[double]$c.Font.Size;family=[string]$c.Font.Name;x=[double]$c.Information(5);y=[double]$c.Information(6);page=[int]$c.Information(3)}
  }
  $result+=@{text=[string]$r.Text;chars=$chars;rule=[int]$p.LineSpacingRule;spacing=[double]$p.LineSpacing}
 }
 return ,$result
}
function Edit($doc,$stage){
 if($stage -eq 'typing'){
  $end=[int]$doc.Paragraphs.Item(13).Range.End-1
  $doc.Activate();$doc.Range($end,$end).Select();$doc.Application.Selection.TypeText('New');return
 }
 for($i=1;$i -le 18;$i+=2){
  $p=$doc.Paragraphs.Item($i)
  if($stage -eq 'double'){$p.LineSpacingRule=2}
  if($stage -eq 'minimum'){$p.LineSpacingRule=3;$p.LineSpacing=50}
  if($stage -eq 'single'){$p.LineSpacingRule=0}
 }
}
$app=$null;$doc=$null;$idle=$false
try{
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Native test needs an idle owned Word instance.'}
 $idle=$true;$alerts=$app.DisplayAlerts;$security=$app.AutomationSecurity;$links=$app.Options.UpdateLinksAtOpen
 $app.Visible=$false;$app.DisplayAlerts=0;$app.AutomationSecurity=3;$app.Options.UpdateLinksAtOpen=$false
 $baseline=Get-Content -Raw (Join-Path $PSScriptRoot '../docs/parity/baseline.json')|ConvertFrom-Json
 $exeHash=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 if($exeHash -ne ($baseline.applications|Where-Object executable -eq 'WINWORD.EXE').sha256){throw 'Word baseline changed.'}
 $source=[string](Join-Path $root 'source.docx')
 if($Create){$doc=$app.Documents.Open([string](Join-Path $root 'authored.docx'),$false,$false,$false);$doc.SaveAs2([ref]$source,[ref]12);$doc.Close([ref]0);$doc=$null;exit}
 $browser=if($Compare){Get-Content -Raw (Join-Path $root 'browser-report.json')|ConvertFrom-Json}else{$null}
 if($Compare -and (!$browser.passed -or $browser.sourceSha256 -ne (Hash $source))){throw 'Stale source receipt.'}
 $stages=@()
 foreach($stage in @('source','double','minimum','single','typing')){
  $doc=$app.Documents.Open($source,$false,$false,$false);Edit $doc $stage
  $expectedPath=[string](Join-Path $root "expected-$stage.docx")
  $doc.SaveAs2([ref]$expectedPath,[ref]12);$doc.Close([ref]0);$doc=$null
  $doc=$app.Documents.Open($expectedPath,$false,$true,$false);$expected=Snapshot $doc
  $doc.ExportAsFixedFormat([string](Join-Path $root "expected-$stage.pdf"),17);$doc.Close([ref]0);$doc=$null
  if($Compare -and $stage -ne 'source'){
   $output=[string](Join-Path $root "browser-$stage.docx")
   if((Hash $output) -ne $browser.exports.$stage){throw 'Stale browser export.'}
   $doc=$app.Documents.Open($output,$false,$true,$false);$actual=Snapshot $doc
   $doc.ExportAsFixedFormat([string](Join-Path $root "browser-$stage.pdf"),17);$doc.Close([ref]0);$doc=$null
   if(($expected|ConvertTo-Json -Depth 12 -Compress) -cne ($actual|ConvertTo-Json -Depth 12 -Compress)){throw "Native export mismatch: $stage"}
  }
  $stages+=@{stage=$stage;paragraphs=$expected;expectedSha256=Hash $expectedPath;expectedPdfSha256=Hash (Join-Path $root "expected-$stage.pdf");exportSha256=$(if($Compare -and $stage -ne 'source'){Hash $output}else{$null});exportPdfSha256=$(if($Compare -and $stage -ne 'source'){Hash (Join-Path $root "browser-$stage.pdf")}else{$null})}
 }
 $report=@{passed=$true;sourceSha256=Hash $source;scriptSha256=Hash $PSCommandPath;executableSha256=$exeHash;fontSha256=Hash (Join-Path $env:WINDIR 'Fonts/arial.ttf');stages=$stages}
 [IO.File]::WriteAllText((Join-Path $root $(if($Compare){'native-report.json'}else{'reference-report.json'})),($report|ConvertTo-Json -Depth 16),(New-Object Text.UTF8Encoding($false)))
 Write-Output 'Native empty-line metrics: 18 paragraphs across five stages.'
}finally{
 try{if($doc){$doc.Close([ref]0)}}finally{if($app){try{if($idle){$app.Options.UpdateLinksAtOpen=$links;$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Quit([ref]0)}}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}}
}
