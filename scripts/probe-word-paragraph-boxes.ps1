param([switch]$MixedWrap)
$ErrorActionPreference='Stop'
$directory=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/word-paragraph-boxes'))
if($MixedWrap){$directory=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/word-mixed-wrap/boxes'))}
[void][IO.Directory]::CreateDirectory($directory)
function Hash($p){(Get-FileHash -LiteralPath $p).Hash.ToLowerInvariant()}
$app=$null;$doc=$null;$idle=$false
try{
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'This probe requires an idle owned Word instance.'}
 $idle=$true;$security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.Options.UpdateLinksAtOpen
 $app.Visible=$false;$app.AutomationSecurity=3;$app.DisplayAlerts=0;$app.Options.UpdateLinksAtOpen=$false
 $baseline=Get-Content -Raw (Join-Path $PSScriptRoot '../docs/parity/baseline.json')|ConvertFrom-Json
 $exe=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 if($exe -ne ($baseline.applications|Where-Object executable -eq 'WINWORD.EXE').sha256){throw 'The Word baseline changed.'}
 $source=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../tests/fixtures/word-wrapped-leading.docx'))
 if($MixedWrap){$source=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../tests/fixtures/word-mixed-wrap.docx'))}
 $stages=@()
 $names=if($MixedWrap){@('source','double','minimum','single','size','typing')}else{@('source','double','minimum','single')}
 foreach($stage in $names){
  $doc=$app.Documents.Open($source,$false,$false,$false)
  foreach($index in @(2,3)){
   $p=$doc.Paragraphs.Item($index)
   if($stage -eq 'double'){$p.LineSpacingRule=2}
   if($stage -eq 'minimum'){$p.LineSpacingRule=3;$p.LineSpacing=$(if($MixedWrap){50}else{18})}
   if($stage -eq 'single'){$p.LineSpacingRule=0}
  }
  if($stage -in @('size','typing')){
   $r=$doc.Paragraphs.Item(2).Range;$start=[int]$r.Start
   if($stage -eq 'size'){
    $offset=([string]$r.Text).IndexOf('alpha5')
    if($offset -lt 0){throw 'Missing authored size target'}
    $doc.Range($start+$offset,$start+$offset+6).Font.Size=20
   }else{
    $doc.Activate();$doc.Range($start+5,$start+5).Select();$app.Selection.TypeText('0')
   }
  }
  $path=[string](Join-Path $directory "$stage.docx")
  $doc.SaveAs2([ref]$path,[ref]12);$doc.Close([ref]0);$doc=$null
  $doc=$app.Documents.Open($path,$false,$false,$false)
  $doc.ActiveWindow.View.Type=3;$doc.Repaginate()
  $doc.ExportAsFixedFormat([string](Join-Path $directory "$stage.pdf"),17)
  $origins=@();foreach($index in @(2,3)){$origins+=@{paragraph=$index;y=[double]$doc.Paragraphs.Item($index).Range.Information(6)}}
  # These are authored measurement copies only. Color changes must be proven
  # layout-neutral by comparing every text matrix in the two PDF exports.
  $doc.Paragraphs.Item(2).Shading.BackgroundPatternColor=16776960
  $doc.Paragraphs.Item(3).Shading.BackgroundPatternColor=16711935
  $doc.Repaginate();$doc.ExportAsFixedFormat([string](Join-Path $directory "$stage-boxes.pdf"),17)
  $doc.Close([ref]0);$doc=$null
  $stages+=@{stage=$stage;origins=$origins;docxSha256=Hash $path;pdfSha256=Hash (Join-Path $directory "$stage.pdf");boxesPdfSha256=Hash (Join-Path $directory "$stage-boxes.pdf")}
 }
 $report=@{sourceSha256=Hash $source;scriptSha256=Hash $PSCommandPath;executableSha256=$exe;stages=$stages;scope='Native paragraph-background geometry research; glyph invariance must be checked separately'}
 [IO.File]::WriteAllText((Join-Path $directory 'native-report.json'),($report|ConvertTo-Json -Depth 12),(New-Object Text.UTF8Encoding($false)))
 Write-Output "Recorded $($stages.Count) native paragraph-box copies and their unshaded controls."
}finally{
 if($doc){$doc.Close([ref]0)}
 if($app){if($idle){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
