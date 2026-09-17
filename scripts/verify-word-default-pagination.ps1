param([switch]$Inherited)
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$matrix=if($Inherited){'word-default-inheritance'}else{'word-default-pagination'}
$root=Join-Path $workspace ('.local/'+$matrix)
$reference=Join-Path $workspace '.local/word-default-fonts/variable-embedded-native.docx'
function Hash($path){(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant()}
function Snapshot($document){
 $document.ActiveWindow.View.ReadingLayout=$false;$document.ActiveWindow.View.Type=3;$document.Repaginate()
 $paragraphs=@()
 foreach($p in $document.Paragraphs){
  $r=$p.Range;$pf=$r.ParagraphFormat;$mark=$document.Range([int]$r.End-1,[int]$r.End);$characters=@()
  for($i=[int]$r.Start;$i -lt [int]$r.End-1;$i++){
   $c=$document.Range($i,$i+1)
   $characters+=@{text=[string]$c.Text;page=[int]$c.Information(3);font=[string]$c.Font.Name;size=[double]$c.Font.Size}
  }
  $paragraphs+=@{text=[string]$r.Text;characters=$characters;markFont=[string]$mark.Font.Name;markSize=[double]$mark.Font.Size;lineRule=[int]$pf.LineSpacingRule;line=[double]$pf.LineSpacing;before=[double]$pf.SpaceBefore;after=[double]$pf.SpaceAfter}
 }
 return @{paragraphs=$paragraphs;pages=[int]$document.ComputeStatistics(2)}
}
$app=$null;$doc=$null;$idle=$false
try{
 $browserPath=Join-Path $root 'browser-report.json';$browser=Get-Content $browserPath -Raw|ConvertFrom-Json
 if(!$browser.passed){throw 'Browser workflow has not passed'}
 $baseline=(Get-Content (Join-Path $workspace 'docs/parity/baseline.json') -Raw|ConvertFrom-Json).applications|Where-Object executable -eq 'WINWORD.EXE'
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Requires an idle owned Word instance'};$idle=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.Options.UpdateLinksAtOpen
 $app.Visible=$false;$app.AutomationSecurity=3;$app.DisplayAlerts=0;$app.Options.UpdateLinksAtOpen=$false
 $exe=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 if($exe -ne $baseline.sha256){throw 'Native baseline changed'}
 if(@($app.FontNames|ForEach-Object {[string]$_}) -contains 'Inter'){throw 'Embedded-font acceptance requires Inter to be absent before opening the reference'}
 $stages=if($Inherited){@('source','edited','restored','reimported')}else{@('empty','source','edited','restored','reimported')}
 $rows=@()
 foreach($stage in $stages){
  # Independently authored Word reference, with its own embedded font. Edit a
  # disposable file copy; preserve the reference and every actual app export.
  $expectedPath=[string](Join-Path $root ('native-'+$stage+'.docx'))
  Copy-Item -LiteralPath $reference -Destination $expectedPath -Force
  $doc=$app.Documents.Open($expectedPath,$false,$false,$false)
  if($stage -eq 'empty'){$doc.Content.Text=''}
  if($stage -eq 'edited'){$end=[int]$doc.Content.End-1;$doc.Range($end,$end).InsertAfter(' added')}
  $doc.Save();$doc.Close([ref]0);$doc=$null
  $actualPath=[string](Join-Path $root ($stage+'.docx'))
  if((Hash $actualPath) -ne $browser.outputs.($stage+'.docx')){throw 'Stale actual DOCX output'}
  $states=@{}
  foreach($kind in @('expected','actual')){
   $input=if($kind -eq 'expected'){$expectedPath}else{$actualPath}
   $doc=$app.Documents.Open($input,$false,$true,$false)
   $snapshot=Snapshot $doc
   $pdf=[string](Join-Path $root ($stage+'-'+$kind+'.pdf'));$doc.ExportAsFixedFormat($pdf,17)
   $states[$kind]=@{snapshot=$snapshot;inputHash=Hash $input;pdfHash=Hash $pdf}
   $doc.Close([ref]0);$doc=$null
  }
  $rows+=@{name=$stage;states=$states}
  Write-Output ('Verified native embedded-default inputs: '+$stage)
 }
 @{captured=$true;cases=$rows;browserHash=Hash $browserPath;referenceHash=Hash $reference;scriptHash=Hash $PSCommandPath;executableHash=$exe}|ConvertTo-Json -Depth 24|Set-Content (Join-Path $root 'native-report.json') -Encoding UTF8
}finally{
 if($doc){$doc.Close([ref]0)}
 if($app){if($idle){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
