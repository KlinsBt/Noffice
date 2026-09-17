$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$root=Join-Path $workspace '.local/word-authored-pagination'
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
  $format=$r.ParagraphFormat;$mark=$document.Range([int]$r.End-1,[int]$r.End)
  $paragraphs+=@{text=[string]$r.Text;characters=$characters;lineRule=[int]$format.LineSpacingRule;line=[double]$format.LineSpacing;before=[double]$format.SpaceBefore;after=[double]$format.SpaceAfter;markFont=[string]$mark.Font.Name;markSize=[double]$mark.Font.Size}
 }
 return @{paragraphs=$paragraphs;pages=[int]$document.ComputeStatistics(2)}
}
$app=$null;$doc=$null;$idle=$false
try{
 $browser=Get-Content (Join-Path $root 'browser-report.json') -Raw|ConvertFrom-Json
 if(!$browser.passed){throw 'Browser workflow has not passed'}
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Requires an idle owned Word instance'};$idle=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.Options.UpdateLinksAtOpen
 $app.Visible=$false;$app.AutomationSecurity=3;$app.DisplayAlerts=0;$app.Options.UpdateLinksAtOpen=$false
 $exe=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 if($exe -ne $baseline.sha256){throw 'Native baseline changed'}
 $stages=@{}
 foreach($stage in @('source','edited','restored','automatic','minimum')){
  # Independently create the stated public editing case in Word, without using
  # the application's exported document as the expected input.
  $doc=$app.Documents.Add();$doc.Activate();$doc.ActiveWindow.View.ReadingLayout=$false;$doc.ActiveWindow.View.Type=3
  $setup=$doc.PageSetup;$setup.PageWidth=595.3;$setup.PageHeight=841.9
  $setup.TopMargin=72;$setup.BottomMargin=72;$setup.LeftMargin=72;$setup.RightMargin=72
  $setup.HeaderDistance=35.4;$setup.FooterDistance=35.4;$setup.Gutter=0
  $text=(1..70|ForEach-Object {'line'+([string]$_).PadLeft(2,'0')}) -join [char]11
  $doc.Content.Text=$text
  $doc.Content.Font.Name='Arial';$doc.Content.Font.Size=10;$doc.Content.Font.Kerning=0
  # The independently repeated actions start from the app's explicit 12pt-after
  # default; font and line-spacing commands do not change paragraph-after spacing.
  $format=$doc.Content.ParagraphFormat;$format.SpaceBefore=0;$format.SpaceAfter=12
  $format.SpaceBeforeAuto=0;$format.SpaceAfterAuto=0;$format.LineSpacingRule=4;$format.LineSpacing=12
  $format.KeepTogether=0;$format.KeepWithNext=0;$format.WidowControl=-1
  if($stage -eq 'automatic'){$format.LineSpacingRule=1}
  if($stage -eq 'minimum'){$format.LineSpacingRule=3;$format.LineSpacing=18}
  if($stage -in @('edited','restored')){
   $end=[int]$doc.Content.End-1;$doc.Range($end,$end).InsertAfter(' added')
   if($stage -eq 'restored'){$end=[int]$doc.Content.End-1;$doc.Range($end-6,$end).Delete()|Out-Null}
  }
  $expected=[string](Join-Path $root "expected-$stage.docx");$doc.SaveAs2([ref]$expected,[ref]12)
  $doc.Close([ref]0);$doc=$app.Documents.Open($expected,$false,$true,$false)
  $wanted=Snapshot $doc;$doc.ExportAsFixedFormat([string](Join-Path $root "expected-$stage.pdf"),17)
  $doc.Close([ref]0);$doc=$null
  $path=[string](Join-Path $root "browser-$stage.docx")
  if((Hash $path) -ne $browser.exports.$stage){throw 'Export binding changed'}
  $doc=$app.Documents.Open($path,$false,$true,$false)
  $actual=Snapshot $doc;$doc.ExportAsFixedFormat([string](Join-Path $root "native-$stage.pdf"),17)
  $doc.Close([ref]0);$doc=$null
  $stages[$stage]=@{expected=$wanted;actual=$actual}
 }
 $hashes=@();foreach($name in @('source','edited','restored','automatic','minimum')){
  if((Hash (Join-Path $root "download-$name.pdf")) -ne $browser.pdfExports.$name){throw 'PDF export binding changed'}
  foreach($artifact in @("expected-$name.docx","expected-$name.pdf","browser-$name.docx","browser-$name.pdf","native-$name.pdf","download-$name.pdf")){
   $hashes+=@{path=$artifact;sha256=Hash (Join-Path $root $artifact)}
  }
 }
 $hashes+=@{path='browser-report.json';sha256=Hash (Join-Path $root 'browser-report.json')}
 $hashes+=@{path='browser-reimported.pdf';sha256=Hash (Join-Path $root 'browser-reimported.pdf')}
 if((Hash (Join-Path $root 'download-reimported.pdf')) -ne $browser.pdfExports.reimported){throw 'Reimported PDF binding changed'}
 $hashes+=@{path='download-reimported.pdf';sha256=Hash (Join-Path $root 'download-reimported.pdf')}
 @{captured=$true;stages=$stages;hashes=$hashes;scriptHash=Hash $PSCommandPath;executableHash=$exe;fontHash=Hash (Join-Path $env:WINDIR 'Fonts/arial.ttf');printer=[string]$app.ActivePrinter}|ConvertTo-Json -Depth 18|Set-Content (Join-Path $root 'native-report.json') -Encoding UTF8
 Write-Output 'Captured five independent authored Word cases and first-open browser exports. Comparison remains required.'
}finally{
 if($doc){$doc.Close([ref]0)}
 if($app){if($idle){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
