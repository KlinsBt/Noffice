$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$root=Join-Path $workspace '.local/word-paragraph-fonts'
$baseline=(Get-Content (Join-Path $workspace 'docs/parity/baseline.json') -Raw|ConvertFrom-Json).applications|Where-Object executable -eq 'WINWORD.EXE'
$cases=(Get-Content (Join-Path $workspace 'tests/fixtures/native-word-paragraph-fonts.json') -Raw|ConvertFrom-Json).cases
function Hash($path){(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant()}
function Snapshot($document){
 $paragraphs=@()
 foreach($p in $document.Paragraphs){
  $r=$p.Range;$mark=$document.Range([int]$r.End-1,[int]$r.End);$characters=@()
  for($i=[int]$r.Start;$i -lt [int]$r.End-1;$i++){
   $c=$document.Range($i,$i+1)
   $characters+=@{text=[string]$c.Text;font=[string]$c.Font.Name;size=[double]$c.Font.Size}
  }
  $paragraphs+=@{text=[string]$r.Text;mark=@{font=[string]$mark.Font.Name;size=[double]$mark.Font.Size};characters=$characters}
 }
 return @{paragraphs=$paragraphs}
}
$app=$null;$doc=$null;$idle=$false
try{
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Requires an idle owned Word instance'};$idle=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.Options.UpdateLinksAtOpen
 $app.Visible=$false;$app.AutomationSecurity=3;$app.DisplayAlerts=0;$app.Options.UpdateLinksAtOpen=$false
 $exe=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 if($exe -ne $baseline.sha256){throw 'Native baseline changed'}
 foreach($case in $cases){
  $dir=Join-Path $root $case.name
  $browser=Get-Content (Join-Path $dir 'browser-report.json') -Raw|ConvertFrom-Json
  if(-not $browser.passed){throw 'Browser acceptance failed'}
  $stages=@('edited','restored');if($browser.outputs.typed){$stages+='typed'}
  foreach($stage in (@('source')+$stages)){
   if((Hash (Join-Path $dir ($stage+'.docx'))) -ne $browser.outputs.$stage){throw "Stale browser output: $stage"}
  }
  $source=[string](Join-Path $dir 'source.docx')
  $doc=$app.Documents.Open($source,$false,$true,$false);$doc.Activate()
  $doc.ActiveWindow.View.ReadingLayout=$false;$doc.ActiveWindow.View.Type=3
  $app.Selection.SetRange([int]$case.selection.start,[int]$case.selection.end)
  $app.Selection.Font.Name='Courier New';$app.Selection.Font.Size=20
  $native=[string](Join-Path $dir 'native.docx');$doc.SaveAs2([ref]$native,[ref]12)
  $doc.Close([ref]0);$doc=$app.Documents.Open($native,$false,$true,$false)
  $expected=Snapshot $doc
  $doc.ExportAsFixedFormat([string](Join-Path $dir 'native.pdf'),17)
  $doc.Close([ref]0);$doc=$null
  $typedExpected=$null
  if($browser.outputs.typed){
   $doc=$app.Documents.Open($native,$false,$true,$false);$doc.Activate()
   $doc.ActiveWindow.View.ReadingLayout=$false;$doc.ActiveWindow.View.Type=3
   $app.Selection.SetRange([int]$case.selection.end,[int]$case.selection.end)
   $app.Selection.TypeText('!')
   $typed=[string](Join-Path $dir 'native-typed.docx');$doc.SaveAs2([ref]$typed,[ref]12)
   $doc.Close([ref]0);$doc=$app.Documents.Open($typed,$false,$true,$false)
   $typedExpected=Snapshot $doc;$doc.ExportAsFixedFormat([string](Join-Path $dir 'native-typed.pdf'),17)
   $doc.Close([ref]0);$doc=$null
  }
  $actual=@{}
  foreach($stage in $stages){
   $doc=$app.Documents.Open([string](Join-Path $dir ($stage+'.docx')),$false,$true,$false)
   $actual[$stage]=Snapshot $doc
   $doc.ExportAsFixedFormat([string](Join-Path $dir ($stage+'-native.pdf')),17)
   $doc.Close([ref]0);$doc=$null
  }
  $hashes=@()
  foreach($name in @('source.docx','edited.docx','restored.docx','native.docx','native.pdf','edited-native.pdf','restored-native.pdf','browser-report.json')){
   $hashes+=@{path=$name;sha256=Hash (Join-Path $dir $name)}
  }
  if($browser.outputs.typed){foreach($name in @('typed.docx','native-typed.docx','native-typed.pdf','typed-native.pdf')){$hashes+=@{path=$name;sha256=Hash (Join-Path $dir $name)}}}
  $fontHashes=@();foreach($fontName in @('arial.ttf','cour.ttf')){$fontHashes+=@{name=$fontName;sha256=Hash (Join-Path $env:WINDIR "Fonts/$fontName")}}
  @{expected=$expected;typedExpected=$typedExpected;actual=$actual;hashes=$hashes;fontHashes=$fontHashes;scriptHash=Hash $PSCommandPath;executableHash=$exe}|ConvertTo-Json -Depth 20|Set-Content (Join-Path $dir 'native-report.json') -Encoding UTF8
  Write-Output "Captured paragraph font exports: $($case.name)"
 }
}finally{
 if($doc){$doc.Close([ref]0)}
 if($app){if($idle){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
