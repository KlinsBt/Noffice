$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/word-layout'))
$baseline=Get-Content -LiteralPath (Join-Path $PSScriptRoot '../docs/parity/baseline.json') -Raw|ConvertFrom-Json
$app=$null;$doc=$null;$idle=$false;$mismatches=@();$observed=@()
function Hash($path){return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()}
function ReadSectionGeometry($section) {
 $setup=$section.PageSetup
 return @{width=[math]::Round($setup.PageWidth*20);height=[math]::Round($setup.PageHeight*20);margins=@{left=[math]::Round($setup.LeftMargin*20);right=[math]::Round($setup.RightMargin*20);top=[math]::Round($setup.TopMargin*20);bottom=[math]::Round($setup.BottomMargin*20);header=[math]::Round($setup.HeaderDistance*20);footer=[math]::Round($setup.FooterDistance*20);gutter=[math]::Round($setup.Gutter*20)}}
}
try {
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Requires an idle owned Word instance'};$idle=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts
 $app.Visible=$false;$app.DisplayAlerts=0;$app.AutomationSecurity=3
 $links=$app.Options.UpdateLinksAtOpen;$app.Options.UpdateLinksAtOpen=$false
 $exeHash=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 if($exeHash -ne ($baseline.applications|Where-Object executable -eq 'WINWORD.EXE').sha256){throw 'Word baseline changed'}
 $expected=Get-Content -LiteralPath (Join-Path $root 'model-expected.json') -Raw|ConvertFrom-Json
 $path=[string](Join-Path $root 'inheritance.docx')
 if((Hash $path) -ne $expected.sourceHash){throw 'Stale inheritance fixture'}
 $doc=$app.Documents.Open($path,$false,$true,$false)
 if($doc.Sections.Count -ne $expected.sections.Count){throw 'Section count mismatch'}
 for($i=1;$i -le $doc.Sections.Count;$i++){
  $section=$doc.Sections.Item($i);$actual=ReadSectionGeometry $section;$wanted=$expected.sections[$i-1]
  foreach($key in @('width','height')){if($actual[$key] -ne $wanted.$key){$mismatches+="section $i $key"}}
  foreach($key in @('left','right','top','bottom','header','footer','gutter')){if($null -ne $wanted.margins.$key -and $actual.margins[$key] -ne $wanted.margins.$key){$mismatches+="section $i margin $key"}}
  $slots=@('default','first','even');$refs=@()
  foreach($kind in @('headers','footers')){for($slot=1;$slot -le 3;$slot++){
   $want=$wanted.$kind.($slots[$slot-1]);if($null -eq $want){continue}
   $reference=if($kind -eq 'headers'){$section.Headers.Item($slot)}else{$section.Footers.Item($slot)}
   $refs+=@{kind=$kind;slot=$slots[$slot-1];text=[string]$reference.Range.Text;inherited=[bool]$reference.LinkToPrevious}
   if([string]$reference.Range.Text -cne $want.text -or [bool]$reference.LinkToPrevious -ne $want.inherited){$mismatches+="section $i $kind $($slots[$slot-1])"}
  }}
  $actual.references=$refs;$observed+=$actual
 }
 $doc.Close([ref]0);$doc=$null
 $browser=Get-Content -LiteralPath (Join-Path $root 'browser-report.json') -Raw|ConvertFrom-Json
 $source=[string](Join-Path $root 'custom.docx');$output=[string](Join-Path $root 'browser.docx')
 if(!$browser.workflowPassed -or (Hash $source) -ne $browser.sourceHash -or (Hash $output) -ne $browser.exportHash){throw 'Stale browser custom-page workflow'}
 $doc=$app.Documents.Open($source,$false,$false,$false)
 $customGeometry=ReadSectionGeometry ($doc.Sections.Item(1))
 $setup=$doc.Sections.Item(1).PageSetup
 $setup.PageWidth=11906/20;$setup.PageHeight=16838/20;$setup.Orientation=0
 $setup.TopMargin=72;$setup.BottomMargin=72;$setup.LeftMargin=72;$setup.RightMargin=72
 $native=[string](Join-Path $root 'native-preset.docx');$doc.SaveAs2([ref]$native,[ref]12)
 $nativeGeometry=ReadSectionGeometry ($doc.Sections.Item(1))
 $pdf=[string](Join-Path $root 'native-preset.pdf');$doc.ExportAsFixedFormat($pdf,17)
 $doc.Close([ref]0);$doc=$null
 $doc=$app.Documents.Open($output,$false,$true,$false);$browserGeometry=ReadSectionGeometry ($doc.Sections.Item(1))
 foreach($key in @('width','height')){if($nativeGeometry[$key] -ne $browserGeometry[$key]){$mismatches+="preset $key"}}
 foreach($key in @('left','right','top','bottom','header','footer','gutter')){if($nativeGeometry.margins[$key] -ne $browserGeometry.margins[$key]){$mismatches+="preset margin $key"}}
 $pdf=[string](Join-Path $root 'browser-native.pdf');$doc.ExportAsFixedFormat($pdf,17)
 $doc.Close([ref]0);$doc=$null
 $hashes=@();foreach($name in @('inheritance.docx','model-expected.json','custom.docx','browser.docx','native-preset.docx','native-preset.pdf','browser-native.pdf','browser-print.pdf','custom-print.pdf')){$hashes+=@{path=$name;sha256=Hash (Join-Path $root $name)}}
 $fonts=@();foreach($name in @('arial.ttf','arialbd.ttf','arialbi.ttf','ariali.ttf')){$fonts+=@{name=$name;sha256=Hash (Join-Path $env:WINDIR "Fonts/$name")}}
 @{version=[string]$app.Version;build=[string]$app.Build;executableHash=$exeHash;fonts=$fonts;printer=[string]$app.ActivePrinter;browser=$browser.browser;uiLanguage=[int]$app.LanguageSettings.LanguageID(2);normalization='Native COM point measurements rounded to the nearest integer twip (1/20 point).';sections=$observed;nativeCustom=$customGeometry;nativePreset=$nativeGeometry;browserPreset=$browserGeometry;mismatches=$mismatches;hashes=$hashes}|ConvertTo-Json -Depth 12|Set-Content -Encoding UTF8 (Join-Path $root 'native-report.json')
 if($mismatches.Count){throw "Native Word mismatches: $mismatches"}
 Write-Output 'Three section geometries and inherited references match Word; browser page presets match native Word.'
} finally {
 try{if($doc){$doc.Close([ref]0)}}finally{if($app){try{if($idle){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}}}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}}
}
