$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/word-ranges'))
$baseline=Get-Content -LiteralPath (Join-Path $PSScriptRoot '../docs/parity/baseline.json') -Raw|ConvertFrom-Json
$app=$null;$doc=$null;$observed=@();$links=$null
function Hash($path){return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()}
function ReadRangeSections($document) {
 $sections=@()
 for($i=1;$i -le $document.Sections.Count;$i++){
  $s=$document.Sections.Item($i);$paragraphs=@()
  for($j=1;$j -le $s.Range.Paragraphs.Count;$j++){$p=$s.Range.Paragraphs.Item($j).Range;$paragraphs+=@{start=[int]$p.Start;end=[int]$p.End;text=[string]$p.Text;font=[string]$p.Font.Name;fontSize=[double]$p.Font.Size;bold=[int]$p.Font.Bold;italic=[int]$p.Font.Italic}}
  $references=@();foreach($kind in @('Headers','Footers')){for($slot=1;$slot -le 3;$slot++){$ref=$s.$kind.Item($slot);$references+=@{kind=$kind;slot=$slot;text=[string]$ref.Range.Text;linked=[bool]$ref.LinkToPrevious}}}
  $sections+=@{start=[int]$s.Range.Start;end=[int]$s.Range.End;text=[string]$s.Range.Text;width=[math]::Round($s.PageSetup.PageWidth*20);height=[math]::Round($s.PageSetup.PageHeight*20);paragraphs=$paragraphs;references=$references}
 }
 return ,$sections
}
try {
 $app=New-Object -ComObject Word.Application
 $app.Visible=$false;$app.DisplayAlerts=0;$app.AutomationSecurity=3
 $links=$app.Options.UpdateLinksAtOpen;$app.Options.UpdateLinksAtOpen=$false
 $exeHash=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 if($exeHash -ne ($baseline.applications|Where-Object executable -eq 'WINWORD.EXE').sha256){throw 'Word baseline changed'}
 $source=[string](Join-Path $root 'source.docx')
 $receipt=Get-Content -LiteralPath (Join-Path $root 'source.json') -Raw|ConvertFrom-Json
 if((Hash $source) -ne $receipt.sourceHash){throw 'Stale source fixture'}
 foreach($case in @('split','delete')){
  $doc=$app.Documents.Open($source,$false,$false,$false)
  $before=ReadRangeSections $doc
  if($before.Count -ne 2 -or $before[0].text -cne "AlphaBeta$([char]12)" -or $before[1].text -cne "Gamma`r"){throw 'Unexpected native starting state'}
  if($case -eq 'split'){$doc.Range(5,5).InsertParagraph()}
  else{$end=[int]$doc.Sections.Item(1).Range.End;[void]$doc.Range($end-1,$end).Delete()}
  $after=ReadRangeSections $doc
  if($case -eq 'split'){
   if($after.Count -ne 2 -or $after[0].text -cne "Alpha`rBeta$([char]12)" -or $after[1].text -cne "Gamma`r" -or $after[0].width -ne 12240){throw 'Native split expectation failed'}
  }else{if($after.Count -ne 1 -or $after[0].text -cne "AlphaBetaGamma`r" -or $after[0].width -ne 16838){throw 'Native delete expectation failed'}}
  $path=[string](Join-Path $root "native-$case.docx");$doc.SaveAs2([ref]$path,[ref]12)
  $doc.Close([ref]0);$doc=$null
  $observed+=@{case=$case;before=$before;after=$after;outputHash=Hash $path}
 }
 $browser=Get-Content -LiteralPath (Join-Path $root 'browser-report.json') -Raw|ConvertFrom-Json
 $browserSource=[string](Join-Path $root 'browser-source.docx');$browserPath=[string](Join-Path $root 'browser.docx')
 if(!$browser.workflowPassed -or (Hash $browserSource) -ne $browser.sourceHash -or (Hash $browserPath) -ne $browser.exportHash){throw 'Stale browser section-range workflow'}
 $doc=$app.Documents.Open($browserSource,$false,$false,$false)
 $doc.ActiveWindow.View.ReadingLayout=$false
 $end=[int]$doc.Content.End-1;$doc.Range($end,$end).InsertAfter(' edited')
 $expected=ReadRangeSections $doc
 $path=[string](Join-Path $root 'native-browser-expected.docx');$doc.SaveAs2([ref]$path,[ref]12)
 $doc.Close([ref]0);$doc=$null
 $doc=$app.Documents.Open($browserPath,$false,$true,$false)
 $actual=ReadRangeSections $doc
 if(($actual|ConvertTo-Json -Depth 10 -Compress) -cne ($expected|ConvertTo-Json -Depth 10 -Compress)){throw 'Actual browser recovery export does not match independent native edit'}
 $doc.Close([ref]0);$doc=$null
 $liveResults=@()
 foreach($operation in @('split','join')){
  $app.Options.UpdateLinksAtOpen=$links
  $app.Quit([ref]0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
  $app=New-Object -ComObject Word.Application
  $app.Visible=$false;$app.DisplayAlerts=0;$app.AutomationSecurity=3
  $links=$app.Options.UpdateLinksAtOpen;$app.Options.UpdateLinksAtOpen=$false
  if((Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')) -ne $exeHash){throw 'Word baseline changed between cases'}
  $export=[string](Join-Path $root "browser-$operation.docx")
  if((Hash $export) -ne $browser.($operation+'Hash')){throw 'Stale live boundary export'}
  $editable=[string](Join-Path $root "native-input-$operation.docx")
  Copy-Item -LiteralPath $browserSource -Destination $editable -Force
  $doc=$app.Documents.Open($editable,$false,$false,$false)
  $doc.Activate()
  $doc.ActiveWindow.View.ReadingLayout=$false
  $doc.ActiveWindow.View.Type=3
  Write-Output ("Native {0}: ReadOnly={1}, ReadingLayout={2}" -f $operation,$doc.ReadOnly,$doc.ActiveWindow.View.ReadingLayout)
  $end=[int]$doc.Content.End-1;$doc.Range($end,$end).InsertAfter(' edited')
  if($operation -eq 'split'){$doc.Range(5,5).InsertParagraph()}else{$end=[int]$doc.Sections.Item(1).Range.End;[void]$doc.Range($end-1,$end).Delete()}
  $wanted=ReadRangeSections $doc
  $copy=[string](Join-Path $root "native-browser-$operation.docx");$doc.SaveAs2([ref]$copy,[ref]12)
  $doc.Close([ref]0);$doc=$null
  $doc=$app.Documents.Open($copy,$false,$true,$false)
  $doc.Repaginate()
  $nativePdf=[string](Join-Path $root "native-$operation.pdf");$doc.ExportAsFixedFormat($nativePdf,17)
  $doc.Close([ref]0);$doc=$null
  $doc=$app.Documents.Open($export,$false,$true,$false);$seen=ReadRangeSections $doc
  $doc.Repaginate()
  $browserPdf=[string](Join-Path $root "browser-$operation.pdf");$doc.ExportAsFixedFormat($browserPdf,17)
  $doc.Close([ref]0);$doc=$null
  if(($wanted|ConvertTo-Json -Depth 10 -Compress) -cne ($seen|ConvertTo-Json -Depth 10 -Compress)){throw "Live boundary mismatch: $operation"}
  $liveResults+=@{operation=$operation;expected=$wanted;actual=$seen;nativeHash=Hash $copy;browserHash=Hash $export;nativePdfHash=Hash $nativePdf;browserPdfHash=Hash $browserPdf}
 }
 $fonts=@();foreach($name in @('times.ttf','timesbd.ttf','timesbi.ttf','timesi.ttf')){$fonts+=@{name=$name;sha256=Hash (Join-Path $env:WINDIR "Fonts/$name")}}
 @{version=[string]$app.Version;build=[string]$app.Build;executableHash=$exeHash;sourceHash=$receipt.sourceHash;scriptHash=Hash $PSCommandPath;fonts=$fonts;printer=[string]$app.ActivePrinter;uiLanguage=[int]$app.LanguageSettings.LanguageID(2);cases=$observed;browser=$browser;browserActual=$actual;browserExpected=$expected;browserNativeHash=Hash $path;liveResults=$liveResults;passed=$true;scope='Native COM paragraph split/delete semantics and actual browser live-boundary exports; no native UI or browser-pagination parity claim.'}|ConvertTo-Json -Depth 12|Set-Content -Encoding UTF8 (Join-Path $root 'native-report.json')
 Write-Output 'Word keeps the section break after the split text; deleting it adopts the following section geometry.'
}finally{
 try{if($doc){$doc.Close([ref]0)}}finally{if($app){try{if($null -ne $links){$app.Options.UpdateLinksAtOpen=$links};$app.Quit([ref]0)}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}}
}
