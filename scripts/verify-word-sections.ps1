param([switch]$Create, [switch]$Render)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/word-sections'))
[IO.Directory]::CreateDirectory($root)|Out-Null
$app=$null;$doc=$null
$baseline=Get-Content -LiteralPath (Join-Path $PSScriptRoot '../docs/parity/baseline.json') -Raw|ConvertFrom-Json
$wordBaseline=$baseline.applications|Where-Object executable -eq 'WINWORD.EXE'
function Snapshot($d) {
 $d.Repaginate();$sections=@();$paragraphs=@();$shapes=@();$notes=@()
 foreach($s in $d.Sections){$headers=@();foreach($h in $s.Headers){$headers+=@{kind=[int]$h.Index;linked=[bool]$h.LinkToPrevious;text=[string]$h.Range.Text}}
 $sections+=@{start=[int]$s.Range.Start;end=[int]$s.Range.End;width=[double]$s.PageSetup.PageWidth;height=[double]$s.PageSetup.PageHeight;orientation=[int]$s.PageSetup.Orientation;oddEven=[int]$s.PageSetup.OddAndEvenPagesHeaderFooter;leftMargin=[double]$s.PageSetup.LeftMargin;rightMargin=[double]$s.PageSetup.RightMargin;topMargin=[double]$s.PageSetup.TopMargin;bottomMargin=[double]$s.PageSetup.BottomMargin;headers=$headers}}
 foreach($p in $d.Paragraphs){$paragraphs+=@{text=[string]$p.Range.Text;font=[string]$p.Range.Font.Name;fontSize=[double]$p.Range.Font.Size;page=[int]$p.Range.Information(3);x=[double]$p.Range.Information(5);y=[double]$p.Range.Information(6)}}
 foreach($s in $d.Shapes){$shapes+=@{name=[string]$s.Name;anchor=[int]$s.Anchor.Start;left=[double]$s.Left;top=[double]$s.Top;width=[double]$s.Width;height=[double]$s.Height}}
 foreach($n in $d.Footnotes){$notes+=@{text=[string]$n.Range.Text;anchor=[int]$n.Reference.Start}}
 return [ordered]@{pages=[int]$d.ComputeStatistics(2);sections=$sections;paragraphs=$paragraphs;shapes=$shapes;notes=$notes}
}
try {
 $app=New-Object -ComObject Word.Application
 $app.Visible=$false;$app.DisplayAlerts=0;$app.AutomationSecurity=3
 $links=$app.Options.UpdateLinksAtOpen;$app.Options.UpdateLinksAtOpen=$false
 $exePath=Join-Path ([string]$app.Path) 'WINWORD.EXE';$exeHash=(Get-FileHash -LiteralPath $exePath -Algorithm SHA256).Hash.ToLowerInvariant()
 if($exeHash -ne $wordBaseline.sha256){throw 'Installed Word differs from the pinned baseline.'}
 $fontHashes=@();foreach($fontName in @('arial.ttf','arialbd.ttf','arialbi.ttf','ariali.ttf')){$fontPath=Join-Path $env:WINDIR "Fonts/$fontName";$fontHashes+=@{file=$fontName;sha256=(Get-FileHash -LiteralPath $fontPath -Algorithm SHA256).Hash.ToLowerInvariant()}}
 $environment=@{executableSha256=$exeHash;fonts=$fontHashes;culture=[Globalization.CultureInfo]::CurrentCulture.Name;uiLanguage=[int]$app.LanguageSettings.LanguageID(2);printer=[string]$app.ActivePrinter}
 if($Create){
  $sourcePath=Join-Path $root 'source.docx';Write-Output 'Open authored fixture';$doc=$app.Documents.Open(([string]$sourcePath),$false,$false,$false)
  Write-Output "Native fixture ReadOnly=$($doc.ReadOnly), ProtectionType=$($doc.ProtectionType)"
  $first=$doc.Paragraphs.Item(1).Range.Duplicate;$first.End=$first.End-1;$first.Text='Section one edit.'
  $savePath=Join-Path $root 'native-expected.docx';Write-Output 'Save native expected edit';$doc.SaveAs2([ref]([string]$savePath),[ref]12)
  $doc.Close([ref]0);$doc=$null
  @{sourceHash=(Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash;expectedHash=(Get-FileHash -LiteralPath $savePath -Algorithm SHA256).Hash;action='Replace first paragraph text with Section one edit. using native Word Range.Text';version=[string]$app.Version;build=[string]$app.Build}|ConvertTo-Json|Set-Content -Encoding UTF8 (Join-Path $root 'native-edit.json')
  if($Render){$doc=$app.Documents.Open(([string]$savePath),$false,$true,$false);$pdfPath=[string](Join-Path $root 'native-expected.pdf');Write-Output 'Render native expected PDF';$doc.ExportAsFixedFormat($pdfPath,17);@{docxHash=(Get-FileHash -LiteralPath $savePath -Algorithm SHA256).Hash;pdfHash=(Get-FileHash -LiteralPath $pdfPath -Algorithm SHA256).Hash}|ConvertTo-Json|Set-Content -Encoding UTF8 (Join-Path $root 'native-pdf.json')}
  Write-Output 'Created native expected edit from the authored section/drawing/note fixture.'
 }else{
  $expectedPath=Join-Path $root 'native-expected.docx';$actualPath=Join-Path $root 'browser.docx'
  $receipt=Get-Content -LiteralPath (Join-Path $root 'native-edit.json') -Raw|ConvertFrom-Json
  if($receipt.sourceHash -ne (Get-FileHash -LiteralPath (Join-Path $root 'source.docx') -Algorithm SHA256).Hash -or $receipt.expectedHash -ne (Get-FileHash -LiteralPath $expectedPath -Algorithm SHA256).Hash){throw 'Native expected edit is stale; rerun -Create against the current source.'}
  $browserReceipt=Get-Content -LiteralPath (Join-Path $root 'browser-report.json') -Raw|ConvertFrom-Json
  if(!$browserReceipt.workflowPassed -or $browserReceipt.sourceHash -ne $receipt.sourceHash -or $browserReceipt.exportHash -ne (Get-FileHash -LiteralPath $actualPath -Algorithm SHA256).Hash){throw 'Browser export is stale or the workflow has not passed against this source.'}
  Write-Output 'Read native expected geometry';$doc=$app.Documents.Open(([string]$expectedPath),$false,$true,$false);$expected=Snapshot $doc;$doc.Close([ref]0);$doc=$null
  Write-Output 'Read browser export geometry';$doc=$app.Documents.Open(([string]$actualPath),$false,$true,$false);$actual=Snapshot $doc
  $doc.Close([ref]0);$doc=$null
  $mismatches=@();foreach($key in $expected.Keys){if(($expected[$key]|ConvertTo-Json -Depth 12 -Compress) -cne ($actual[$key]|ConvertTo-Json -Depth 12 -Compress)){$mismatches+= $key}}
  $hashes=@();foreach($name in @('source.docx','native-expected.docx','browser.docx')){$hashes+=@{path=$name;sha256=(Get-FileHash -LiteralPath (Join-Path $root $name) -Algorithm SHA256).Hash.ToLowerInvariant()}}
  @{observedAt=[DateTime]::UtcNow.ToString('o');version=[string]$app.Version;build=[string]$app.Build;browser=$browserReceipt.browser;environment=$environment;expected=$expected;actual=$actual;mismatches=$mismatches;hashes=$hashes}|ConvertTo-Json -Depth 16|Set-Content -Encoding UTF8 (Join-Path $root 'native-report.json')
  if($mismatches.Count){throw "Native Word mismatch: $mismatches"}
  Write-Output 'Native Word sections, paragraphs/page positions, anchored drawing and footnote match.'
  if($Render){$doc=$app.Documents.Open(([string]$actualPath),$false,$true,$false);$pdfPath=[string](Join-Path $root 'browser.pdf');Write-Output 'Render browser export PDF';$doc.ExportAsFixedFormat($pdfPath,17);@{docxHash=$browserReceipt.exportHash;pdfHash=(Get-FileHash -LiteralPath $pdfPath -Algorithm SHA256).Hash}|ConvertTo-Json|Set-Content -Encoding UTF8 (Join-Path $root 'browser-pdf.json')}
 }
}finally{
 try {if($doc){$doc.Close([ref]0)}} finally {
  if($app){try{if($null -ne $links){$app.Options.UpdateLinksAtOpen=$links};$app.Quit([ref]0)}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
 }
}

