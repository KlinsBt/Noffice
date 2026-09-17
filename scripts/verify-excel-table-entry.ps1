$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
Add-Type -AssemblyName System.IO.Compression.FileSystem
$app=$null;$source=$null;$actual=$null;$oldAuto=$null;$oldExpand=$null
function Invoke-EntryCall([scriptblock]$Action) {
 for($retry=0;$retry -lt 20;$retry++){try{return (& $Action)}catch{
  if($_.Exception.HResult -notin @(-2147418111,-2147417846,-2146777998) -or $retry -eq 19){throw}
  [ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250
 }}
}
function Read-EntryBook($book,[bool]$newBook=$false) {
 $ws=Invoke-EntryCall {$book.Worksheets.Item(1)};$cells=@();$columns=@();$totalFunctions=@();$table=$ws.ListObjects.Item(1)
 foreach($c in $table.ListColumns){$columns += [string]$c.Name;$totalFunctions += [int]$c.TotalsCalculation}
 $rows=if($newBook){5}else{8};$cols=if($newBook){2}else{8}
 for($r=1;$r -le $rows;$r++){for($c=1;$c -le $cols;$c++){
  $cell=Invoke-EntryCall {$ws.Cells.Item($r,$c)}
  $cells += @{ref=[string]$cell.Address();value=$cell.Value2;valueType=$(if($null -eq $cell.Value2){'empty'}else{$cell.Value2.GetType().FullName});formula=[string]$cell.Formula;format=[string]$cell.NumberFormat}
 }}
 $result=@{totalFunctions=$totalFunctions;columns=$columns;cells=$cells;name=[string]$table.DisplayName;ref=[string]$table.Range.Address();filter=[string]$table.AutoFilter.Range.Address();totals=[bool]$table.ShowTotals}
 if(-not $newBook){$result.note=[string]$ws.Range('F1').Comment.Text();$result.validation=[string]$ws.Range('K2').Validation.Formula1;$result.conditional=[string]$ws.Range('J2').FormatConditions.Item(1).Formula1;$result.names=@($book.Names|ForEach-Object {@{name=[string]$_.Name;formula=[string]$_.RefersTo}})}
 return $result
}
function Read-EntryMasters([string]$path) {
 $zip=[IO.Compression.ZipFile]::OpenRead($path)
 try {
  $entry=$zip.Entries|Where-Object {$_.FullName -match '^xl/tables/[^/]+\.xml$'}|Select-Object -First 1
  $reader=New-Object IO.StreamReader($entry.Open())
  try {$xml=[xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
  $formulas=@();foreach($c in $xml.SelectNodes('//*[local-name()="tableColumn"]')){
   $f=$c.SelectSingleNode('*[local-name()="calculatedColumnFormula"]')
   $t=$c.SelectSingleNode('*[local-name()="totalsRowFormula"]')
   $formulas += [ordered]@{calculated=$(if($f){[string]$f.InnerText}else{$null});total=$(if($t){[string]$t.InnerText}else{$null});label=$c.GetAttribute('totalsRowLabel');function=$(if($c.HasAttribute('totalsRowFunction')){$c.GetAttribute('totalsRowFunction')}else{'none'})}
  }
  return ,$formulas
 }finally{$zip.Dispose()}
}
try {
 $app=New-Object -ComObject Excel.Application
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=3;$app.EnableEvents=$false
 [uint32]$testProcessId=0;[void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$testProcessId)
 for($startup=0;$startup -lt 8;$startup++){[ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250}
 $oldExpand=Invoke-EntryCall {$app.AutoCorrect.AutoExpandListRange};Invoke-EntryCall {$app.AutoCorrect.AutoExpandListRange=$true};$oldAuto=Invoke-EntryCall {$app.AutoCorrect.AutoFillFormulasInLists};Invoke-EntryCall {$app.AutoCorrect.AutoFillFormulasInLists=$true}
 $root=Join-Path (Get-Location) '.local/xlsx-table-entry';$sourcePath=Join-Path $root 'source.xlsx'
 $source=Invoke-EntryCall {$app.Workbooks.Open($sourcePath,0,$true)};$ws=Invoke-EntryCall {$source.Worksheets.Item(1)}
 $results=@();$mismatches=@();$snapshots=0
 foreach($stage in @('first','quantity','next','formula','exception','neighbor')) {
  switch($stage){
   'first' {Invoke-EntryCall {$ws.Range('A5').Value2='East'}}
   'quantity' {Invoke-EntryCall {$ws.Range('B5').Value2=7}}
   'next' {Invoke-EntryCall {$ws.Range('B6').Value2=11}}
   'formula' {Invoke-EntryCall {$ws.Range('B7').Formula='=2+3'}}
   'exception' {Invoke-EntryCall {$ws.Range('C7').Value2=99}}
   'neighbor' {Invoke-EntryCall {$app.AutoCorrect.AutoExpandListRange=$false;$ws.Range('C8').Value2=55;$app.AutoCorrect.AutoExpandListRange=$true;$ws.Range('B8').Value2=13}}
  }
  Invoke-EntryCall {$app.CalculateFull()};$isNew=$stage.StartsWith('new-');$expected=Read-EntryBook $source $isNew
  $nativePath=Join-Path $root "native-$stage.xlsx";Invoke-EntryCall {$source.SaveCopyAs($nativePath)}
  $expected.masters=Read-EntryMasters $nativePath
  $path=Join-Path $root "$stage.xlsx";$actual=Invoke-EntryCall {$app.Workbooks.Open($path,0,$true)}
  Invoke-EntryCall {$app.CalculateFull()};$observed=Read-EntryBook $actual $isNew;$observed.masters=Read-EntryMasters $path
  foreach($key in @('columns','name','ref','filter','totals','totalFunctions','masters','note','validation','conditional','names')){
   if(($expected[$key]|ConvertTo-Json -Depth 8 -Compress) -cne ($observed[$key]|ConvertTo-Json -Depth 8 -Compress)){$mismatches += @{stage=$stage;part=$key;expected=$expected[$key];actual=$observed[$key]}}
  }
  for($i=0;$i -lt $expected.cells.Count;$i++){foreach($key in @('value','valueType','formula','format')){
   if([string]$expected.cells[$i][$key] -cne [string]$observed.cells[$i][$key]){$mismatches += @{stage=$stage;ref=$expected.cells[$i].ref;part=$key;expected=$expected.cells[$i][$key];actual=$observed.cells[$i][$key]}}
  }}
  $snapshots += $expected.cells.Count
  $results += @{stage=$stage;sha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $path).Hash;nativeSha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $nativePath).Hash;expected=$expected;actual=$observed}
  Invoke-EntryCall {$actual.Close($false)};$actual=$null
 }
 @{normalization='Missing totalsRowFunction is compared as none; native COM TotalsCalculation is compared independently.';excelVersion=[string]$app.Version;excelBuild=[string]$app.Build;sourceSha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $sourcePath).Hash;stages=$results;cellSnapshots=$snapshots;mismatches=$mismatches} | ConvertTo-Json -Depth 16 | Set-Content -Encoding UTF8 (Join-Path $root 'native-report.json')
 Write-Output "Native Excel table entry: $($mismatches.Count) mismatches across $snapshots cell snapshots and six table/master-formula comparisons."
 if($mismatches.Count){throw 'Table entry comparison failed; see native-report.json.'}
} finally {
 if($actual){Invoke-EntryCall {$actual.Close($false)}};if($source){Invoke-EntryCall {$source.Close($false)}}
 if($app){if($null -ne $oldExpand){Invoke-EntryCall {$app.AutoCorrect.AutoExpandListRange=$oldExpand}};if($null -ne $oldAuto){Invoke-EntryCall {$app.AutoCorrect.AutoFillFormulasInLists=$oldAuto}};Invoke-EntryCall {$app.Quit()}}
}
