$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
Add-Type -AssemblyName System.IO.Compression.FileSystem
$app=$null;$source=$null;$actual=$null;$oldAuto=$null
function Invoke-ResizeCall([scriptblock]$Action) {
 for($retry=0;$retry -lt 20;$retry++){try{return (& $Action)}catch{
  if($_.Exception.HResult -notin @(-2147418111,-2147417846,-2146777998) -or $retry -eq 19){throw}
  [ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250
 }}
}
function Read-ResizeBook($book,[bool]$newBook=$false) {
 $ws=Invoke-ResizeCall {$book.Worksheets.Item(1)};$cells=@();$columns=@();$table=$ws.ListObjects.Item(1)
 foreach($c in $table.ListColumns){$columns += [string]$c.Name}
 $rows=if($newBook){5}else{8};$cols=if($newBook){2}else{8}
 for($r=1;$r -le $rows;$r++){for($c=1;$c -le $cols;$c++){
  $cell=Invoke-ResizeCall {$ws.Cells.Item($r,$c)}
  $cells += @{ref=[string]$cell.Address();value=$cell.Value2;formula=[string]$cell.Formula;format=[string]$cell.NumberFormat}
 }}
 $result=@{columns=$columns;cells=$cells;name=[string]$table.DisplayName;ref=[string]$table.Range.Address();filter=[string]$table.AutoFilter.Range.Address()}
 if(-not $newBook){$result.note=[string]$ws.Range('F1').Comment.Text();$result.validation=[string]$ws.Range('K2').Validation.Formula1;$result.conditional=[string]$ws.Range('J2').FormatConditions.Item(1).Formula1;$result.names=@($book.Names|ForEach-Object {@{name=[string]$_.Name;formula=[string]$_.RefersTo}})}
 return $result
}
function Read-ResizeMasters([string]$path) {
 $zip=[IO.Compression.ZipFile]::OpenRead($path)
 try {
  $entry=$zip.Entries|Where-Object {$_.FullName -match '^xl/tables/[^/]+\.xml$'}|Select-Object -First 1
  $reader=New-Object IO.StreamReader($entry.Open())
  try {$xml=[xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
  $formulas=@();foreach($c in $xml.SelectNodes('//*[local-name()="tableColumn"]')){
   $f=$c.SelectSingleNode('*[local-name()="calculatedColumnFormula"]')
   $formulas += $(if($f){[string]$f.InnerText}else{$null})
  }
  return ,$formulas
 }finally{$zip.Dispose()}
}
try {
 $app=New-Object -ComObject Excel.Application
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=3;$app.EnableEvents=$false
 [uint32]$testProcessId=0;[void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$testProcessId)
 for($startup=0;$startup -lt 8;$startup++){[ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250}
 $oldAuto=Invoke-ResizeCall {$app.AutoCorrect.AutoFillFormulasInLists};Invoke-ResizeCall {$app.AutoCorrect.AutoFillFormulasInLists=$true}
 $root=Join-Path (Get-Location) '.local/xlsx-table-resize';$sourcePath=Join-Path $root 'source.xlsx'
 $source=Invoke-ResizeCall {$app.Workbooks.Open($sourcePath,0,$true)};$ws=Invoke-ResizeCall {$source.Worksheets.Item(1)}
 $results=@();$mismatches=@();$snapshots=0
 foreach($stage in @('grown','shrunk','regrown','cleared','empty-grown','new-grown','new-shrunk')) {
  switch($stage){
   'grown' {Invoke-ResizeCall {$ws.Range('C3').Value2=99;$ws.ListObjects.Item(1).Resize($ws.Range('A1:C6'))}}
   'shrunk' {Invoke-ResizeCall {$ws.ListObjects.Item(1).Resize($ws.Range('A1:C3'))}}
   'regrown' {Invoke-ResizeCall {$ws.ListObjects.Item(1).Resize($ws.Range('A1:C6'))}}
   'cleared' {Invoke-ResizeCall {[void]$ws.Range('C2:C6').ClearContents()}}
   'empty-grown' {Invoke-ResizeCall {$ws.ListObjects.Item(1).Resize($ws.Range('A1:C8'))}}
   'new-grown' {
    Invoke-ResizeCall {$source.Close($false)};$source=Invoke-ResizeCall {$app.Workbooks.Add()};$ws=Invoke-ResizeCall {$source.Worksheets.Item(1)}
    Invoke-ResizeCall {$ws.Range('A1').Value2='Quantity';$ws.Range('B1').Value2='Total';$ws.Range('A2').Value2=2;$ws.Range('A3').Value2=3}
    $t=Invoke-ResizeCall {$ws.ListObjects.Add(1,$ws.Range('A1:B3'),$null,1)}
    Invoke-ResizeCall {$t.Name='Sales';$ws.Range('B2').Formula='=A2*10+$A$2';$t.Resize($ws.Range('A1:B5'));$ws.Range('A4').Value2=7;$ws.Range('A5').Value2=11}
    # New workbooks inherit this machine's OpenDocument default; choose XLSX explicitly.
    $scratchPath=Join-Path $root 'native-scratch.xlsx';Invoke-ResizeCall {$source.SaveAs($scratchPath,51)}
   }
   'new-shrunk' {Invoke-ResizeCall {$ws.ListObjects.Item(1).Resize($ws.Range('A1:B2'))}}
  }
  Invoke-ResizeCall {$app.CalculateFull()};$isNew=$stage.StartsWith('new-');$expected=Read-ResizeBook $source $isNew
  $nativePath=Join-Path $root "native-$stage.xlsx";Invoke-ResizeCall {$source.SaveCopyAs($nativePath)}
  $expected.masters=Read-ResizeMasters $nativePath
  $path=Join-Path $root "$stage.xlsx";$actual=Invoke-ResizeCall {$app.Workbooks.Open($path,0,$true)}
  Invoke-ResizeCall {$app.CalculateFull()};$observed=Read-ResizeBook $actual $isNew;$observed.masters=Read-ResizeMasters $path
  foreach($key in @('columns','name','ref','filter','masters','note','validation','conditional','names')){
   if(($expected[$key]|ConvertTo-Json -Depth 8 -Compress) -cne ($observed[$key]|ConvertTo-Json -Depth 8 -Compress)){$mismatches += @{stage=$stage;part=$key;expected=$expected[$key];actual=$observed[$key]}}
  }
  for($i=0;$i -lt $expected.cells.Count;$i++){foreach($key in @('value','formula','format')){
   if([string]$expected.cells[$i][$key] -cne [string]$observed.cells[$i][$key]){$mismatches += @{stage=$stage;ref=$expected.cells[$i].ref;part=$key;expected=$expected.cells[$i][$key];actual=$observed.cells[$i][$key]}}
  }}
  $snapshots += $expected.cells.Count
  $results += @{stage=$stage;sha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $path).Hash;nativeSha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $nativePath).Hash;expected=$expected;actual=$observed}
  Invoke-ResizeCall {$actual.Close($false)};$actual=$null
 }
 @{excelVersion=[string]$app.Version;excelBuild=[string]$app.Build;sourceSha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $sourcePath).Hash;stages=$results;cellSnapshots=$snapshots;mismatches=$mismatches} | ConvertTo-Json -Depth 16 | Set-Content -Encoding UTF8 (Join-Path $root 'native-report.json')
 Write-Output "Native Excel table resizing: $($mismatches.Count) mismatches across $snapshots cell snapshots and seven table/master-formula comparisons."
 if($mismatches.Count){throw 'Table resize comparison failed; see native-report.json.'}
} finally {
 if($actual){Invoke-ResizeCall {$actual.Close($false)}};if($source){Invoke-ResizeCall {$source.Close($false)}}
 if($app){if($null -ne $oldAuto){Invoke-ResizeCall {$app.AutoCorrect.AutoFillFormulasInLists=$oldAuto}};Invoke-ResizeCall {$app.Quit()}}
}
