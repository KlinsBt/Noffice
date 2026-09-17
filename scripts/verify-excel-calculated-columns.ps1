$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
Add-Type -AssemblyName System.IO.Compression.FileSystem
$app=$null;$source=$null;$actual=$null;$oldAuto=$null
function Invoke-ColumnCall([scriptblock]$Action) {
 for($retry=0;$retry -lt 20;$retry++){try{return (& $Action)}catch{
  if($_.Exception.HResult -notin @(-2147418111,-2147417846,-2146777998) -or $retry -eq 19){throw}
  [ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250
 }}
}
function Read-ColumnBook($book,[int]$rows=6,[int]$cols=7) {
 $ws=Invoke-ColumnCall {$book.Worksheets.Item(1)};$cells=@();$columns=@()
 foreach($c in $ws.ListObjects.Item(1).ListColumns){$columns += [string]$c.Name}
 for($r=1;$r -le $rows;$r++){for($c=1;$c -le $cols;$c++){
  $cell=Invoke-ColumnCall {$ws.Cells.Item($r,$c)}
  $cells += @{ref=[string]$cell.Address();value=$cell.Value2;formula=[string]$cell.Formula;format=[string]$cell.NumberFormat}
 }}
 return @{columns=$columns;cells=$cells;name=[string]$ws.ListObjects.Item(1).DisplayName;hidden=[bool]$ws.Rows.Item(4).Hidden}
}
function Read-ColumnMasters([string]$path) {
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
 $oldAuto=Invoke-ColumnCall {$app.AutoCorrect.AutoFillFormulasInLists}
 Invoke-ColumnCall {$app.AutoCorrect.AutoFillFormulasInLists=$true}
 $root=Join-Path (Get-Location) '.local/xlsx-calculated-columns';$sourcePath=Join-Path $root 'source.xlsx'
 $source=Invoke-ColumnCall {$app.Workbooks.Open($sourcePath,0,$true)};$ws=Invoke-ColumnCall {$source.Worksheets.Item(1)}
 $results=@();$mismatches=@();$snapshots=0
 foreach($stage in @('structured','updated','renamed','exception','isolated','a1','a1updated','new')) {
  switch($stage){
   'structured' {Invoke-ColumnCall {$ws.Range('C3').Formula='=Sales[[#This Row],[Quantity]]*10'}}
   'updated' {Invoke-ColumnCall {$ws.Range('C2').Formula='=Sales[[#This Row],[Quantity]]*20'}}
   'renamed' {Invoke-ColumnCall {$ws.Range('B1').Value2='Units'}}
   'exception' {Invoke-ColumnCall {$ws.Range('C3').Value2=99}}
   'isolated' {Invoke-ColumnCall {$ws.Range('C2').Formula='=Sales[[#This Row],[Units]]*30'}}
   'a1' {Invoke-ColumnCall {[void]$ws.Range('C2:C4').ClearContents();$ws.Range('C3').Formula='=B3*10+$B$2'}}
   'a1updated' {Invoke-ColumnCall {$ws.Range('C2').Formula='=B2*20+$B$2'}}
   'new' {
    Invoke-ColumnCall {$source.Close($false)};$source=Invoke-ColumnCall {$app.Workbooks.Add()};$ws=Invoke-ColumnCall {$source.Worksheets.Item(1)}
    Invoke-ColumnCall {$ws.Range('A1').Value2='Quantity';$ws.Range('B1').Value2='Total';$ws.Range('A2').Value2=2;$ws.Range('A3').Value2=3;$ws.Range('A4').Value2=5}
    $t=Invoke-ColumnCall {$ws.ListObjects.Add(1,$ws.Range('A1:B4'),$null,1)}
    Invoke-ColumnCall {$t.Name='Sales';$ws.Range('B3').Formula='=A3*10';$ws.Range('B4').Value2=99;$ws.Range('B3').Formula='=A3*20'}
   }
  }
  Invoke-ColumnCall {$app.CalculateFull()}
  $expected=if($stage -eq 'new'){Read-ColumnBook $source 4 2}else{Read-ColumnBook $source}
  $nativePath=Join-Path $root "native-$stage.xlsx"
  if($stage -eq 'new'){
   Invoke-ColumnCall {$source.SaveAs($nativePath,51);$source.Close($false)};$source=$null
  }else{Invoke-ColumnCall {$source.SaveCopyAs($nativePath)}}
  $expected.masters=Read-ColumnMasters $nativePath
  $path=Join-Path $root "$stage.xlsx";$actual=Invoke-ColumnCall {$app.Workbooks.Open($path,0,$true)}
  Invoke-ColumnCall {$app.CalculateFull()}
  $observed=if($stage -eq 'new'){Read-ColumnBook $actual 4 2}else{Read-ColumnBook $actual}
  $observed.masters=Read-ColumnMasters $path
  foreach($key in @('columns','name','hidden','masters')){
   if(($expected[$key]|ConvertTo-Json -Depth 8 -Compress) -cne ($observed[$key]|ConvertTo-Json -Depth 8 -Compress)){$mismatches += @{stage=$stage;part=$key;expected=$expected[$key];actual=$observed[$key]}}
  }
  for($i=0;$i -lt $expected.cells.Count;$i++){foreach($key in @('value','formula','format')){
   if([string]$expected.cells[$i][$key] -cne [string]$observed.cells[$i][$key]){$mismatches += @{stage=$stage;ref=$expected.cells[$i].ref;part=$key;expected=$expected.cells[$i][$key];actual=$observed.cells[$i][$key]}}
  }}
  $snapshots += $expected.cells.Count
  $results += @{stage=$stage;sha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $path).Hash;nativeSha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $nativePath).Hash;expected=$expected;actual=$observed}
  Invoke-ColumnCall {$actual.Close($false)};$actual=$null
 }
 @{excelVersion=[string]$app.Version;excelBuild=[string]$app.Build;sourceSha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $sourcePath).Hash;stages=$results;cellSnapshots=$snapshots;mismatches=$mismatches} | ConvertTo-Json -Depth 16 | Set-Content -Encoding UTF8 (Join-Path $root 'native-report.json')
 Write-Output "Native Excel calculated columns: $($mismatches.Count) mismatches across $snapshots cell snapshots and eight master-formula comparisons."
 if($mismatches.Count){throw 'Calculated-column comparison failed; see native-report.json.'}
} finally {
 if($actual){Invoke-ColumnCall {$actual.Close($false)}};if($source){Invoke-ColumnCall {$source.Close($false)}}
 if($app){if($null -ne $oldAuto){Invoke-ColumnCall {$app.AutoCorrect.AutoFillFormulasInLists=$oldAuto}};Invoke-ColumnCall {$app.Quit()}}
}
