$ErrorActionPreference = 'Stop'
$app=$null;$source=$null;$actual=$null
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
function Invoke-StructureCall([scriptblock]$Action) {
 for($retry=0;$retry -lt 20;$retry++) {
  try {return (& $Action)} catch {
   if($_.Exception.HResult -notin @(-2147418111,-2147417846,-2146777998) -or $retry -eq 19){throw}
   [ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250
  }
 }
}
function Read-StructureSnapshot($book) {
 $cells=@();$sheets=@()
 foreach($name in @('Data',"O'Brien")) {
  $ws=Invoke-StructureCall {$book.Worksheets.Item($name)}
  for($r=1;$r -le 9;$r++) {for($c=1;$c -le 7;$c++) {
   $cell=Invoke-StructureCall {$ws.Cells.Item($r,$c)}
   $formula=if(Invoke-StructureCall {$cell.HasFormula}){[string](Invoke-StructureCall {$cell.Formula})}else{''}
   $value=Invoke-StructureCall {$cell.Value2}
   if($value -is [int]){$value=[string](Invoke-StructureCall {$cell.Text})}
   $cells += [ordered]@{sheet=$name;row=$r;column=$c;formula=$formula;value=$value;numberFormat=[string](Invoke-StructureCall {$cell.NumberFormat});bold=[bool](Invoke-StructureCall {$cell.Font.Bold});fill=[int](Invoke-StructureCall {$cell.Interior.ColorIndex});merge=[string](Invoke-StructureCall {$cell.MergeArea.Address()})}
  }}
  $rowDimensions=@();$columnDimensions=@();$validation=@()
  if($name -eq 'Data') {
   for($i=1;$i -le 9;$i++){$rowDimensions += [ordered]@{height=[double](Invoke-StructureCall {$ws.Rows.Item($i).RowHeight});hidden=[bool](Invoke-StructureCall {$ws.Rows.Item($i).Hidden})}}
   for($i=1;$i -le 7;$i++){$columnDimensions += [ordered]@{width=[double](Invoke-StructureCall {$ws.Columns.Item($i).ColumnWidth});hidden=[bool](Invoke-StructureCall {$ws.Columns.Item($i).Hidden})}}
   foreach($ref in @('A2','A3','A4','A5','A6','A7','B2','B3')) {
    try {$rule=Invoke-StructureCall {$ws.Range($ref).Validation.Type}}catch{$rule=0}
    $validation += [ordered]@{ref=$ref;type=$rule;formula1=$(if($rule){[string](Invoke-StructureCall {$ws.Range($ref).Validation.Formula1})}else{''});formula2=$(if($rule){[string](Invoke-StructureCall {$ws.Range($ref).Validation.Formula2})}else{''})}
   }
  }
  $sheets += [ordered]@{name=$name;printArea=[string](Invoke-StructureCall {$ws.PageSetup.PrintArea});printTitles=[string](Invoke-StructureCall {$ws.PageSetup.PrintTitleRows});rowDimensions=$rowDimensions;columnDimensions=$columnDimensions;validation=$validation}
 }
 return [ordered]@{cells=$cells;sheets=$sheets;amounts=[string](Invoke-StructureCall {$book.Names.Item('Amounts').RefersTo});link=[string](Invoke-StructureCall {$book.Worksheets.Item("O'Brien").Range('A5').Hyperlinks.Item(1).SubAddress})}
}
try {
 $app=New-Object -ComObject Excel.Application
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=3;$app.EnableEvents=$false
 [uint32]$testProcessId=0
 [void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$testProcessId)
 for($startup=0;$startup -lt 8;$startup++){[ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250}
 $root=Join-Path (Get-Location) '.local/xlsx-structure'
 $sourcePath=Join-Path $root 'source.xlsx'
 $source=Invoke-StructureCall {$app.Workbooks.Open($sourcePath,0,$true)}
 $ws=Invoke-StructureCall {$source.Worksheets.Item('Data')}
 $results=@();$mismatches=@()
 foreach($stage in @('insert-row','insert-column','delete-row','edited','insert-first-row','insert-first-column','delete-first-row','delete-first-column')) {
  switch($stage) {
   'insert-row' {[void](Invoke-StructureCall {$ws.Range('3:3').Insert()})}
   'insert-column' {[void](Invoke-StructureCall {$ws.Range('B:B').Insert()})}
   'delete-row' {[void](Invoke-StructureCall {$ws.Range('4:4').Delete()})}
   'edited' {[void](Invoke-StructureCall {$ws.Range('D:D').Delete()});Invoke-StructureCall {$ws.Range('A2').Value2=[double]12}}
   'insert-first-row' {[void](Invoke-StructureCall {$ws.Range('1:1').Insert()})}
   'insert-first-column' {[void](Invoke-StructureCall {$ws.Range('A:A').Insert()})}
   'delete-first-row' {[void](Invoke-StructureCall {$ws.Range('1:1').Delete()})}
   'delete-first-column' {[void](Invoke-StructureCall {$ws.Range('A:A').Delete()})}
  }
  Invoke-StructureCall {$app.CalculateFull()}
  $expected=Read-StructureSnapshot $source
  $path=Join-Path $root "$stage.xlsx"
  $actual=Invoke-StructureCall {$app.Workbooks.Open($path,0,$true)}
  Invoke-StructureCall {$app.CalculateFull()}
  $observed=Read-StructureSnapshot $actual
  for($i=0;$i -lt $expected.cells.Count;$i++) {
   foreach($key in @('formula','value','numberFormat','bold','fill','merge')) {
    if([string]$expected.cells[$i][$key] -cne [string]$observed.cells[$i][$key]) {
     $mismatches += [ordered]@{stage=$stage;sheet=$expected.cells[$i].sheet;row=$expected.cells[$i].row;column=$expected.cells[$i].column;field=$key;expected=$expected.cells[$i][$key];actual=$observed.cells[$i][$key]}
    }
   }
  }
  foreach($key in @('amounts','link')) {if($expected[$key] -cne $observed[$key]){$mismatches += [ordered]@{stage=$stage;field=$key;expected=$expected[$key];actual=$observed[$key]}}}
  if(($expected.sheets|ConvertTo-Json -Depth 8 -Compress) -cne ($observed.sheets|ConvertTo-Json -Depth 8 -Compress)){$mismatches += [ordered]@{stage=$stage;field='sheetMetadata';expected=$expected.sheets;actual=$observed.sheets}}
  $results += [ordered]@{stage=$stage;sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();expected=$expected;actual=$observed}
  Invoke-StructureCall {$actual.Close($false)};$actual=$null
 }
 [ordered]@{application='Microsoft Excel';version=[string]$app.Version;build=[string]$app.Build;sourceSha256=(Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToLowerInvariant();stages=$results;mismatches=$mismatches}|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $root 'native-report.json') -Encoding UTF8
 if($mismatches.Count){throw "$($mismatches.Count) differences from native Excel. See .local/xlsx-structure/native-report.json."}
 Write-Output 'Native Excel matched 1008 cell snapshots across eight structural stages, including formulas, calculated values, cell formatting and merges; names, links, dimensions, validation and print settings match.'
} finally {
 if($actual){try{Invoke-StructureCall {$actual.Close($false)}}catch{Write-Warning $_}}
 if($source){try{Invoke-StructureCall {$source.Close($false)}}catch{Write-Warning $_}}
 if($app){try{Invoke-StructureCall {$app.Quit()}}catch{Write-Warning $_}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
}
