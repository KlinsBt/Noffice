param([switch]$Compare)
# Authored shaped-reference/SUMPRODUCT contract; compare real browser exports with native edits.
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root '.local/excel-sumproduct'
[void][IO.Directory]::CreateDirectory($directory)
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
function Hash($path){(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant()}
function WriteJson($name,$value){[IO.File]::WriteAllText((Join-Path $directory $name),($value|ConvertTo-Json -Depth 16),(New-Object Text.UTF8Encoding($false)))}
$formulas=[ordered]@{
 A1='=SUMPRODUCT(Data!A1:A3,Data!B1:B3)'
 A2='=SUMPRODUCT(Data!A1:C2,Data!A2:C3)'
 A3='=SUMPRODUCT(Data!A1:C2,Data!A1:B3)'
 A4='=SUMPRODUCT(Data!A1:A3,Data!A1:C1)'
 A5='=SUMPRODUCT(Data!A1:A3)'
 A6='=SUMPRODUCT(Data!A1:A3,2)'
 A7='=SUMPRODUCT(Data!A1:A3*2)'
 A8='=SUMPRODUCT((Data!A1:A3>1)*Data!B1:B3)'
 A9='=SUMPRODUCT(--(Data!A1:A3>1),Data!B1:B3)'
 A10='=SUMPRODUCT(Data!A1:A3>1,Data!B1:B3)'
 A11='=SUMPRODUCT(Data!A1:A3+Data!B1:B3)'
 A12='=SUMPRODUCT(Data!A1:A3^2)'
 B1='=SUMPRODUCT(Data!A1:C1*Data!A1:A3)'
 B2='=SUMPRODUCT(Data!A1:C2*Data!A1:B3)'
 B3='=SUMPRODUCT(Data!E1:E4,Data!F1:F4)'
 B4='=SUMPRODUCT(Data!E1:E4*Data!F1:F4)'
 B5='=SUMPRODUCT(Data!G1:G3,Data!F1:F3)'
 B6='=SUMPRODUCT(Data!G1:G3*Data!F1:F3)'
 B7='=SUMPRODUCT(Data!E1:E5,Data!F1:F5)'
 B8='=SUMPRODUCT(0*Data!E1:E5)'
 B9='=SUMPRODUCT(2,3)'
 B10='=SUMPRODUCT(TRUE,3)'
 B11='=SUMPRODUCT("2",3)'
 B12='=SUMPRODUCT({1,2;3,4},{5,6;7,8})'
 C1='=SUMPRODUCT({1,2,3;4,5,6},{1,2;3,4;5,6})'
 C2='=SUMPRODUCT({1,"2",TRUE}, {2,3,4})'
 C3='=SUMPRODUCT({1,"2",TRUE}*{2,3,4})'
 C4='=SUMPRODUCT(Matrix,Data!A2:C3)'
 C5='=SUMPRODUCT(Matrix*2)'
 C6='=SUMPRODUCT(Items[Qty],Items[Price])'
 C7='=SUMPRODUCT(Items[Qty]*Items[Price])'
 C8='=SUMPRODUCT(INDEX(Data!A1:C3,0,2),Data!A1:A3)'
 C9='=SUMPRODUCT(INDIRECT("Data!A1:A3"),Data!B1:B3)'
 C10='=IFERROR(SUMPRODUCT(Data!E1:E5),99)'
 C11='=SUMPRODUCT(IFERROR(Data!E1:E5,0),Data!F1:F5)'
 C12='=SUMPRODUCT(IF(Data!A1:A3>1,Data!B1:B3,0))'
 D1='=SUMPRODUCT(Data!A1:A3/Data!G1:G3)'
 D2='=SUMPRODUCT(-Data!A1:A3)'
 D3='=SUMPRODUCT(Data!A1:A3%)'
 D4='=SUMPRODUCT(Data!A1:A3&"x")'
 D5='=SUMPRODUCT({1;2;3}+{10,20})'
 D6='=SUMPRODUCT(Data!A1:C2,Data!A2:C3,Data!A1:C2)'
 D7='=SUMPRODUCT(Data!A1:A3=4)'
 D8='=SUMPRODUCT(--(Data!A1:A3=4))'
 D9='=SUMPRODUCT(Data!E4*Data!F4)'
 D10='=SUMPRODUCT(Data!G2*Data!F2)'
 D11='=SUMPRODUCT({1,#N/A},{0,0})'
 D12='=SUMPRODUCT({1;2}+{10;20;30})'
}
function Snapshot($book){
 $rows=@();$sheet=$book.Worksheets.Item('Calls')
 foreach($key in $formulas.Keys){
  $cell=$sheet.Range($key);$value=$cell.Value2;$error=$app.WorksheetFunction.IsError($cell)
  if($error){$code=[int]([long]$value -band 65535);$value=@{2000='#NULL!';2007='#DIV/0!';2015='#VALUE!';2023='#REF!';2029='#NAME?';2036='#NUM!';2042='#N/A'}[$code];if($null -eq $value){throw "Unknown native error $code"}}
  $rows+=,[ordered]@{ref=$key;formula=[string]$cell.Formula;value=$value;kind=$(if($error){'error'}else{'value'})}
 };return ,$rows
}
function SanitizeFixture($path){
 # Excel overwrites the author properties when saving. Sanitize the closed owned
 # package before capturing its reference hash; do not alter Office user settings.
 Add-Type -AssemblyName System.IO.Compression
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $zip=[IO.Compression.ZipFile]::Open($path,[IO.Compression.ZipArchiveMode]::Update)
 try{
  foreach($part in @('docProps/core.xml','docProps/app.xml')){
   $entry=$zip.GetEntry($part);if(!$entry){continue}
   $reader=New-Object IO.StreamReader($entry.Open())
   try{[xml]$xml=$reader.ReadToEnd()}finally{$reader.Dispose()}
   foreach($node in $xml.SelectNodes('//*[local-name()="creator" or local-name()="lastModifiedBy" or local-name()="Company" or local-name()="Manager"]')){$node.InnerText='Noffice'}
   $entry.Delete();$replacement=$zip.CreateEntry($part)
   $writer=New-Object IO.StreamWriter($replacement.Open(),(New-Object Text.UTF8Encoding($false)))
   try{$writer.Write($xml.OuterXml)}finally{$writer.Dispose()}
  }
 }finally{$zip.Dispose()}
}
$app=$null;$book=$null
try{
 $app=New-Object -ComObject Excel.Application
 if($app.Workbooks.Count -ne 0){throw 'This probe requires an idle owned Excel instance.'}
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.AskToUpdateLinks
 $app.AutomationSecurity=3;$app.DisplayAlerts=$false;$app.AskToUpdateLinks=$false
 $baseline=Get-Content -Raw (Join-Path $root 'docs/parity/baseline.json')|ConvertFrom-Json
 $native=$baseline.applications|Where-Object executable -eq 'EXCEL.EXE'
 if((Hash (Join-Path ([string]$app.Path) 'EXCEL.EXE')) -ne $native.sha256){throw 'The running Excel baseline changed.'}
 [uint32]$owned=0;[void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$owned);[ExcelTestWindow]::DismissReminder($owned)
 if(!$Compare){
 $book=$app.Workbooks.Add();while($book.Worksheets.Count -lt 2){[void]$book.Worksheets.Add()}
 $book.Worksheets.Item(1).Name='Data';$book.Worksheets.Item(2).Name='Calls'
 $data=$book.Worksheets.Item('Data');$calls=$book.Worksheets.Item('Calls')
 foreach($row in 1..3){foreach($col in 1..3){$data.Cells.Item($row,$col).Value2=[double](($row-1)*3+$col)}}
 $data.Range('E1').Value2=[double]2;$data.Range('E2').NumberFormat='@';$data.Range('E2').Value2='3';$data.Range('E3').Formula='=TRUE()';$data.Range('E5').Formula='=1/0'
 foreach($row in 1..5){$data.Range("F$row").Value2=[double]($row*10)}
 $data.Range('G1').Value2='x';$data.Range('G2').Formula='=""';$data.Range('G3').Value2=[double]0
 $data.Range('H1').Value2='Qty';$data.Range('I1').Value2='Price'
 foreach($row in 2..4){$data.Range("H$row").Value2=[double]($row-1);$data.Range("I$row").Value2=[double](($row-1)*10)}
 $table=$data.ListObjects.Add(1,$data.Range('H1:I4'),$null,1);$table.Name='Items'
 [void]$book.Names.Add('Matrix','=Data!$A$1:$C$2')
 foreach($key in $formulas.Keys){$calls.Range($key).Formula=$formulas[$key]}
 $app.CalculateFullRebuild();$initial=Snapshot $book
 $source=Join-Path $directory 'source.xlsx';$book.SaveAs($source,51)
 $stages=@()
 foreach($stage in @('edit','delete','error','formula','table')){
  $data.Range('A2').Value2=[double]4
  if($stage -eq 'edit'){$data.Range('A2').Value2=[double]64}
  elseif($stage -eq 'delete'){[void]$data.Range('A2').ClearContents()}
  elseif($stage -eq 'error'){$data.Range('A2').Formula='=1/0'}
  elseif($stage -eq 'formula'){$calls.Range('D8').Formula='=SUMPRODUCT((Data!A1:A3>=4)*Data!B1:B3)'}
  elseif($stage -eq 'table'){$calls.Range('D8').Formula=$formulas.D8;$data.Range('H3').Value2=[double]4}
  $app.CalculateFullRebuild();$stages+=,[ordered]@{stage=$stage;cells=(Snapshot $book)}
  $book.SaveCopyAs((Join-Path $directory "native-$stage.xlsx"))
 }
 $book.Close($false);$book=$null
 SanitizeFixture $source
 $book=$app.Workbooks.Open($source,0,$true);$app.CalculateFullRebuild();$reopened=Snapshot $book
 if(($initial|ConvertTo-Json -Compress) -cne ($reopened|ConvertTo-Json -Compress)){throw 'Native saved-file array results changed on reopening.'}
 WriteJson 'reference.json' ([ordered]@{scope='Bounded shaped-reference/SUMPRODUCT values; full array/function/spill semantics remain open';scriptSha256=Hash $PSCommandPath;sourceSha256=Hash $source;executableSha256=$native.sha256;locale=$app.LanguageSettings.LanguageID(2);initial=$initial;stages=$stages;reopened=$reopened})
 $book.Close($false);$book=$null
  Copy-Item -LiteralPath $source -Destination (Join-Path $root 'tests/fixtures/excel-sumproduct.xlsx')
  Copy-Item -LiteralPath (Join-Path $directory 'reference.json') -Destination (Join-Path $root 'tests/fixtures/native-excel-sumproduct.json')
 }else{
  $fixture=Join-Path $root 'tests/fixtures/excel-sumproduct.xlsx'
  $browser=Get-Content -Raw (Join-Path $directory 'browser-report.json')|ConvertFrom-Json
  if(!$browser.passed -or (Hash $fixture) -ne $browser.sourceSha256){throw 'Stale browser source evidence.'}
  $reports=@();$keeper=$app.Workbooks.Add()
  foreach($stage in @('edit','delete','error','formula','table')){
   $output=Join-Path $directory "$stage.xlsx"
   if((Hash $output) -ne $browser.exports.$stage){throw 'Stale browser export evidence.'}
   $book=$app.Workbooks.Open($fixture,0,$false);$data=$book.Worksheets.Item('Data')
   if($stage -eq 'edit'){$data.Range('A2').Value2=[double]64}elseif($stage -eq 'delete'){[void]$data.Range('A2').ClearContents()}elseif($stage -eq 'error'){$data.Range('A2').Formula='=1/0'}elseif($stage -eq 'formula'){$book.Worksheets.Item('Calls').Range('D8').Formula='=SUMPRODUCT((Data!A1:A3>=4)*Data!B1:B3)'}elseif($stage -eq 'table'){$data.Range('H3').Value2=[double]4}
   $app.CalculateFullRebuild();$expected=Snapshot $book
   $expectedPath=Join-Path $directory "expected-$stage.xlsx";$book.SaveAs($expectedPath,51);$book.ExportAsFixedFormat(0,(Join-Path $directory "expected-$stage.pdf"));$book.Close($false);$book=$null
   $book=$app.Workbooks.Open($output,0,$true);$app.CalculateFullRebuild();$actual=Snapshot $book
   WriteJson "comparison-$stage.json" @{expected=$expected;actual=$actual}
   if(($expected|ConvertTo-Json -Compress) -cne ($actual|ConvertTo-Json -Compress)){throw "Native scalar/reference results differ: $stage"}
   $book.ExportAsFixedFormat(0,(Join-Path $directory "browser-$stage.pdf"))
   $reports+=,[ordered]@{stage=$stage;exportSha256=Hash $output;expectedSha256=Hash $expectedPath;cells=$actual}
   $book.Close($false);$book=$null
   $reedits=@()
   foreach($inputPath in @($expectedPath,$output)){
    $book=$app.Workbooks.Open($inputPath,0,$false);$book.Worksheets.Item('Data').Range('B2').Value2=[double]13
    $app.CalculateFullRebuild();$reeditPath=Join-Path $directory ("reedit-$stage-"+$reedits.Count+'.xlsx');$book.SaveAs($reeditPath,51);$book.Close($false);$book=$null
    $book=$app.Workbooks.Open($reeditPath,0,$true);$app.CalculateFullRebuild();$reedits+=,(Snapshot $book);$book.ExportAsFixedFormat(0,([IO.Path]::ChangeExtension($reeditPath,'pdf')));$book.Close($false);$book=$null
   }
   if(($reedits[0]|ConvertTo-Json -Compress) -cne ($reedits[1]|ConvertTo-Json -Compress)){throw "Native re-edit differs: $stage"}
   $reports[-1]['reeditCells']=$reedits[1]
  }
  $keeper.Close($false);$keeper=$null
  $artifacts=@();foreach($stage in @('edit','delete','error','formula','table')){foreach($name in @("$stage.xlsx","expected-$stage.xlsx","expected-$stage.pdf","browser-$stage.pdf","reedit-$stage-0.xlsx","reedit-$stage-1.xlsx","reedit-$stage-0.pdf","reedit-$stage-1.pdf")){$artifacts+=,@{path=$name;sha256=Hash (Join-Path $directory $name)}}}
  WriteJson 'native-report.json' ([ordered]@{passed=$true;sourceSha256=Hash $fixture;scriptSha256=Hash $PSCommandPath;browserReceiptSha256=Hash (Join-Path $directory 'browser-report.json');executableSha256=$native.sha256;stages=$reports;artifacts=$artifacts})
  Write-Output "All $($formulas.Count*5) native shaped-reference/SUMPRODUCT snapshots match five actual browser exports."
 }
 if(!$Compare){Write-Output "Recorded $($initial.Count) native array/range cases, five edit stages and unchanged-source reopen."}
}finally{
 if($book){$book.Close($false)}
 if($keeper){$keeper.Close($false)}
 if($app){if($null -ne $security){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.AskToUpdateLinks=$links;if($app.Workbooks.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
