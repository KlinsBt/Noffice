param([switch]$Compare)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root '.local/excel-arrays'
[void][IO.Directory]::CreateDirectory($directory)
. (Join-Path $root 'scripts/excel-test-window.ps1')
function Hash($path){(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant()}
function WriteJson($name,$value){[IO.File]::WriteAllText((Join-Path $directory $name),($value|ConvertTo-Json -Depth 20),(New-Object Text.UTF8Encoding($false)))}
$fixture=Join-Path $root 'tests/fixtures/excel-array-contexts.xlsx'
$oracle=Get-Content (Join-Path $root 'tests/fixtures/native-excel-array-contexts.json') -Raw|ConvertFrom-Json
if((Hash $fixture) -ne $oracle.sourceSha256){throw 'Changed array fixture'}
function SetArray($cell,[string]$formula){[void]$cell.GetType().InvokeMember('FormulaArray',[Reflection.BindingFlags]::SetProperty,$null,$cell,@($formula),[Globalization.CultureInfo]::GetCultureInfo('en-US'))}
function Snapshot($book){
 $sheet=$book.Worksheets.Item('Calls');$rows=@()
 foreach($expected in $oracle.stages[0].cells){
  $cell=$sheet.Range($expected.ref);$value=$cell.Value2;$error=$app.WorksheetFunction.IsError($cell)
  if($error){$code=[int]([long]$value -band 65535);$value=@{2000='#NULL!';2007='#DIV/0!';2015='#VALUE!';2023='#REF!';2029='#NAME?';2036='#NUM!';2042='#N/A'}[$code];if($null -eq $value){throw "Unknown error $code"}}
  $rows+=,[ordered]@{ref=$expected.ref;formula=[string]$cell.Formula;value=$value;kind=$(if($error){'error'}else{'value'});hasArray=[bool]$cell.HasArray;arrayAddress=$(if($cell.HasArray){[string]$cell.CurrentArray.Address()}else{$null})}
 };return ,$rows
}
function Apply($book,$stage){
 $data=$book.Worksheets.Item('Data');$calls=$book.Worksheets.Item('Calls')
 if($stage -eq 'edit'){$data.Range('A2').Value2=[double]0}
 if($stage -eq 'error'){$data.Range('A2').Formula='=1/0'}
 if($stage -eq 'entry'){SetArray $calls.Range('B1') ([string]$calls.Range('B1').Formula)}
 if($stage -eq 'ordinary'){$formula=[string]$calls.Range('C1').Formula;[void]$calls.Range('C1').ClearContents();$calls.Range('C1').Formula=$formula}
 if($stage -eq 'formula'){SetArray $calls.Range('C1') '=SUMPRODUCT(IF(Data!A1:A3>6,Data!B1:B3,0))'}
 $app.CalculateFullRebuild()
}
$app=$null;$book=$null;$keeper=$null;$idle=$false
try{
 $app=New-Object -ComObject Excel.Application
 if($app.Workbooks.Count -ne 0){throw 'Requires an idle owned Excel instance.'};$idle=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.AskToUpdateLinks
 WriteJson 'preferences.json' @{security=$security;alerts=$alerts;links=$links}
 $app.AutomationSecurity=3;$app.DisplayAlerts=$false;$app.AskToUpdateLinks=$false
 $native=(Get-Content (Join-Path $root 'docs/parity/baseline.json') -Raw|ConvertFrom-Json).applications|Where-Object executable -eq 'EXCEL.EXE'
 if((Hash (Join-Path ([string]$app.Path) 'EXCEL.EXE')) -ne $native.sha256){throw 'Excel baseline changed'}
 [uint32]$owned=0;[void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$owned);[ExcelTestWindow]::DismissReminder($owned)
 $keeper=$app.Workbooks.Add();$stages=@();$artifacts=@()
 foreach($stage in @('edit','error','entry','ordinary','formula')){
  $book=$app.Workbooks.Open($fixture,0,$false);Apply $book $stage;$before=Snapshot $book
  [string]$expectedPath=Join-Path $directory "expected-$stage.xlsx";$book.SaveAs($expectedPath,51);$book.Close($false);$book=$null
  $book=$app.Workbooks.Open($expectedPath,0,$true);$app.CalculateFullRebuild();$expected=Snapshot $book
  if(($before|ConvertTo-Json -Compress) -cne ($expected|ConvertTo-Json -Compress)){throw "Native reopen changed $stage"}
  $book.ExportAsFixedFormat(0,(Join-Path $directory "expected-$stage.pdf"));$book.Close($false);$book=$null
  $row=[ordered]@{stage=$stage;cells=$expected;expectedSha256=Hash $expectedPath}
  if($Compare){
   [string]$output=Join-Path $directory "$stage.xlsx";$book=$app.Workbooks.Open($output,0,$true);$app.CalculateFullRebuild();$actual=Snapshot $book
   WriteJson "comparison-$stage.json" @{expected=$expected;actual=$actual}
   if(($actual|ConvertTo-Json -Compress) -cne ($expected|ConvertTo-Json -Compress)){throw "Actual export differs at $stage"}
   $book.ExportAsFixedFormat(0,(Join-Path $directory "browser-$stage.pdf"));$book.Close($false);$book=$null
   $row['exportSha256']=Hash $output;$reedits=@()
   foreach($inputPath in @($expectedPath,$output)){
    $book=$app.Workbooks.Open($inputPath,0,$false);$book.Worksheets.Item('Data').Range('B2').Value2=[double]13;$app.CalculateFullRebuild()
    [string]$reedit=Join-Path $directory ("reedit-$stage-"+$reedits.Count+'.xlsx');$book.SaveAs($reedit,51);$book.Close($false);$book=$null
    $book=$app.Workbooks.Open($reedit,0,$true);$app.CalculateFullRebuild();$reedits+=,(Snapshot $book);$book.ExportAsFixedFormat(0,([IO.Path]::ChangeExtension($reedit,'pdf')));$book.Close($false);$book=$null
   }
   if(($reedits[0]|ConvertTo-Json -Compress) -cne ($reedits[1]|ConvertTo-Json -Compress)){throw "Native re-edit differs at $stage"}
   $row['reeditCells']=$reedits[1]
  }
  $stages+=,$row
  $names=@("expected-$stage.xlsx","expected-$stage.pdf")
  if($Compare){$names+=@("$stage.xlsx","browser-$stage.pdf","reedit-$stage-0.xlsx","reedit-$stage-1.xlsx","reedit-$stage-0.pdf","reedit-$stage-1.pdf")}
  foreach($name in $names){$artifacts+=,@{path=$name;sha256=Hash (Join-Path $directory $name)}}
 }
 $report=@{passed=$true;sourceSha256=Hash $fixture;executableSha256=$native.sha256;scriptSha256=Hash $PSCommandPath;stages=$stages;artifacts=$artifacts}
 if($Compare){$report['browserReceiptSha256']=Hash (Join-Path $directory 'browser-report.json')}
 WriteJson $(if($Compare){'native-report.json'}else{'reference.json'}) $report
 Write-Output "Verified 160 native value/type/formula/array-identity snapshots; actual exports compared: $Compare."
}finally{
 if($book){$book.Close($false)}
 if($keeper){$keeper.Close($false)}
 if($app){if($idle){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.AskToUpdateLinks=$links;if($app.Workbooks.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
