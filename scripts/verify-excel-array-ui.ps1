$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root '.local/excel-array-ui'
[void][IO.Directory]::CreateDirectory($directory)
. (Join-Path $root 'scripts/excel-test-window.ps1')
Add-Type -AssemblyName System.Windows.Forms
Add-Type @'
using System;using System.Runtime.InteropServices;
public class ArrayInput {
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint pid);
}
'@
function Keys([string]$keys){
 Write-Output "Keys: $keys"
 [ExcelTestWindow]::DismissReminder($owned)
 [uint32]$front=0;[void][ArrayInput]::GetWindowThreadProcessId([ArrayInput]::GetForegroundWindow(),[ref]$front)
 if($front -ne $owned){throw 'Refusing keys outside the owned Excel process.'}
 [Windows.Forms.SendKeys]::SendWait($keys);Start-Sleep -Milliseconds 350
}
function Retry([scriptblock]$action){for($i=0;$i -lt 30;$i++){try{return (& $action)}catch{if($_.Exception.ToString() -notmatch 'RPC_E_CALL_REJECTED|0x80010001|0x8001010A' -or $i -eq 29){throw};[ExcelTestWindow]::DismissReminder($owned);Start-Sleep -Milliseconds 200}}}
function Snapshot($name){
 Write-Host "Snapshot: $name"
 $row=Retry { $cell=$sheet.Range('B1');@{stage=$name;formula=[string]$cell.Formula;value=$cell.Value2;hasArray=[bool]$cell.HasArray;selection=[string]$app.Selection.Address()} }
 $script:rows+=,$row;return $row
}
$app=$null;$book=$null;$idle=$false;$rows=@()
try{
 $app=New-Object -ComObject Excel.Application
 if($app.Workbooks.Count -ne 0){throw 'Requires an idle owned Excel instance.'};$idle=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.AskToUpdateLinks
 $app.AutomationSecurity=3;$app.DisplayAlerts=$false;$app.AskToUpdateLinks=$false
 $baseline=(Get-Content (Join-Path $root 'docs/parity/baseline.json') -Raw|ConvertFrom-Json).applications|Where-Object executable -eq 'EXCEL.EXE'
 if((Get-FileHash (Join-Path ([string]$app.Path) 'EXCEL.EXE')).Hash.ToLowerInvariant() -ne $baseline.sha256){throw 'Excel baseline changed'}
 [uint32]$owned=0;[void][ArrayInput]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$owned);[ExcelTestWindow]::DismissReminder($owned)
 @{pid=$owned;security=$security;alerts=$alerts;links=$links}|ConvertTo-Json|Set-Content (Join-Path $directory 'owned.json')
 $book=$app.Workbooks.Open([string](Join-Path $root 'tests/fixtures/excel-array-contexts.xlsx'),0,$false)
 $sheet=$book.Worksheets.Item('Calls');$sheet.Activate();$sheet.Range('B1').Select()
 $app.Visible=$true;Start-Sleep -Milliseconds 500;[ExcelTestWindow]::DismissReminder($owned);[void][ArrayInput]::SetForegroundWindow([IntPtr]$app.Hwnd);Start-Sleep -Milliseconds 500
 $before=Snapshot 'ordinary';if($before.value -ne 0 -or $before.hasArray){throw 'Unexpected source'}
 Keys '{F2}';Keys '^+{ENTER}'
 $entered=Snapshot 'CSE';if($entered.value -ne 13 -or !$entered.hasArray -or $entered.selection -ne '$B$1'){throw 'CSE entry differs'}
 Keys '^z';$undo=Snapshot 'undo';if($undo.value -ne 0 -or $undo.hasArray -or $undo.selection -ne '$B$1'){throw 'Undo differs'}
 Keys '^y';$redo=Snapshot 'redo';if($redo.value -ne 13 -or !$redo.hasArray -or $redo.selection -ne '$B$1'){throw 'Redo differs'}
 Keys '{F2}';Keys '{ENTER}'
 $ordinary=Snapshot 'ordinary-reentry'
 if($ordinary.value -ne 0 -or $ordinary.hasArray -or $ordinary.selection -ne '$B$2'){throw 'Ordinary reentry differs'}
 [string]$saved=Join-Path $directory 'ui.xlsx';$book.SaveAs($saved,51);$book.Close($false);$book=$null
 $savedHash=(Get-FileHash $saved).Hash.ToLowerInvariant()
 $book=$app.Workbooks.Open($saved,0,$true);$sheet=$book.Worksheets.Item('Calls');$reopen=Snapshot 'reopen'
 if($reopen.value -ne 0 -or $reopen.hasArray -or $reopen.selection -ne '$B$2'){throw 'Saved ordinary reentry differs'}
 @{passed=$true;scope='Owned native F2/Ctrl+Shift+Enter, Undo/Redo and ordinary Enter reentry';rows=$rows;executableSha256=$baseline.sha256;scriptSha256=(Get-FileHash $PSCommandPath).Hash.ToLowerInvariant();savedSha256=$savedHash;sourceSha256=(Get-FileHash (Join-Path $root 'tests/fixtures/excel-array-contexts.xlsx')).Hash.ToLowerInvariant()}|ConvertTo-Json -Depth 10|Set-Content (Join-Path $directory 'report.json') -Encoding UTF8
 Write-Output ($rows|ConvertTo-Json -Depth 6)
}catch{Write-Output $_.Exception.ToString();throw}finally{
 if($book){Retry {$book.Close($false)}}
 if($app){if($idle){Retry {$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.AskToUpdateLinks=$links;if($app.Workbooks.Count -eq 0){$app.Quit()}}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
