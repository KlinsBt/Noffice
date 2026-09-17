param([string]$InputPath = '.local/filter-validation/browser-export.xlsx')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
$excel=$null; $book=$null
try {
  $excel=New-Object -ComObject Excel.Application
  $alerts=$excel.DisplayAlerts; $security=$excel.AutomationSecurity
  $excel.DisplayAlerts=$false; $excel.AutomationSecurity=3
  [uint32]$testProcessId=0
  [void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$testProcessId)
  $inputAbsolute=[IO.Path]::GetFullPath((Join-Path (Get-Location) $InputPath))
  $book=$excel.Workbooks.Open($inputAbsolute,0,$true)
  $sheet=$null
  for($attempt=0;$attempt -lt 20 -and $null -eq $sheet;$attempt++) {
    [ExcelTestWindow]::DismissReminder($testProcessId); Start-Sleep -Milliseconds 250
    try {$sheet=$book.Worksheets.Item(1)} catch {if($attempt -eq 19){throw}}
  }
  $excel.CalculateFullRebuild()
  $filtered=$sheet.Range('D1').Value2
  if($filtered -ne 2 -or !$sheet.Rows.Item(3).Hidden -or !$sheet.Rows.Item(4).Hidden) {throw 'Native Excel did not retain the browser filter result.'}
  $sheet.AutoFilter.ApplyFilter()
  $excel.CalculateFullRebuild()
  $reapplied=$sheet.Range('D1').Value2
  if($reapplied -ne 2) {throw 'Reapplying the exported criteria in Excel changed the expected total.'}
  $sheet.ShowAllData(); $excel.CalculateFullRebuild()
  $cleared=$sheet.Range('D1').Value2
  if($cleared -ne 12) {throw 'Clearing the exported filter in Excel did not restore the total.'}
  $report=[ordered]@{application='Microsoft Excel';version=$excel.Version;build=$excel.Build;inputSha256=(Get-FileHash -LiteralPath $inputAbsolute -Algorithm SHA256).Hash.ToLowerInvariant();filtered=$filtered;reapplied=$reapplied;cleared=$cleared}
  $report | ConvertTo-Json | Set-Content -LiteralPath (Join-Path ([IO.Path]::GetDirectoryName($inputAbsolute)) 'desktop-report.json') -Encoding UTF8
  Write-Output 'Native Excel browser-export check passed: filtered 2, reapplied 2, cleared 12.'
} finally {
  if($null -ne $book){try{$book.Close($false)}catch{Write-Warning $_}}
  if($null -ne $excel){try{$excel.DisplayAlerts=$alerts;$excel.AutomationSecurity=$security;if($excel.Workbooks.Count -eq 0){$excel.Quit()}}catch{Write-Warning $_};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
}
