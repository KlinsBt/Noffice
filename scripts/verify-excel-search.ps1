$ErrorActionPreference = 'Stop'
$inputPath = Join-Path (Get-Location) 'tests/fixtures/search-oracle-cases.json'
$outputPath = Join-Path (Get-Location) 'tests/fixtures/native-excel-search.json'
$cases = Get-Content -LiteralPath $inputPath -Raw -Encoding UTF8 | ConvertFrom-Json
$app = $null; $book = $null
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
function Invoke-ExcelSearchCall([scriptblock]$Action) {
  for ($retry = 0; $retry -lt 20; $retry++) {
    try { return (& $Action) } catch {
      if ($_.Exception.HResult -notin @(-2147418111, -2147417846) -or $retry -eq 19) { throw }
      [ExcelTestWindow]::DismissReminder($testProcessId)
      Start-Sleep -Milliseconds 250
    }
  }
}
try {
  $app = New-Object -ComObject Excel.Application
  $app.Visible = $false
  $app.DisplayAlerts = $false
  $app.AutomationSecurity = 3
  $app.EnableEvents = $false
  [uint32]$testProcessId = 0
  [void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd, [ref]$testProcessId)
  $book = $app.Workbooks.Add()
  $sheet = $null
  for ($attempt = 0; $attempt -lt 20 -and $null -eq $sheet; $attempt++) {
    [ExcelTestWindow]::DismissReminder($testProcessId)
    Start-Sleep -Milliseconds 250
    try { $sheet = $book.Worksheets.Item(1) } catch { if ($attempt -eq 19) { throw } }
  }
  $cell = $sheet.Range('A1')
  $results = @()
  foreach ($case in $cases) {
    Invoke-ExcelSearchCall { $cell.NumberFormat = '@'; $cell.Value2 = [string]$case.text }
    $lookAt = 2
    if ($case.wholeCell) { $lookAt = 1 }
    $found = Invoke-ExcelSearchCall { $null -ne $cell.Find([string]$case.query, $cell, -4123, $lookAt, 1, 1, [bool]$case.matchCase, $false, $false) }
    Invoke-ExcelSearchCall { [void]$cell.Replace([string]$case.query, [string]$case.replacement, $lookAt, 1, [bool]$case.matchCase, $false, $false, $false) }
    $results += [ordered]@{ text=$case.text; query=$case.query; replacement=$case.replacement; matchCase=[bool]$case.matchCase; wholeCell=[bool]$case.wholeCell; found=[bool]$found; result=[string]$cell.Value2 }
  }
  $report = [ordered]@{ application='Microsoft Excel'; version=[string]$app.Version; build=[string]$app.Build; inputSha256=(Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant(); cases=$results }
  [IO.File]::WriteAllText($outputPath, ($report | ConvertTo-Json -Depth 8) + "`n", (New-Object Text.UTF8Encoding($false)))
  Write-Output ("Recorded {0} native Excel search replacements ({1}, build {2})." -f $results.Count, $app.Version, $app.Build)
} finally {
  if ($book) { try { Invoke-ExcelSearchCall { $book.Close($false) } } catch { Write-Warning $_ } }
  if ($app) { try { Invoke-ExcelSearchCall { $app.Quit() } } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } }
}
