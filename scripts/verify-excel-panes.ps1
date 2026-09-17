$ErrorActionPreference = 'Stop'
$app = $null; $book = $null
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
function Invoke-PaneCall([scriptblock]$Action) {
  for ($retry = 0; $retry -lt 20; $retry++) {
    try { return (& $Action) } catch {
      if ($_.Exception.HResult -notin @(-2147418111, -2147417846, -2146777998) -or $retry -eq 19) { throw }
      [ExcelTestWindow]::DismissReminder($testProcessId)
      Start-Sleep -Milliseconds 250
    }
  }
}
try {
  $app = New-Object -ComObject Excel.Application
  $app.Visible = $false; $app.DisplayAlerts = $false; $app.AutomationSecurity = 3; $app.EnableEvents = $false
  [uint32]$testProcessId = 0
  [void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd, [ref]$testProcessId)
  for ($startup = 0; $startup -lt 8; $startup++) {
    [ExcelTestWindow]::DismissReminder($testProcessId)
    Start-Sleep -Milliseconds 250
  }
  $results = @()
  foreach ($name in @('frozen-panes','unfrozen-panes')) {
    $path = [IO.Path]::GetFullPath((Join-Path (Get-Location) ".local/xlsx-validation/$name.xlsx"))
    $book = Invoke-PaneCall { $app.Workbooks.Open($path,0,$true) }
    if ($null -eq $book) { throw 'Excel returned no workbook when opening the generated test export.' }
    $window = Invoke-PaneCall { $book.Windows.Item(1) }
    $frozen = Invoke-PaneCall { [bool]$window.FreezePanes }
    $rows = Invoke-PaneCall { [int]$window.SplitRow }; $columns = Invoke-PaneCall { [int]$window.SplitColumn }
    $expectedRows=0; $expectedColumns=0; $expectedFrozen=$false
    if ($name -eq 'frozen-panes') { $expectedRows=3; $expectedColumns=2; $expectedFrozen=$true }
    if ($frozen -ne $expectedFrozen -or $rows -ne $expectedRows -or $columns -ne $expectedColumns) { throw "Native pane settings do not match $name." }
    $value = Invoke-PaneCall { $book.Worksheets.Item(1).Range('D4').Value2 }
    if ($value -ne 3) { throw 'Pane changes altered the calculated test value.' }
    $results += [ordered]@{name=$name;inputSha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();frozen=$frozen;rows=$rows;columns=$columns;value=$value}
    Invoke-PaneCall { $book.Close($false) }; $book=$null
  }
  [ordered]@{application='Microsoft Excel';version=[string]$app.Version;build=[string]$app.Build;cases=$results} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath .local/xlsx-validation/panes-report.json -Encoding UTF8
  Write-Output 'Native Excel verified frozen and unfrozen browser exports, splits and calculated value.'
} finally {
  if ($book) { try { Invoke-PaneCall { $book.Close($false) } } catch { Write-Warning $_ } }
  if ($app) { try { Invoke-PaneCall { $app.Quit() } } catch { Write-Warning $_ } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } }
}
