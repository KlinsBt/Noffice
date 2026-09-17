param([string]$InputPath = 'tests/fixtures/excel-formulas-input.json', [string]$OutputPath = '.local/excel-formulas-oracle.json', [switch]$UseRunningTestInstance)
$ErrorActionPreference = 'Stop'
$spec = Get-Content -LiteralPath $InputPath -Raw | ConvertFrom-Json
$excel = $null
$book = $null
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
try {
  $excel = if ($UseRunningTestInstance) { [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } else { New-Object -ComObject Excel.Application }
  $alerts = $excel.DisplayAlerts
  $security = $excel.AutomationSecurity
  $excel.DisplayAlerts = $false
  $excel.AutomationSecurity = 3
  [uint32]$testProcessId = 0
  [void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$testProcessId)
  $book = $excel.Workbooks.Add()
  $sheet = $null
  for ($attempt = 0; $attempt -lt 20 -and $null -eq $sheet; $attempt++) {
    # Close only the ordinary reminder in this test instance; never alter activation.
    [ExcelTestWindow]::DismissReminder($testProcessId)
    Start-Sleep -Milliseconds 250
    try { $sheet = $book.Worksheets.Item(1) } catch { if ($attempt -eq 19) { throw } }
  }
  $sheet.Name = $spec.sheet
  foreach ($property in $spec.cells.PSObject.Properties) {
    $cell = $sheet.Range($property.Name)
    $value = $property.Value
    if ($value -is [string] -and $value.StartsWith('=')) { $cell.Formula = $value }
    elseif ($value -is [string]) { $cell.NumberFormat = '@'; $cell.Value2 = $value }
    elseif ($value -is [bool]) { $cell.Formula = $(if ($value) { '=TRUE()' } else { '=FALSE()' }) }
    else { $cell.Value2 = [double]$value }
  }
  if ($spec.table) {
    $table = $sheet.ListObjects.Add(1, $sheet.Range($spec.table.ref), $null, 1)
    $table.Name = $spec.table.name
  }
  foreach ($property in $spec.names.PSObject.Properties) {
    if ($property.Value -match "^'Data Set'!(.+)$") { $sheet.Range($Matches[1]).Name = [string]$property.Name }
    else { $entry = $book.Names.Add([string]$property.Name, '=1'); $entry.RefersTo = '=' + [string]$property.Value }
  }
  foreach ($case in $spec.cases) {
    # Reference/error names inside strings follow Excel's installed language.
    $formula = $case.formula
    if ($excel.LanguageSettings.LanguageID(2) -eq 1031) {
      $formula = $formula.Replace('"R2C2"', '"Z2S2"').Replace('"R[-15]C[-8]"', '"Z(-15)S(-8)"').Replace('"#N/A"', '"#NV"').Replace('"<>#N/A"', '"<>#NV"')
    }
    try { $sheet.Range($case.ref).Formula = $formula }
    catch { throw "Excel rejected $($case.ref): $formula ($($_.Exception.Message))" }
  }
  foreach ($filter in $spec.filters) {
    [void]$sheet.Range($filter.ref).AutoFilter([int]$filter.field, [string]$filter.criteria)
  }
  foreach ($row in $spec.hiddenRows) { $sheet.Rows.Item([int]$row + 1).Hidden = $true }
  foreach ($column in $spec.hiddenColumns) { $sheet.Columns.Item([int]$column + 1).Hidden = $true }
  foreach ($property in $spec.edits.PSObject.Properties) { $sheet.Range($property.Name).Value2 = [string]$property.Value }
  $excel.CalculateFullRebuild()
  $cases = @()
  foreach ($case in $spec.cases) {
    $range = $sheet.Range($case.ref)
    $value = $range.Value2
    $isError = $excel.WorksheetFunction.IsError($range)
    if ($isError) {
      $code = [int]([long]$value -band 65535)
      $value = @{2000='#NULL!'; 2007='#DIV/0!'; 2015='#VALUE!'; 2023='#REF!'; 2029='#NAME?'; 2036='#NUM!'; 2042='#N/A'; 2043='#GETTING_DATA'}[$code]
      if ($null -eq $value) { throw "Unrecognized Excel error code $code" }
    }
    $cases += [ordered]@{ ref = $case.ref; formula = $case.formula; nativeFormula = $range.Formula; value = $value; kind = $(if ($isError) { 'error' } else { 'value' }) }
  }
  $outputAbsolute = [IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputPath))
  [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($outputAbsolute)) | Out-Null
  $hiddenRows = @()
  for ($row = 1; $row -le $sheet.UsedRange.Rows.Count; $row++) { if ($sheet.Rows.Item($row).Hidden) { $hiddenRows += $row - 1 } }
  [ordered]@{ application = 'Microsoft Excel'; version = $excel.Version; build = $excel.Build; locale = $excel.LanguageSettings.LanguageID(2); inputSha256 = (Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLowerInvariant(); hiddenRows = $hiddenRows; cases = $cases } |
    ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $outputAbsolute -Encoding UTF8
  $xlsxPath = [IO.Path]::ChangeExtension($outputAbsolute, '.xlsx')
  $book.SaveAs($xlsxPath, 51)
  Write-Output "Recorded $($cases.Count) native formula results in $OutputPath"
} catch {
  Write-Output $_.ScriptStackTrace
  throw
} finally {
  if ($null -ne $book) { try { $book.Close($false) } catch { Write-Warning $_ } }
  if ($null -ne $excel) {
    try {
      $excel.DisplayAlerts = $alerts
      $excel.AutomationSecurity = $security
      if (!$UseRunningTestInstance -and $excel.Workbooks.Count -eq 0) { $excel.Quit() }
    } catch { Write-Warning $_ }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
  }
}

