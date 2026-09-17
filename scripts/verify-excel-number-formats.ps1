param([string]$InputPath = 'tests/fixtures/number-formats-input.json', [string]$OutputPath = '.local/number-formats-oracle.json')
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
$spec = Get-Content -LiteralPath $InputPath -Raw -Encoding UTF8 | ConvertFrom-Json
$excel = $null
$book = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $alerts = $excel.DisplayAlerts
  $security = $excel.AutomationSecurity
  $system = $excel.UseSystemSeparators
  $decimal = $excel.DecimalSeparator
  $thousands = $excel.ThousandsSeparator
  $excel.DisplayAlerts = $false
  $excel.AutomationSecurity = 3
  $excel.UseSystemSeparators = $false
  $excel.DecimalSeparator = '.'
  $excel.ThousandsSeparator = ','
  [uint32]$testProcessId = 0
  [void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$testProcessId)
  $book = $excel.Workbooks.Add()
  $sheet = $null
  for ($attempt = 0; $attempt -lt 20 -and $null -eq $sheet; $attempt++) {
    [ExcelTestWindow]::DismissReminder($testProcessId)
    Start-Sleep -Milliseconds 250
    try { $sheet = $book.Worksheets.Item(1) } catch { if ($attempt -eq 19) { throw } }
  }
  $sheet.Name = 'Format oracle'
  $cases = @()
  $row = 0
  foreach ($case in $spec.cases) {
    $row++
    $nativeCode = $case.code
    if ($excel.LanguageSettings.LanguageID(2) -eq 1031 -and $case.nativeCode) { $nativeCode = $case.nativeCode }
    if ($excel.LanguageSettings.LanguageID(2) -eq 1031) { $nativeCode = $nativeCode.Replace('[Red]', '[Rot]') }
    $nativeValue = [double]$case.value + $(if ($case.date1904) { 1462 } else { 0 })
    $sheet.Range("A$row").Value2 = $nativeValue
    $sheet.Range("B$row").NumberFormat = '@'
    $sheet.Range("B$row").Value2 = [string]$nativeCode
    $sheet.Range("C$row").Formula = "=TEXT(A$row,B$row)"
  }
  $excel.CalculateFullRebuild()
  $row = 0
  foreach ($case in $spec.cases) {
    $row++
    $range = $sheet.Range("C$row")
    $value = $range.Value2
    $kind = 'value'
    if ($excel.WorksheetFunction.IsError($range)) {
      $kind = 'error'
      $code = [int]([long]$value -band 65535)
      $value = @{2000='#NULL!'; 2007='#DIV/0!'; 2015='#VALUE!'; 2023='#REF!'; 2029='#NAME?'; 2036='#NUM!'; 2042='#N/A'}[$code]
      if ($null -eq $value) { throw "Unrecognized Excel error code $code" }
    }
    $cases += [ordered]@{ code=$case.code; input=$case.value; date1904=[bool]$case.date1904; nativeCode=$sheet.Range("B$row").Value2; nativeValue=$sheet.Range("A$row").Value2; value=$value; kind=$kind }
  }
  $outputAbsolute = [IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputPath))
  [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($outputAbsolute)) | Out-Null
  [ordered]@{ application='Microsoft Excel'; version=$excel.Version; build=$excel.Build; locale=$excel.LanguageSettings.LanguageID(2); inputSha256=(Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLowerInvariant(); cases=$cases } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $outputAbsolute -Encoding UTF8
  $book.SaveAs([IO.Path]::ChangeExtension($outputAbsolute, '.xlsx'), 51)
  Write-Output "Recorded $($cases.Count) native format results in $OutputPath"
} finally {
  if ($null -ne $book) { try { $book.Close($false) } catch { Write-Warning $_ } }
  if ($null -ne $excel) {
    try {
      $excel.DisplayAlerts = $alerts; $excel.AutomationSecurity = $security
      $excel.DecimalSeparator = $decimal; $excel.ThousandsSeparator = $thousands; $excel.UseSystemSeparators = $system
      if ($excel.Workbooks.Count -eq 0) { $excel.Quit() }
    } catch { Write-Warning $_ }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
  }
}
