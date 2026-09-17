$ErrorActionPreference = 'Stop'
$existed=@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue).Count -gt 0
$app=$null; $document=$null
try {
  $app=New-Object -ComObject Word.Application
  $alerts=$app.DisplayAlerts; $security=$app.AutomationSecurity
  $app.Visible=$false; $app.DisplayAlerts=0; $app.AutomationSecurity=3
  $path=[IO.Path]::GetFullPath((Join-Path (Get-Location) '.local/docx-validation/links.docx'))
  $document=$app.Documents.Open($path,$false,$true,$false)
  if ($document.Hyperlinks.Count -ne 3) { throw 'Unexpected native Word hyperlink count.' }
  $expected=@('https://example.com/revised?a=1&b=2','https://example.com/original','mailto:team@example.com')
  $links=@()
  for ($i=1; $i -le 3; $i++) {
    $link=$document.Hyperlinks.Item($i)
    if ($link.Address -ne $expected[$i-1]) { throw "Unexpected native Word hyperlink target $i." }
    $links += [ordered]@{address=[string]$link.Address;text=[string]$link.TextToDisplay}
  }
  $first=$document.Hyperlinks.Item(1)
  # A hyperlink range includes hidden field-code positions; inspect the displayed result.
  $result=$first.Range.Fields.Item(1).Result
  if ($result.Text -ne 'Alpha beta') { throw 'Native Word lost the hyperlink display text.' }
  for ($n=0; $n -lt 10; $n++) {
    $character=$document.Range($result.Start+$n,$result.Start+$n+1)
    $expectedBold=if($n -lt 5){-1}else{0}; $expectedItalic=if($n -lt 5){0}else{-1}
    if ($character.Font.Bold -ne $expectedBold -or $character.Font.Italic -ne $expectedItalic) { throw "Native Word formatting mismatch at display character $n, bold=$($character.Font.Bold), italic=$($character.Font.Italic)." }
  }
  if ($document.Hyperlinks.Item(3).TextToDisplay -ne 'Visit now') { throw 'Native Word lost the inserted display text.' }
  $pdf=[IO.Path]::ChangeExtension($path,'.pdf')
  $document.ExportAsFixedFormat($pdf,17)
  [ordered]@{application='Microsoft Word';version=[string]$app.Version;inputSha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();links=$links;mixedFormattingVerified=$true;pdf=$pdf} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath .local/docx-validation/links-report.json -Encoding UTF8
  Write-Output 'Native Word verified hyperlink destinations and rendered the browser export.'
} finally {
  $saveOption=0
  if ($document) { try {$document.Close([ref]$saveOption)}catch{Write-Warning $_} }
  if ($app) { try {$app.DisplayAlerts=$alerts;$app.AutomationSecurity=$security;if(!$existed -and $app.Documents.Count -eq 0){$app.Quit([ref]$saveOption)}}catch{Write-Warning $_}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
}
