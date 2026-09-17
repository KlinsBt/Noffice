param([string[]]$Kinds = @('word', 'excel', 'powerpoint'), [string[]]$FixtureIds, [switch]$All)
$ErrorActionPreference = 'Stop'
$corpusRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/office-corpus'))
$manifestPath = Join-Path $PSScriptRoot '../corpus/manifest.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$outputRoot = Join-Path $corpusRoot 'desktop'
New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
$selection = @('poi-docx-ComplexNumberedLists', 'poi-docx-headerFooter', 'poi-docx-VariousPictures', 'poi-xlsx-simple-monthly-budget', 'poi-xlsx-ConditionalFormattingSamples', 'poi-xlsx-WithThreeCharts', 'poi-pptx-WithMaster', 'poi-pptx-SmartArt', 'poi-pptx-bar-chart')
if ($FixtureIds) { $selection = $FixtureIds }
$applications = @{ word='Word.Application'; excel='Excel.Application'; powerpoint='PowerPoint.Application' }
$processes = @{ word='WINWORD'; excel='EXCEL'; powerpoint='POWERPNT' }
$reports = [Collections.Generic.List[object]]::new()
foreach ($kind in $Kinds) {
    $app = $null
    $existed = @(Get-Process -Name $processes[$kind] -ErrorAction SilentlyContinue).Count -gt 0
    $security = $null
    $alerts = $null
    $wordLinks = $null
    try {
        $app = New-Object -ComObject $applications[$kind]
        $version = $app.Version
        $security = $app.AutomationSecurity
        $app.AutomationSecurity = 3 # Force-disable document macros during tests.
        $alerts = $app.DisplayAlerts
        if ($kind -eq 'powerpoint') { $app.DisplayAlerts = 1 } else { $app.DisplayAlerts = 0 }
        if ($kind -eq 'word') { $wordLinks = $app.Options.UpdateLinksAtOpen; $app.Options.UpdateLinksAtOpen = $false }
        $fixtures = @($manifest.files | Where-Object { $_.kind -eq $kind -and ($All -or $selection -contains $_.id) -and $_.id -ne 'poi-docx-deep-table-cell' })
        foreach ($fixture in $fixtures) {
            foreach ($stage in @('source', 'unchanged', 'edited')) {
                $extension = [IO.Path]::GetExtension($fixture.filename)
                $inputPath = if ($stage -eq 'source') { Join-Path $corpusRoot "input/$kind/$($fixture.filename)" } else { Join-Path $corpusRoot "results/$($fixture.id)/$stage$extension" }
                if (!(Test-Path -LiteralPath $inputPath)) { continue }
                $record = [ordered]@{ id=$fixture.id; kind=$kind; stage=$stage; version=$version; status='failed' }
                $document = $null
                try {
                    $inputPath = [IO.Path]::GetFullPath($inputPath)
                    $pdf = Join-Path $outputRoot "$($fixture.id)-$stage.pdf"
                    if ($kind -eq 'word') {
                        $document = $app.Documents.Open($inputPath, $false, $true, $false)
                        $record.pages = $document.ComputeStatistics(2)
                        $record.tables = $document.Tables.Count
                        $record.inlineShapes = $document.InlineShapes.Count
                        $record.floatingShapes = $document.Shapes.Count
                        $document.ExportAsFixedFormat($pdf, 17)
                    } elseif ($kind -eq 'excel') {
                        $document = $app.Workbooks.Open($inputPath, 0, $true)
                        $record.sheets = $document.Sheets.Count
                        $record.names = $document.Names.Count
                        $document.ExportAsFixedFormat(0, $pdf)
                    } else {
                        $document = $app.Presentations.Open($inputPath, -1, 0, 0)
                        $record.slides = $document.Slides.Count
                        $record.widthPoints = $document.PageSetup.SlideWidth
                        $record.heightPoints = $document.PageSetup.SlideHeight
                        $document.SaveAs($pdf, 32)
                    }
                    $record.status = 'opened-and-rendered'
                    $record.pdf = $pdf
                } catch { $record.error = $_.Exception.Message }
                finally {
                    if ($null -ne $document) {
                        if ($kind -eq 'powerpoint') { $document.Close() } else { $document.Close($false) }
                        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
                    }
                }
                $reports.Add([pscustomobject]$record)
                $reports | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $outputRoot 'report.json') -Encoding UTF8
                Write-Output "$kind $($fixture.filename) $stage : $($record.status)"
            }
        }
    } catch {
        $reports.Add([pscustomobject]@{kind=$kind;status='automation-unavailable';error=$_.Exception.Message})
        $reports | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $outputRoot 'report.json') -Encoding UTF8
        Write-Output "$kind automation unavailable: $($_.Exception.Message)"
    } finally {
        if ($null -ne $app) {
            if ($null -ne $wordLinks) { $app.Options.UpdateLinksAtOpen = $wordLinks }
            if ($null -ne $security) { $app.AutomationSecurity = $security }
            if ($null -ne $alerts) { $app.DisplayAlerts = $alerts }
            # Never quit an Office application that was already running before these tests.
            if (!$existed) {
                $openDocuments = if ($kind -eq 'word') { $app.Documents.Count } elseif ($kind -eq 'excel') { $app.Workbooks.Count } else { $app.Presentations.Count }
                if ($openDocuments -eq 0) { $app.Quit() }
            }
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
        }
    }
}
foreach ($kind in $Kinds) {
    @($reports | Where-Object kind -eq $kind) | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $outputRoot "report-$kind.json") -Encoding UTF8
}
Write-Output "Desktop test evidence: $outputRoot"
if (@($reports | Where-Object status -ne 'opened-and-rendered').Count -gt 0) { exit 1 }
