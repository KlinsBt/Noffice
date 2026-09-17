$ErrorActionPreference='Stop'
$app=$null; $presentation=$null; $reports=@()
try {
  $app=New-Object -ComObject PowerPoint.Application
  $alerts=$app.DisplayAlerts; $security=$app.AutomationSecurity
  $app.DisplayAlerts=1; $app.AutomationSecurity=3
  foreach($name in @('edited','additions')) {
    $path=[IO.Path]::GetFullPath((Join-Path (Get-Location) ".local/pptx-validation/$name.pptx"))
    Write-Output "Checking native export: $name.pptx"
    $presentation=$app.Presentations.Open($path,-1,0,0)
    if($presentation.Slides.Count -ne 2) {throw 'The retained slide count changed.'}
    if($presentation.PageSetup.SlideWidth -ne 720 -or $presentation.PageSetup.SlideHeight -ne 540) {throw 'The original 4:3 slide dimensions changed.'}
    $chart=$false
    foreach($shape in $presentation.Slides.Item(2).Shapes) {if($shape.HasChart -eq -1){$chart=$true}}
    if(!$chart){throw 'Native PowerPoint could not find the retained chart.'}
    $text=@();$pictures=0
    foreach($shape in $presentation.Slides.Item(1).Shapes) {
      if($shape.HasTextFrame -eq -1){$text+=$shape.TextFrame.TextRange.Text}
      if($shape.Type -eq 13){$pictures++}
    }
    $notes=@()
    foreach($shape in $presentation.Slides.Item(1).NotesPage.Shapes) {if($shape.HasTextFrame -eq -1){$notes+=$shape.TextFrame.TextRange.Text}}
    if($name -eq 'edited') {
      if(($text -join "`n") -notlike '*Browser revised title*' -or ($notes -join "`n") -notlike '*Browser revised notes*'){throw 'Native PowerPoint did not retain edited text and notes.'}
      $title=$presentation.Slides.Item(1).Shapes.Item(1)
      if([Math]::Abs($title.Left-112.5) -gt 0.01){throw 'Native PowerPoint did not retain edited geometry.'}
      if([Math]::Abs($title.Rotation-30) -gt 0.01){throw 'Native PowerPoint did not retain the edited rotation.'}
    } elseif($pictures -ne 1 -or ($notes -join "`n") -notlike '*First browser notes*'){throw 'Native PowerPoint did not retain the inserted picture and first notes.'}
    $pdf=[IO.Path]::ChangeExtension($path,'.pdf');$presentation.SaveAs($pdf,32)
    $reports+=[ordered]@{name=$name;application='Microsoft PowerPoint';version=$app.Version;inputSha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();slides=$presentation.Slides.Count;width=$presentation.PageSetup.SlideWidth;height=$presentation.PageSetup.SlideHeight;chart=$chart;pictures=$pictures;text=$text;notes=$notes;pdf=$pdf}
    $presentation.Close();$presentation=$null
  }
  $reports|ConvertTo-Json -Depth 8|Set-Content -LiteralPath .local/pptx-validation/desktop-report.json -Encoding UTF8
  Write-Output 'Native PowerPoint opened both exports and retained text, geometry, notes, image, chart and 4:3 dimensions; PDFs rendered.'
} finally {
  if($null -ne $presentation){try{$presentation.Close()}catch{Write-Warning $_}}
  if($null -ne $app){try{$app.DisplayAlerts=$alerts;$app.AutomationSecurity=$security;if($app.Presentations.Count -eq 0){$app.Quit()}}catch{Write-Warning $_};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
