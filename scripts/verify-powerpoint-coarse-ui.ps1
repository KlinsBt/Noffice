param([switch]$Compare)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root '.local/powerpoint-coarse-ui'
[void][IO.Directory]::CreateDirectory($directory)
. (Join-Path $root 'scripts/excel-test-window.ps1')
function Retry([scriptblock]$action){for($attempt=0;$attempt -lt 24;$attempt++){try{return (& $action)}catch{if($_.Exception.ToString() -notmatch 'RPC_E_CALL_REJECTED|0x80010001|0x8001010A' -or $attempt -eq 23){throw};if($script:officeProcess){[ExcelTestWindow]::DismissReminder($script:officeProcess)};Start-Sleep -Milliseconds 250}}}
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Transforms($path){
 $zip=[IO.Compression.ZipFile]::OpenRead($path);$result=@()
 try {foreach($i in 1..4){
  $reader=New-Object IO.StreamReader($zip.GetEntry("ppt/slides/slide$i.xml").Open())
  try{[xml]$xml=$reader.ReadToEnd()}finally{$reader.Dispose()}
  $frames=@();foreach($group in $xml.SelectNodes("//*[local-name()='grpSp']")){
   $id=$group.SelectSingleNode("./*[local-name()='nvGrpSpPr']/*[local-name()='cNvPr']").GetAttribute('id')
   $frame=[ordered]@{id=$id}
   foreach($node in $group.SelectSingleNode("./*[local-name()='grpSpPr']/*[local-name()='xfrm']").ChildNodes){foreach($attr in $node.Attributes){$frame["$($node.LocalName).$($attr.LocalName)"]=[double]$attr.Value}}
   $frames+=,$frame
  };$result+=,$frames
 };return ,$result}finally{$zip.Dispose()}
}
function Hash($p){(Get-FileHash -LiteralPath $p).Hash.ToLowerInvariant()}
function Frame($shape){
 for($attempt=0;$attempt -lt 30;$attempt++){
  try{
   $books=$app.Presentations
   if($attempt -eq 0){Write-Host "Frame diagnostic: app $($app.GetType().FullName), presentations $($books.Count), shape $($shape.GetType().FullName)"}
   if(!$books -or $books.Count -ne 1){Start-Sleep -Milliseconds 100;continue}
   $slides=$books.Item(1).Slides
   if(!$slides -or $slides.Count -ne 4){Start-Sleep -Milliseconds 100;continue}
   $shape=FindLeaf $slides.Item(1).Shapes 2
   if(!$shape){Start-Sleep -Milliseconds 100;continue}
   $id=$shape.Id;$x=$shape.Left;$y=$shape.Top;$w=$shape.Width;$h=$shape.Height;$rotation=$shape.Rotation;$name=$shape.Name
   if($null -ne $id -and $id -gt 0 -and $null -ne $x -and $null -ne $y -and $w -gt 0 -and $h -gt 0 -and $null -ne $rotation){return [ordered]@{id=[int]$id;name=[string]$name;x=[double]$x;y=[double]$y;w=[double]$w;h=[double]$h;rotation=[double]$rotation}}
  }catch{if($_.Exception.ToString() -notmatch 'RPC_E_CALL_REJECTED|0x80010001|0x8001010A'){throw}}
  Start-Sleep -Milliseconds 100
 }
 throw 'Native shape frame did not become available; refusing null/zero evidence.'
}
function Shapes($shapes,$parents){
 $result=@()
 for($shapeIndex=1;$shapeIndex -le $shapes.Count;$shapeIndex++){
  $shape=$shapes.Item($shapeIndex)
  $frame=Frame $shape;$frame['parents']=$parents;$frame['group']=($shape.Type -eq 6);$result+=,$frame
  if($shape.Type -eq 6){$result+=Shapes $shape.GroupItems @($parents+[int]$shape.Id)}
 }
 return $result
}
function Snapshot($deck){
 $result=@()
 foreach($slide in $deck.Slides){$result+=,@(Shapes $slide.Shapes @())}
 return ,$result
}
function SelectBounds($deck){
 $result=@();$window=$deck.Windows.Item(1)
 foreach($slide in $deck.Slides){
  $window.View.GotoSlide($slide.SlideIndex)
  $group=$slide.Shapes.Item(1);$group.Select(-1)
  $selection=$window.Selection.ShapeRange
  $result+=,[ordered]@{slide=[int]$slide.SlideIndex;count=[int]$selection.Count;x=[double]$selection.Left;y=[double]$selection.Top;w=[double]$selection.Width;h=[double]$selection.Height;rotation=[double]$selection.Rotation}
 }
 return ,$result
}
function FindLeaf($shapes,$id){foreach($shape in $shapes){if($shape.Type -eq 6){$found=FindLeaf $shape.GroupItems $id;if($found){return $found}}elseif($shape.Id -eq $id){return $shape}}}
function Ungroup($deck){foreach($slide in $deck.Slides){$more=$true;while($more){$more=$false;for($i=$slide.Shapes.Count;$i -ge 1;$i--){if($slide.Shapes.Item($i).Type -eq 6){[void]$slide.Shapes.Item($i).Ungroup();$more=$true}}}}}
function WriteJson($name,$value){[IO.File]::WriteAllText((Join-Path $directory $name),($value|ConvertTo-Json -Depth 24),(New-Object Text.UTF8Encoding($false)))}
$app=$null;$deck=$null
try{
 $baseline=Get-Content -Raw (Join-Path $root 'docs/parity/baseline.json')|ConvertFrom-Json
 $native=$baseline.applications|Where-Object executable -eq 'POWERPNT.EXE'
 $app=New-Object -ComObject PowerPoint.Application
 if($app.Presentations.Count -ne 0){throw 'The native test requires an idle owned PowerPoint instance.'}
 $processes=@(Get-CimInstance Win32_Process -Filter "name = 'POWERPNT.EXE'")
 if($processes.Count -ne 1){throw 'Cannot uniquely identify the owned PowerPoint process.'}
 $script:officeProcess=[uint32]$processes[0].ProcessId
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$app.AutomationSecurity=3;$app.DisplayAlerts=1
 if((Hash (Join-Path ([string]$app.Path) 'POWERPNT.EXE')) -ne $native.sha256){throw 'The native baseline changed.'}

 Add-Type -AssemblyName UIAutomationClient
 Add-Type -AssemblyName UIAutomationTypes
 $source=Join-Path $root 'tests/fixtures/powerpoint-coarse.pptx'
 $deck=$app.Presentations.Open($source,0,0,0)
 if(!$deck){throw 'PowerPoint did not return the owned authored presentation.'}
 $window=$deck.NewWindow();Retry {$window.WindowState=3}
 Retry {$window.View.GotoSlide(1)}
 $leaf=FindLeaf $deck.Slides.Item(1).Shapes 2
 Retry {$leaf.Select(-1)}
 Retry {$app.CommandBars.ExecuteMso('ObjectSizeAndPositionDialog')}
 Start-Sleep -Milliseconds 700
 $nativeHandle=(Get-Process -Id $script:officeProcess).MainWindowHandle
 if($nativeHandle -eq [IntPtr]::Zero){throw "The owned native frame handle is unavailable."}
 $element=[Windows.Automation.AutomationElement]::FromHandle($nativeHandle)
 function Controls($element){
 $all=New-Object 'System.Collections.Generic.List[Windows.Automation.AutomationElement]'
 $queue=New-Object 'System.Collections.Generic.Queue[Windows.Automation.AutomationElement]'
 $queue.Enqueue($element);$failures=0
 while($queue.Count -gt 0 -and $all.Count -lt 2500){
  $current=$queue.Dequeue();$all.Add($current)
  try{$children=$current.FindAll([Windows.Automation.TreeScope]::Children,[Windows.Automation.Condition]::TrueCondition);foreach($child in $children){$queue.Enqueue($child)}}catch{$failures++}
 }
 return ,$all
 }
 $all=Controls $element
 $position=@($all|Where-Object {$_.Current.Name -eq 'Position' -and $_.Current.ControlType.ProgrammaticName -eq 'ControlType.Button'})
 if($position.Count -ne 1){throw 'Expected one native Position expander.'}
 $toggle=$position[0].GetCurrentPattern([Windows.Automation.TogglePattern]::Pattern);if($toggle.Current.ToggleState -eq [Windows.Automation.ToggleState]::Off){$toggle.Toggle()}
 Start-Sleep -Milliseconds 300
 $all=Controls $element
 $rows=@()
 foreach($el in $all){
  $c=$el.Current
  if($c.ControlType.ProgrammaticName -match 'Edit|Pane|Button|ComboBox|CheckBox|RadioButton|Group'){
   $value='';$pattern=$null
   if($el.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern,[ref]$pattern)){$value=$pattern.Current.Value}
   $rows+=,@{name=$c.Name;type=$c.ControlType.ProgrammaticName;id=$c.AutomationId;enabled=$c.IsEnabled;offscreen=$c.IsOffscreen;value=$value}
  }
 }
 WriteJson 'ui-tree.json' $rows

 Add-Type -AssemblyName System.Windows.Forms
 Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class OwnedPowerPointInput {
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint id);
}
'@
 function Field($name){
  $items=@((Controls $element)|Where-Object {$_.Current.Name -eq $name -and $_.Current.ControlType.ProgrammaticName -eq 'ControlType.Edit'})
  if($items.Count -ne 1){throw "Expected one native field: $name"};return $items[0]
 }
 function Keys($keys){
  [uint32]$front=0;[void][OwnedPowerPointInput]::GetWindowThreadProcessId([OwnedPowerPointInput]::GetForegroundWindow(),[ref]$front)
  if($front -ne $script:officeProcess){throw 'Refusing keyboard input outside the owned native window.'}
  [Windows.Forms.SendKeys]::SendWait($keys)
 }
 function ReadField($name){return (Field $name).GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).Current.Value}
 $stages=@();$stages+=,@{stage='source';frame=(Frame $leaf);horizontal=(ReadField 'Horizontale Position');vertical=(ReadField 'Vertikale Position')}
 foreach($edit in @(@{field='Horizontale Position';value='4,2897777778 cm';stage='left'},@{field='Vertikale Position';value='4,9829861111 cm';stage='top'})){
  [void][OwnedPowerPointInput]::SetForegroundWindow($nativeHandle)
  $field=Field $edit.field;$field.SetFocus();Keys '^a';Keys $edit.value
  $before=$field.GetCurrentPattern([Windows.Automation.ValuePattern]::Pattern).Current.Value
  Keys '{TAB}';Start-Sleep -Milliseconds 250
  $stages+=,@{stage=$edit.stage;entered=$edit.value;beforeCommit=$before;afterCommit=(ReadField $edit.field);frame=(Frame $leaf)}
 }
 WriteJson 'before-history.json' $stages
 Retry {$leaf.Select(-1)}
 $canvas=@((Controls $element)|Where-Object {$_.Current.Name -eq 'Folie' -and $_.Current.ControlType.ProgrammaticName -eq 'ControlType.Pane'})
 if($canvas.Count -ne 1){throw 'Expected one owned native slide surface.'};$canvas[0].SetFocus()
 foreach($action in @('Undo','Undo','Redo','Redo')){
  if(!$app.CommandBars.GetEnabledMso($action)){$stages+=,@{stage=$action;enabled=$false;frame=(Frame $leaf)};continue}
  Keys $(if($action -eq 'Undo'){'^z'}else{'^y'});Start-Sleep -Milliseconds 150
  $stages+=,@{stage=$action;frame=(Frame $leaf)}
 }
 $output=Join-Path $directory 'ui-edited.pptx';$deck.SaveAs($output,24);$deck.Close();$deck=$null
 $deck=$app.Presentations.Open($output,-1,0,0);$leaf=FindLeaf $deck.Slides.Item(1).Shapes 2
 $stages+=,@{stage='reopened';frame=(Frame $leaf)}
 $deck.SaveAs((Join-Path $directory 'ui-edited.pdf'),32)
 WriteJson 'ui-edit-report.json' @{sourceSha256=Hash $source;scriptSha256=Hash $PSCommandPath;exportSha256=Hash $output;pdfSha256=Hash (Join-Path $directory 'ui-edited.pdf');executableSha256=$native.sha256;stages=$stages}

 WriteJson 'snapshot.json' @{sourceSha256=Hash $source;scriptSha256=Hash $PSCommandPath;frame=(Frame $leaf);executableSha256=$native.sha256}
 Write-Output "Recorded owned native UI input, history and reopened frames."
}finally{
 if($deck){Retry {$deck.Saved=-1;$deck.Close()}}
 if($app){if($null -ne $security){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;if($app.Presentations.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
