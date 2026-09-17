param([string]$Output = '.local/office-baseline.json')
$ErrorActionPreference = 'Stop'
# Read installed executable metadata only; never attach to a user's Office process.
$records = @()
foreach ($name in @('WINWORD.EXE','EXCEL.EXE','POWERPNT.EXE')) {
  $key = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$name"
  $path = (Get-ItemProperty -LiteralPath $key).'(default)'
  $file = Get-Item -LiteralPath $path
  $stream = [IO.File]::OpenRead($file.FullName)
  $reader = New-Object IO.BinaryReader($stream)
  try { $stream.Position=0x3c; $offset=$reader.ReadInt32(); $stream.Position=$offset+4; $machine=$reader.ReadUInt16() }
  finally { $reader.Dispose(); $stream.Dispose() }
  $records += [ordered]@{ executable=$name; path=$file.FullName; product=$file.VersionInfo.ProductName; version=$file.VersionInfo.FileVersion; architecture=$(if($machine -eq 34404){'x64'}elseif($machine -eq 332){'x86'}else{"PE-$machine"}); sha256=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant() }
}
$os = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$result = [ordered]@{ schemaVersion=1; inspectedAt=[DateTime]::UtcNow.ToString('o'); method='Executable file metadata, PE machine type and SHA-256; no Office documents opened'; culture=[Globalization.CultureInfo]::CurrentCulture.Name; uiCulture=[Globalization.CultureInfo]::CurrentUICulture.Name; osBuild="$($os.CurrentBuild).$($os.UBR)"; applications=$records }
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$target = [IO.Path]::GetFullPath((Join-Path $root $Output))
if(-not $target.StartsWith($root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Baseline output must stay in the workspace.'}
[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target)) | Out-Null
[IO.File]::WriteAllText($target,($result|ConvertTo-Json -Depth 8)+[Environment]::NewLine,(New-Object Text.UTF8Encoding($false)))
$result | ConvertTo-Json -Depth 8
