# Locates signtool.exe (matching the current machine's processor architecture) from the
# highest-versioned installed Windows 10/11 SDK. Falls back to the signtool.exe that ships
# with Visual Studio (Microsoft SDKs\ClickOnce\SignTool) if no Windows 10/11 SDK is installed.
# Writes the full path to signtool.exe to stdout, or nothing if not found.
$sdkBinRoot = 'C:\Program Files (x86)\Windows Kits\10\bin'

$arch = $env:PROCESSOR_ARCHITECTURE
switch ($arch) {
	'AMD64' { $toolArch = 'x64' }
	'ARM64' { $toolArch = 'arm64' }
	'x86'   { $toolArch = 'x86' }
	default { $toolArch = 'x64' }
}

if (Test-Path $sdkBinRoot) {
	$dir = Get-ChildItem -Path $sdkBinRoot -Directory -ErrorAction SilentlyContinue |
		Where-Object { Test-Path (Join-Path $_.FullName "$toolArch\signtool.exe") } |
		Sort-Object Name -Descending |
		Select-Object -First 1

	if ($dir) {
		Write-Output (Join-Path $dir.FullName "$toolArch\signtool.exe")
		exit 0
	}
}

$fallbackPaths = @(
	'C:\Program Files (x86)\Microsoft SDKs\ClickOnce\SignTool\signtool.exe',
	'C:\Program Files\Microsoft SDKs\ClickOnce\SignTool\signtool.exe'
)
foreach ($fallback in $fallbackPaths) {
	if (Test-Path $fallback) {
		Write-Output $fallback
		exit 0
	}
}

exit 0
