# Builds and runs the export, alignment and shape tests against the PowerPoint installed on this machine.
# Usage: powershell -ExecutionPolicy Bypass -File tests\run.ps1 [align]
# It starts PowerPoint, builds throw-away presentations, exports many kinds of selections,
# aligns and distributes shapes (comparing with PowerPoint's own commands) and prints one line
# per case. Exported files are written to tests\out. "align" runs everything but the export tests;
# "align matrix", "align size" or "align slim" only the matrix layout / cut, slide size or Slim Down tests.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

$csc = @("$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe", "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe") |
    Where-Object { Test-Path $_ } | Select-Object -First 1
function Find-Pia([string]$name) {
    (Get-ChildItem (Join-Path $env:WINDIR "assembly\GAC_MSIL\$name") -Recurse -Filter *.dll | Sort-Object FullName -Descending | Select-Object -First 1).FullName
}

$out = Join-Path $root 'tests\out'
New-Item -ItemType Directory -Force $out | Out-Null
$exe = Join-Path $out 'TestHarness.exe'
$core = 'Strings.cs', 'SvgTools.cs', 'PdfCropper.cs', 'Exporter.cs', 'Aligner.cs', 'ShapeTools.cs', 'StyleTools.cs', 'LineCaps.cs', 'MatrixTools.cs', 'Slimmer.cs', 'SlideResizer.cs', 'ColorPicker.cs', 'Settings.cs' |
    ForEach-Object { Join-Path $root "src\$_" }
$tests = 'TestHarness.cs', 'AlignTests.cs', 'ShapeToolsTests.cs', 'StyleTests.cs', 'CapTests.cs', 'MatrixTests.cs', 'ResizeTests.cs', 'SlimTests.cs' | ForEach-Object { Join-Path $root "tests\$_" }

& $csc /nologo /codepage:65001 /target:exe "/out:$exe" "/link:$(Find-Pia 'Microsoft.Office.Interop.PowerPoint')" "/link:$(Find-Pia 'office')" `
    /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Xml.dll /r:System.IO.Compression.dll $core $tests
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
& $exe (Join-Path $out 'files') @args
