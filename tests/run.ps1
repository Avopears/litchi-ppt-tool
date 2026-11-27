# Builds and runs the export tests against the PowerPoint installed on this machine.
# Usage: powershell -ExecutionPolicy Bypass -File tests\run.ps1
# It starts PowerPoint, builds a throw-away presentation, exports many kinds of selections
# and prints one line per case. Results are written to tests\out.
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
$core = 'Strings.cs', 'SvgTools.cs', 'PdfCropper.cs', 'Exporter.cs' | ForEach-Object { Join-Path $root "src\$_" }

& $csc /nologo /codepage:65001 /target:exe "/out:$exe" "/link:$(Find-Pia 'Microsoft.Office.Interop.PowerPoint')" "/link:$(Find-Pia 'office')" `
    /r:System.Drawing.dll $core (Join-Path $root 'tests\TestHarness.cs')
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
& $exe (Join-Path $out 'files')
