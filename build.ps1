# 编译插件：生成 dist\PptSelectionExport.dll
# 只需要 Windows 自带的 .NET Framework 4 编译器和已安装的 Office，不需要 Visual Studio。
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$csc = @("$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe", "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe") |
    Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $csc) { throw '找不到 .NET Framework 4 的编译器 csc.exe。' }

function Find-Pia([string]$name) {
    $dir = Join-Path $env:WINDIR "assembly\GAC_MSIL\$name"
    $dll = Get-ChildItem $dir -Recurse -Filter *.dll -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -First 1
    if (-not $dll) { throw "找不到 Office 互操作程序集 $name（需要先安装 Microsoft Office）。" }
    $dll.FullName
}
$piaPpt = Find-Pia 'Microsoft.Office.Interop.PowerPoint'
$piaOffice = Find-Pia 'office'

$dist = Join-Path $root 'dist'
New-Item -ItemType Directory -Force $dist | Out-Null
$out = Join-Path $dist 'PptSelectionExport.dll'
$src = Get-ChildItem (Join-Path $root 'src') -Filter *.cs | ForEach-Object { $_.FullName }

& $csc /nologo /codepage:65001 /target:library /platform:anycpu /optimize+ /warn:4 `
    "/out:$out" "/link:$piaPpt" "/link:$piaOffice" `
    /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll `
    "/resource:$(Join-Path $root 'src\Ribbon.xml'),PptSelectionExport.Ribbon.xml" `
    $src
if ($LASTEXITCODE -ne 0) { throw '编译失败。' }
Write-Host "编译完成：$out"

# 安装程序：把上面的 DLL 作为资源打进一个 exe，文件名带版本号（Litchi-Setup-1.4.1.exe）
function Get-Version([string]$file) {
    $m = [regex]::Match((Get-Content $file -Raw -Encoding UTF8), 'AssemblyVersion\("(\d+\.\d+\.\d+)')
    if (-not $m.Success) { throw "在 $file 里找不到版本号。" }
    $m.Groups[1].Value
}
$version = Get-Version (Join-Path $root 'installer\Setup.cs')
$addinVersion = Get-Version (Join-Path $root 'src\Connect.cs')
if ($version -ne $addinVersion) { throw "版本号不一致：安装程序 $version，插件 $addinVersion。请把 installer\Setup.cs 和 src\Connect.cs 改成同一个版本。" }
Get-ChildItem $dist -Filter 'Litchi-Setup*.exe' | Remove-Item -Force     # 只保留当前版本的安装包
$setup = Join-Path $dist "Litchi-Setup-$version.exe"
& $csc /nologo /codepage:65001 /target:winexe /platform:anycpu /optimize+ /warn:4 `
    "/out:$setup" "/win32icon:$(Join-Path $root 'installer\setup.ico')" `
    /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll `
    "/resource:$out,PptSelectionExport.dll" `
    (Join-Path $root 'installer\Setup.cs')
if ($LASTEXITCODE -ne 0) { throw '安装程序编译失败。' }
Write-Host "安装程序：$setup"
