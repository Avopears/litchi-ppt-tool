# 卸载“选区导出”插件：删除注册表登记和已复制的 DLL。
$ErrorActionPreference = 'Stop'
$progId = 'PptSelectionExport.Connect'
$clsid  = '{A0181EA2-7061-43C3-9086-B72894478235}'

if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    Write-Host '检测到 PowerPoint 正在运行，请先关闭所有 PowerPoint 窗口，然后重新运行卸载。' -ForegroundColor Yellow
    Write-Host 'PowerPoint is running. Close every PowerPoint window, then run the uninstaller again.' -ForegroundColor Yellow
    exit 1
}

$keys = @(
    "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$progId",
    "HKCU:\Software\Classes\$progId",
    "HKCU:\Software\Classes\CLSID\$clsid",
    "HKCU:\Software\Classes\Wow6432Node\CLSID\$clsid",
    'HKCU:\Software\PptSelectionExport'
)
foreach ($k in $keys) {
    if (Test-Path $k) { Remove-Item -Path $k -Recurse -Force }
}

$target = Join-Path $env:LOCALAPPDATA 'PptSelectionExport'
if (Test-Path $target) { Remove-Item -Path $target -Recurse -Force }

Write-Host '已卸载。 / Uninstalled.' -ForegroundColor Green
