# 安装“选区导出”插件（仅当前用户，不需要管理员权限）。
# 做两件事：1) 把 DLL 复制到 %LOCALAPPDATA%\Litchi Tool
#           2) 在当前用户的注册表里登记这个 COM 加载项，让 PowerPoint 启动时加载它
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$progId = 'PptSelectionExport.Connect'
$clsid  = '{A0181EA2-7061-43C3-9086-B72894478235}'
$dll    = Join-Path $root 'dist\PptSelectionExport.dll'
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    Write-Host '检测到 PowerPoint 正在运行，请先保存并关闭所有 PowerPoint 窗口，然后重新运行安装。' -ForegroundColor Yellow
    Write-Host 'PowerPoint is running. Save your work, close every PowerPoint window, then run the installer again.' -ForegroundColor Yellow
    exit 1
}

# 从源码仓库直接安装时还没有 DLL：先编译一次（只用 Windows 自带的编译器）
if (-not (Test-Path $dll)) {
    Write-Host '没有找到编译好的 DLL，先从源码编译… / No prebuilt DLL found, building from source…'
    & (Join-Path $root 'build.ps1')
    if (-not (Test-Path $dll)) { throw "编译失败，找不到 $dll / Build failed." }
}

$target = Join-Path $env:LOCALAPPDATA 'Litchi Tool'
New-Item -ItemType Directory -Force $target | Out-Null
Copy-Item $dll $target -Force
$installed = Join-Path $target 'PptSelectionExport.dll'
# 从网上下载的文件带有“来自 Internet”标记，.NET 会拒绝加载这样的 DLL，这里去掉标记
Unblock-File -Path $installed -ErrorAction SilentlyContinue
$asmName  = [Reflection.AssemblyName]::GetAssemblyName($installed)
$codeBase = 'file:///' + $installed.Replace('\', '/')   # 不做 URL 转义，路径里有中文或空格也能加载

# 先把旧版本的登记清干净，再写入新的
foreach ($old in "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$progId", "HKCU:\Software\Classes\$progId",
                 "HKCU:\Software\Classes\CLSID\$clsid", "HKCU:\Software\Classes\Wow6432Node\CLSID\$clsid") {
    if (Test-Path $old) { Remove-Item -Path $old -Recurse -Force }
}

function Set-Default([string]$key, [string]$value) {
    New-Item -Path $key -Force | Out-Null
    Set-Item -Path $key -Value $value
}

# 64 位和 32 位 Office 各看一个注册表视图，两边都登记
foreach ($classes in 'HKCU:\Software\Classes', 'HKCU:\Software\Classes\Wow6432Node') {
    $k = "$classes\CLSID\$clsid"
    if (Test-Path $k) { Remove-Item -Path $k -Recurse -Force }   # 清掉旧版本的登记
    Set-Default $k $progId
    Set-Default "$k\ProgId" $progId
    New-Item -Path "$k\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}" -Force | Out-Null
    foreach ($s in "$k\InprocServer32", "$k\InprocServer32\$($asmName.Version)") {
        New-Item -Path $s -Force | Out-Null
        Set-ItemProperty -Path $s -Name 'Class' -Value $progId
        Set-ItemProperty -Path $s -Name 'Assembly' -Value $asmName.FullName
        Set-ItemProperty -Path $s -Name 'RuntimeVersion' -Value 'v4.0.30319'
        Set-ItemProperty -Path $s -Name 'CodeBase' -Value $codeBase
    }
    Set-Item -Path "$k\InprocServer32" -Value 'mscoree.dll'
    Set-ItemProperty -Path "$k\InprocServer32" -Name 'ThreadingModel' -Value 'Both'
}
Set-Default "HKCU:\Software\Classes\$progId" $progId
Set-Default "HKCU:\Software\Classes\$progId\CLSID" $clsid

$addin = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$progId"
New-Item -Path $addin -Force | Out-Null
Set-ItemProperty -Path $addin -Name 'FriendlyName' -Value '选区导出 Export Selection (PDF / SVG / PNG / JPG)'
Set-ItemProperty -Path $addin -Name 'Description' -Value '把选中的对象按最小外接矩形导出为 PDF、SVG、PNG 或 JPG，可设置背景颜色或透明；另有对齐、旋转和布尔运算工具。 Export the selected objects, cropped to their bounding box, as PDF, SVG, PNG or JPG; plus align, rotate and merge-shape tools.'
Set-ItemProperty -Path $addin -Name 'LoadBehavior' -Value 3 -Type DWord

Write-Host "安装完成 / Installed: $installed" -ForegroundColor Green
Write-Host '现在打开 PowerPoint，功能区“视图”后面会出现“Litchi”选项卡。'
Write-Host 'Start PowerPoint: the "Litchi" tab appears after the View tab.'
