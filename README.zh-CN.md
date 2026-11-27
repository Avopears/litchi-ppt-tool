# Litchi PPT Tool

[English](README.md) | **简体中文**

**选区导出**（PowerPoint 插件）——给习惯用 PowerPoint 画图的人准备的小工具：在幻灯片上选中几个对象，点一下按钮，
就能把**只有这些对象**导出成 **PDF**、**SVG**、**PNG** 或 **JPG**，
范围是刚好包住它们的最小矩形，背景可以选透明或任意纯色。

![插件在 PowerPoint 里的“Litchi”选项卡](docs/ribbon.jpg)

（中文版 Office 显示中文界面，其他语言的 Office 显示英文界面。）

## 为什么做这个

写论文、毕业设计或 LaTeX 文档时，要把幻灯片里的一张图单独拿出来，通常得导出整页再手动裁剪，
或者把幻灯片尺寸改成和图一样大。这个插件一步完成：

- **贴边裁剪。** 输出范围就是选中内容的外接矩形，和 PowerPoint 自带的“另存为图片”一致，
  线条粗细、箭头、阴影都算在内，边框不会被切掉一半。
- **只导出选中的对象。** 没选中的对象、幻灯片背景、母版上的图案都不会出现。
- **PDF 和 SVG 是矢量图。** 文字还是文字（PDF 内嵌字体），形状还是形状。
- **PNG / JPG 可选 96 – 600 DPI**，并把 DPI 写进文件，插入到别的软件里物理尺寸是对的。
- **背景透明或纯色**，所有格式都支持（JPG 不支持透明，选了透明时用白色）。
- **不动你的文档。** 不修改演示文稿，不影响撤销记录和剪贴板。

## 安装

要求：Windows，桌面版 PowerPoint（Microsoft 365 或 Office 2019 / 2021 / 2024，32 位、64 位均可）。
不需要管理员权限。

1. 下载 **[Litchi-Setup.exe](https://github.com/Avopears/litchi-ppt-tool/releases/latest/download/Litchi-Setup.exe)**
   （也可以在 [Releases](https://github.com/Avopears/litchi-ppt-tool/releases) 页面选版本）。
2. 关闭 PowerPoint，运行安装程序。安装位置可以用默认的，也可以自己选文件夹。
3. 打开 PowerPoint，插件的选项卡叫 **Litchi**，就在“视图”后面。

![安装程序](docs/installer.png)

安装程序没有代码签名，第一次运行时 Windows SmartScreen 可能提示“Windows 已保护你的电脑”，
点 **更多信息 → 仍要运行** 即可；不放心的话也可以自己从源码编译（见下文）。

卸载：Windows **设置 → 应用 → 安装的应用 → 选区导出 Export Selection for PowerPoint → 卸载**，
或者再运行一次安装程序，选“卸载”。

<details>
<summary>其他安装方式</summary>

- **不用 exe：** 把本仓库下载为 ZIP 并解压，关闭 PowerPoint 后双击 `安装.bat`（`卸载.bat` 用来卸载）。
- **静默安装：** `Litchi-Setup.exe /S` 安装（加 `/DIR="D:\Tools\Litchi"` 可指定位置），
  `Litchi-Setup.exe /uninstall /S` 卸载。

所有内容都只装给当前用户：插件复制到 `%LOCALAPPDATA%\PptSelectionExport` 或你选的文件夹
（如果那个文件夹里已有别的东西，会在其中新建 `PptSelectionExport` 文件夹），注册信息写在 `HKEY_CURRENT_USER` 下。
卸载时只删除插件自己的文件，不会动该位置里的其他东西。
</details>

## 使用

1. 在幻灯片上选中要导出的对象（Shift / Ctrl 点选，或拖框多选），也可以只选组合里的某几个子对象。
2. 在“背景”组里选 **透明背景** 或 **纯色**（点“纯色”下面的小箭头换颜色，有预设色和“其他颜色…”）。
3. 点 **导出 PDF**、**导出 SVG** 或 **导出图片**，选择保存位置。
   “导出图片”在保存对话框里选 PNG 或 JPG，清晰度由 **图片分辨率** 决定。

## 需要知道的事

- 导出 PDF 需要在普通视图里选中幻灯片上的对象；母版、版式、备注页里的对象可以导出 SVG 或图片。
- 阴影、发光等效果在 PDF 里是位图，这是 PowerPoint 自身的行为。
- SVG 里的文字仍是文字，按字体名引用字体；别的电脑没有这个字体时会被替换。对外观要求严格时请用 PDF。
- 导出 PDF 时插件会先存一份演示文稿的临时副本，所以特别大的文件（几百 MB）要多等一会儿。
- 选中内容超过 56 英寸（PowerPoint 幻灯片的最大尺寸）时，PDF 会被裁到这个范围。

已在 Windows 11 上的 Microsoft 365（64 位、中文界面、版本 16.0.20625）测试过。
英文界面和其他 Office 版本按设计应当可用，但还没有实际测试，
欢迎[反馈问题](https://github.com/Avopears/litchi-ppt-tool/issues)。

## 从源码编译

不需要 Visual Studio，用 Windows 自带的 C# 编译器（.NET Framework 4）加上已安装的 Office 就够了：

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

会生成 `dist\PptSelectionExport.dll`（插件本体）和 `dist\Litchi-Setup.exe`（安装程序，DLL 已打包在里面）。

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\run.ps1
```

会用本机的 PowerPoint 跑一遍导出测试：新建一个临时演示文稿，把各种选区按每种格式导出一遍。

| 文件 | 作用 |
| --- | --- |
| `src/Connect.cs` | 加载项入口、选项卡按钮的响应 |
| `src/Ribbon.xml` | 选项卡界面定义 |
| `src/Exporter.cs` | SVG、PNG/JPG、PDF 三条导出流程 |
| `src/PdfCropper.cs` | 把 PowerPoint 生成的整页 PDF 精确裁到选区（PDF 增量更新） |
| `src/SvgTools.cs` | 读取选区范围、给 SVG 加背景 |
| `src/Strings.cs` | 全部界面文字（中文和英文） |
| `src/Icons.cs`、`src/Settings.cs` | 用代码画的图标、记住的设置 |
| `installer/Setup.cs` | exe 安装 / 卸载程序 |
| `tests/TestHarness.cs` | 直接驱动 PowerPoint 的导出测试 |

### 实现原理

- **SVG 和 PNG/JPG** 用的是 PowerPoint 自己的 `ShapeRange.Export`，它本身就按可见范围裁剪并保留透明。
  插件在此基础上加背景、`viewBox`、指定分辨率和 DPI 标记。
- **PDF**：PowerPoint 只能导出整张幻灯片。插件先存一份临时副本，把没选中的对象全部隐藏、去掉幻灯片背景，
  导出这一页，再往 PDF 末尾追加一段“增量更新”，把页面框改成选区范围。全程不栅格化，也不改写原有字节。

## 许可证

[MIT](LICENSE)
