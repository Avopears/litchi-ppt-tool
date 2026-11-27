using System.Collections.Generic;
using System.Globalization;

namespace PptSelectionExport
{
    /// <summary>
    /// All user-visible text, in Chinese and English. The add-in follows the display
    /// language of Office: Chinese for Chinese editions, English for everything else.
    /// </summary>
    internal static class Strings
    {
        public static bool Chinese = CultureInfo.CurrentUICulture.Name.StartsWith("zh", System.StringComparison.OrdinalIgnoreCase);

        /// <summary>Office UI language ids (LCIDs) that should get the Chinese interface.</summary>
        public static bool IsChineseLcid(int lcid)
        {
            return (lcid & 0x3FF) == 0x04;   // primary language id: Chinese (zh-CN, zh-TW, zh-HK, zh-SG, zh-MO)
        }

        public static string T(string key)
        {
            string[] pair;
            if (!Table.TryGetValue(key, out pair)) return key;
            return Chinese ? pair[0] : pair[1];
        }

        public static string T(string key, params object[] args)
        {
            return string.Format(CultureInfo.InvariantCulture, T(key), args);
        }

        public static IEnumerable<string> Keys { get { return Table.Keys; } }

        private static readonly Dictionary<string, string[]> Table = new Dictionary<string, string[]>
        {
            // ---- ribbon
            { "tab", new[] { "Litchi", "Litchi" } },
            { "grp.export", new[] { "导出选中内容", "Export selection" } },
            { "pdf.label", new[] { "导出 PDF", "Export PDF" } },
            { "pdf.tip", new[] { "把选中内容导出为 PDF", "Export the selection as PDF" } },
            { "pdf.super", new[] {
                "把当前选中的对象导出为矢量 PDF。页面大小就是刚好包住所有选中对象的最小矩形（含线条粗细、箭头和阴影）。未选中的对象不会出现。",
                "Saves the selected objects as a vector PDF. The page is the smallest rectangle that contains everything drawn for them (outlines, arrowheads and shadows included). Objects that are not selected are left out." } },
            { "svg.label", new[] { "导出 SVG", "Export SVG" } },
            { "svg.tip", new[] { "把选中内容导出为 SVG", "Export the selection as SVG" } },
            { "svg.super", new[] {
                "把当前选中的对象导出为 SVG 矢量图。画布大小就是刚好包住所有选中对象的最小矩形（含线条粗细、箭头和阴影）。",
                "Saves the selected objects as an SVG vector image. The canvas is the smallest rectangle that contains everything drawn for them (outlines, arrowheads and shadows included)." } },
            { "img.label", new[] { "导出图片", "Export Image" } },
            { "img.tip", new[] { "把选中内容导出为 PNG / JPG 图片", "Export the selection as a PNG / JPG picture" } },
            { "img.super", new[] {
                "把当前选中的对象导出为 PNG 或 JPG 图片（在保存对话框里选格式），清晰度由“图片分辨率”决定。PNG 支持透明背景；JPG 不支持透明，选了透明背景时会用白色。",
                "Saves the selected objects as a PNG or JPG picture (pick the format in the save dialog) at the chosen image resolution. PNG can be transparent; JPG cannot, so a transparent background becomes white." } },
            { "grp.background", new[] { "背景", "Background" } },
            { "transparent.label", new[] { "透明背景", "Transparent" } },
            { "transparent.super", new[] { "导出的文件不带任何底色。", "The exported file has no background at all." } },
            { "color.label", new[] { "纯色", "Color" } },
            { "color.tip", new[] { "纯色背景", "Solid background" } },
            { "color.super", new[] {
                "用纯色铺满整个矩形作为底色。当前颜色：#{0}。点下面的箭头可以换颜色。",
                "Fills the whole rectangle with one colour. Current colour: #{0}. Use the arrow below to change it." } },
            { "color.white", new[] { "白色", "White" } },
            { "color.lightgray", new[] { "浅灰", "Light gray" } },
            { "color.black", new[] { "黑色", "Black" } },
            { "color.cream", new[] { "米黄", "Cream" } },
            { "color.paleblue", new[] { "浅蓝", "Pale blue" } },
            { "color.more", new[] { "其他颜色…", "More colors…" } },
            { "grp.options", new[] { "选项", "Options" } },
            { "dpi.label", new[] { "图片分辨率", "Image resolution" } },
            { "dpi.super", new[] {
                "导出 PNG / JPG 时每英寸的像素数。数值越大越清晰、文件也越大：96 适合屏幕，300 适合论文和打印，600 适合精细线条图。对 PDF 和 SVG 没有影响（它们是矢量图）。",
                "Pixels per inch for PNG / JPG export. Higher is sharper and bigger: 96 for screens, 300 for papers and print, 600 for fine line art. It does not affect PDF or SVG, which are vector formats." } },
            { "openfolder.label", new[] { "完成后打开文件夹", "Show file when done" } },
            { "openfolder.super", new[] { "导出成功后在资源管理器中定位到导出的文件。", "Opens File Explorer with the exported file selected." } },

            // ---- dialogs
            { "title", new[] { "选区导出", "Export Selection" } },
            { "dlg.pdf", new[] { "导出选中内容为 PDF", "Export selection as PDF" } },
            { "dlg.svg", new[] { "导出选中内容为 SVG", "Export selection as SVG" } },
            { "dlg.img", new[] { "导出选中内容为图片", "Export selection as image" } },
            { "filter.pdf", new[] { "PDF 文件 (*.pdf)|*.pdf", "PDF file (*.pdf)|*.pdf" } },
            { "filter.svg", new[] { "SVG 矢量图 (*.svg)|*.svg", "SVG vector image (*.svg)|*.svg" } },
            { "filter.img", new[] { "PNG 图片 (*.png)|*.png|JPEG 图片 (*.jpg)|*.jpg", "PNG image (*.png)|*.png|JPEG image (*.jpg)|*.jpg" } },
            { "name.selection", new[] { "选区", "selection" } },

            // ---- messages
            { "msg.nowindow", new[] {
                "请先打开一个演示文稿，并在普通编辑视图中选中要导出的对象。\n（受保护的视图、放映视图下无法导出。）",
                "Open a presentation and select the objects to export in the normal editing view.\n(Exporting is not available in Protected View or during a slide show.)" } },
            { "msg.noselection", new[] {
                "请先在幻灯片上选中要导出的对象。\n可以按住 Shift / Ctrl 逐个点选，或用鼠标拖出选框一次选中多个。",
                "Select the objects to export on the slide first.\nShift- or Ctrl-click to pick several, or drag a selection box around them." } },
            { "msg.pdfneedsslide", new[] {
                "导出 PDF 需要在普通视图中选中幻灯片上的对象。\n（母版、版式、备注页中的对象请改用“导出 SVG”或“导出图片”。）",
                "PDF export needs objects selected on a slide in the normal view.\n(For objects on a master, layout or notes page use Export SVG or Export Image instead.)" } },
            { "msg.failed", new[] { "导出失败：\n{0}", "Export failed:\n{0}" } },
            { "msg.locked", new[] {
                "\n\n如果目标文件正被其他程序打开（例如 PDF 阅读器或看图软件），请先关闭它再导出。",
                "\n\nIf the target file is open in another program (a PDF reader or image viewer, for example), close it and try again." } },

            // ---- errors raised by the export code
            { "err.svg", new[] { "PowerPoint 导出的 SVG 内容无法识别。", "The SVG written by PowerPoint could not be understood." } },
            { "err.bounds", new[] { "无法确定选中内容的范围。", "Could not determine the bounds of the selection." } },
            { "err.notfound", new[] { "在临时副本中找不到选中的对象，请重新选择后再试。", "The selected objects could not be found in the temporary copy. Select them again and retry." } },
            { "err.pdf", new[] { "PDF 结构无法识别（{0}）。", "The PDF written by PowerPoint has an unexpected structure ({0})." } },
            { "err.outside", new[] { "选中的内容不在幻灯片范围内，无法导出。", "The selection lies outside the slide and cannot be exported." } },
            { "err.toolarge", new[] {
                "图片太大（{0} × {1} 像素）。请在“图片分辨率”里选一个更小的数值，或者少选一些内容。",
                "The picture would be too large ({0} × {1} pixels). Choose a lower image resolution or select less." } },
            { "err.image", new[] { "PowerPoint 没有生成图片，请重新选择后再试。", "PowerPoint did not produce a picture. Select the objects again and retry." } },
        };
    }
}
