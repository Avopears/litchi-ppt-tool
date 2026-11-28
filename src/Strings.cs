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

        // the part every align tooltip shares: what the objects are aligned to
        private const string AlignRefZh = "只选一个对象时对齐到幻灯片；打开“对齐到对象”后，以第一个选中的对象为准，它本身不动。";
        private const string AlignRefEn = " One object on its own is aligned to the slide. With Align to Object on, the first object you selected stays put and the others align to it.";

        // the part every boolean tooltip shares
        private const string MergeRefZh = "结果沿用第一个选中的形状的格式（填充、线条）：先单击它，再按住 Shift 或 Ctrl 单击其他形状。至少选 2 个形状。";
        private const string MergeRefEn = " The result keeps the format of the first shape you select: click it first, then Shift- or Ctrl-click the others. Select at least 2 shapes.";

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

            // ---- ribbon: align
            { "grp.align", new[] { "对齐", "Align" } },
            { "align.left.label", new[] { "左对齐", "Align Left" } },
            { "align.left.super", new[] { "让选中的对象左边缘对齐，对齐到最左边的对象。" + AlignRefZh, "Lines up the left edges of the selected objects with the leftmost one." + AlignRefEn } },
            { "align.center.label", new[] { "左右居中", "Align Center" } },
            { "align.center.super", new[] { "让选中的对象在左右方向上居中，对齐到它们整体范围的竖直中线。" + AlignRefZh, "Lines up the selected objects on the vertical centre line of the whole selection." + AlignRefEn } },
            { "align.right.label", new[] { "右对齐", "Align Right" } },
            { "align.right.super", new[] { "让选中的对象右边缘对齐，对齐到最右边的对象。" + AlignRefZh, "Lines up the right edges of the selected objects with the rightmost one." + AlignRefEn } },
            { "align.top.label", new[] { "上对齐", "Align Top" } },
            { "align.top.super", new[] { "让选中的对象上边缘对齐，对齐到最上面的对象。" + AlignRefZh, "Lines up the top edges of the selected objects with the highest one." + AlignRefEn } },
            { "align.middle.label", new[] { "上下居中", "Align Middle" } },
            { "align.middle.super", new[] { "让选中的对象在上下方向上居中，对齐到它们整体范围的水平中线。" + AlignRefZh, "Lines up the selected objects on the horizontal centre line of the whole selection." + AlignRefEn } },
            { "align.bottom.label", new[] { "下对齐", "Align Bottom" } },
            { "align.bottom.super", new[] { "让选中的对象下边缘对齐，对齐到最下面的对象。" + AlignRefZh, "Lines up the bottom edges of the selected objects with the lowest one." + AlignRefEn } },
            { "distribute.h.label", new[] { "横向分布", "Distribute Horizontally" } },
            { "distribute.h.super", new[] {
                "让选中的对象左右间距相等：最左和最右的对象不动，中间的对象等距排开（至少选 3 个）。只选一个对象时在幻灯片上左右居中；打开“对齐到对象”后，其他对象在第一个选中的对象的宽度范围内等距排开。",
                "Spaces the selected objects evenly from left to right: the outermost two stay put and the ones between them move (select at least 3). One object on its own is centred on the slide. With Align to Object on, the other objects are spread evenly across the width of the first object you selected." } },
            { "distribute.v.label", new[] { "纵向分布", "Distribute Vertically" } },
            { "distribute.v.super", new[] {
                "让选中的对象上下间距相等：最上和最下的对象不动，中间的对象等距排开（至少选 3 个）。只选一个对象时在幻灯片上上下居中；打开“对齐到对象”后，其他对象在第一个选中的对象的高度范围内等距排开。",
                "Spaces the selected objects evenly from top to bottom: the outermost two stay put and the ones between them move (select at least 3). One object on its own is centred on the slide. With Align to Object on, the other objects are spread evenly across the height of the first object you selected." } },
            // ---- ribbon: rotate / flip
            { "grp.rotate", new[] { "旋转", "Rotate" } },
            { "rotate.left.label", new[] { "向左旋转 90°", "Rotate Left 90°" } },
            { "rotate.left.super", new[] { "把选中的每个对象逆时针旋转 90°，各自绕自己的中心转，位置不变。", "Turns each selected object 90° counter-clockwise about its own centre." } },
            { "rotate.right.label", new[] { "向右旋转 90°", "Rotate Right 90°" } },
            { "rotate.right.super", new[] { "把选中的每个对象顺时针旋转 90°，各自绕自己的中心转，位置不变。", "Turns each selected object 90° clockwise about its own centre." } },
            { "flip.h.label", new[] { "水平翻转", "Flip Horizontal" } },
            { "flip.h.super", new[] { "把选中的每个对象左右镜像翻转，位置不变。", "Mirrors each selected object left to right, in place." } },
            { "flip.v.label", new[] { "垂直翻转", "Flip Vertical" } },
            { "flip.v.super", new[] { "把选中的每个对象上下镜像翻转，位置不变。", "Mirrors each selected object top to bottom, in place." } },

            // ---- ribbon: boolean (merge shapes)
            { "grp.merge", new[] { "布尔", "Merge" } },
            { "merge.union.label", new[] { "联合", "Union" } },
            { "merge.union.super", new[] { "把选中的形状合并成一个形状，包含它们覆盖的全部区域。" + MergeRefZh, "Merges the selected shapes into one shape that covers all of them." + MergeRefEn } },
            { "merge.combine.label", new[] { "组合", "Combine" } },
            { "merge.combine.super", new[] { "把选中的形状合并成一个形状，重叠的部分镂空。" + MergeRefZh, "Merges the selected shapes into one shape with the overlapping parts cut out." + MergeRefEn } },
            { "merge.subtract.label", new[] { "剪除", "Subtract" } },
            { "merge.subtract.super", new[] { "从第一个选中的形状中剪掉其他形状盖住的部分。" + MergeRefZh, "Cuts the other shapes out of the first shape you selected." + MergeRefEn } },
            { "merge.fragment.label", new[] { "拆分", "Fragment" } },
            { "merge.fragment.super", new[] { "沿所有形状的轮廓，把它们切成多个独立的形状。" + MergeRefZh, "Cuts the shapes apart along all their outlines into separate pieces." + MergeRefEn } },

            // ---- ribbon: shape style
            { "grp.style", new[] { "样式", "Style" } },
            { "style.fill.label", new[] { "填充", "Fill" } },
            { "style.fill.tip", new[] { "形状填充颜色", "Shape fill color" } },
            { "style.fill.super", new[] {
                "给选中的形状换填充颜色：点色块立即生效。下面是“无填充”和“其他颜色…”（打开选色窗口，可以输入色值）。线条没有填充，会被跳过；组合里的每个形状都会改。",
                "Changes the fill color of the selected shapes: click a swatch to apply it. Below are No Fill and More Colors… (opens the color picker, where you can type a value). Lines have no fill and are skipped; every shape inside a group is changed." } },
            { "style.fill.none", new[] { "无填充", "No Fill" } },
            { "style.outline.label", new[] { "轮廓", "Outline" } },
            { "style.outline.tip", new[] { "轮廓 / 线条颜色", "Outline / line color" } },
            { "style.outline.super", new[] {
                "给选中的形状换轮廓颜色，对线条和连接线就是线条本身的颜色：点色块立即生效。下面是“无轮廓”和“其他颜色…”。",
                "Changes the outline color of the selected shapes; for lines and connectors that is the color of the line itself. Click a swatch to apply it. Below are No Outline and More Colors…." } },
            { "style.outline.none", new[] { "无轮廓", "No Outline" } },
            { "style.weight.label", new[] { "粗细", "Weight" } },
            { "style.weight.tip", new[] { "轮廓 / 线条粗细", "Outline / line weight" } },
            { "style.weight.super", new[] {
                "设置选中形状的轮廓粗细（线条的粗细），单位是磅。列表里没有想要的数值时，点“其他粗细…”自己输入。原来没有轮廓的形状会显示出轮廓。",
                "Sets the outline weight of the selected shapes (the thickness of lines), in points. Choose More Weights… to type a value that is not in the list. A shape that had no outline gets one." } },
            { "style.weight.more", new[] { "其他粗细…", "More Weights…" } },
            { "style.dash.label", new[] { "虚线", "Dashes" } },
            { "style.dash.tip", new[] { "轮廓 / 线条的虚线类型", "Outline / line dash type" } },
            { "style.dash.super", new[] {
                "把选中形状的轮廓（或线条）改成实线、圆点、方点、短划线等类型。",
                "Makes the outline of the selected shapes (or the line) solid, dotted, dashed and so on." } },
            { "dash.solid", new[] { "实线", "Solid" } },
            { "dash.rounddot", new[] { "圆点", "Round Dot" } },
            { "dash.squaredot", new[] { "方点", "Square Dot" } },
            { "dash.dash", new[] { "短划线", "Dash" } },
            { "dash.dashdot", new[] { "划线-点", "Dash Dot" } },
            { "dash.longdash", new[] { "长划线", "Long Dash" } },
            { "dash.longdashdot", new[] { "长划线-点", "Long Dash Dot" } },
            { "dash.longdashdotdot", new[] { "长划线-点-点", "Long Dash Dot Dot" } },

            { "toobject.label", new[] { "对齐到对象", "Align to Object" } },
            { "toobject.super", new[] {
                "打开后，对齐和分布以第一个选中的对象为参照，它本身不动。先单击作为参照的对象，再按住 Shift 或 Ctrl 单击其他对象，然后点对齐或分布按钮。分布时，其他对象在参照对象的范围内等距排开，适合把几个对象均匀地放进一个框里。",
                "When on, aligning and distributing use the first object you select as the reference, and that object does not move. Click the reference object first, Shift- or Ctrl-click the others, then click an align or distribute button. Distributing spreads the other objects evenly inside the reference object, which is handy for placing items evenly in a box." } },

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
            { "title.align", new[] { "对齐", "Align" } },
            { "msg.align.nowindow", new[] {
                "请先打开一个演示文稿，并在普通编辑视图中选中要对齐的对象。",
                "Open a presentation and select the objects to align in the normal editing view." } },
            { "msg.align.noselection", new[] {
                "请先选中要对齐的对象。\n可以按住 Shift / Ctrl 逐个点选，或用鼠标拖出选框一次选中多个。",
                "Select the objects to align first.\nShift- or Ctrl-click to pick several, or drag a selection box around them." } },
            { "msg.align.needthree", new[] {
                "分布至少需要选中 3 个对象。\n（打开“对齐到对象”后，也可以让其他对象在第一个选中的对象范围内等距排开。）",
                "Distributing needs at least 3 selected objects.\n(With Align to Object on, you can also spread the other objects evenly inside the first object you select.)" } },
            { "msg.align.unmoved", new[] {
                "有 {0} 个对象无法移动（可能被锁定了），其他对象已经对齐。",
                "{0} of the objects could not be moved (they may be locked). The others were aligned." } },
            { "msg.align.failed", new[] { "对齐失败：\n{0}", "Could not align:\n{0}" } },
            { "title.rotate", new[] { "旋转", "Rotate" } },
            { "title.merge", new[] { "布尔运算", "Merge Shapes" } },
            { "msg.shape.nowindow", new[] {
                "请先打开一个演示文稿，并在普通编辑视图中选中对象。",
                "Open a presentation and select objects in the normal editing view." } },
            { "msg.shape.noselection", new[] {
                "请先在幻灯片上选中对象。\n可以按住 Shift / Ctrl 逐个点选，或用鼠标拖出选框一次选中多个。",
                "Select objects on the slide first.\nShift- or Ctrl-click to pick several, or drag a selection box around them." } },
            { "msg.shape.failed", new[] { "操作失败：\n{0}", "That did not work:\n{0}" } },
            { "msg.merge.needtwo", new[] {
                "布尔运算至少需要选中 2 个形状。\n先单击要保留格式的形状，再按住 Shift 或 Ctrl 单击其他形状。",
                "Merging needs at least 2 selected shapes.\nClick the shape whose format you want to keep first, then Shift- or Ctrl-click the others." } },
            { "msg.merge.grouped", new[] {
                "组合不能直接做布尔运算。请先取消组合（Ctrl+Shift+G），再选中要运算的形状。",
                "Groups cannot be merged. Ungroup them first (Ctrl+Shift+G), then select the shapes to merge." } },
            { "msg.merge.failed", new[] {
                "无法对这些对象做布尔运算：\n{0}\n\n表格、图表、SmartArt 等对象不能参与布尔运算。",
                "These objects cannot be merged:\n{0}\n\nTables, charts, SmartArt and similar objects cannot take part in merging." } },

            { "title.style", new[] { "样式", "Style" } },
            { "msg.style.none", new[] {
                "选中的对象里没有可以这样设置的形状。\n（线条没有填充；表格、图表等对象请用 PowerPoint 自带的格式工具。）",
                "None of the selected objects can take this setting.\n(Lines have no fill; for tables, charts and similar objects use PowerPoint's own format tools.)" } },

            // ---- weight dialog
            { "weight.title", new[] { "线条粗细", "Line weight" } },
            { "weight.caption", new[] { "粗细（磅）", "Weight (points)" } },
            { "weight.hint", new[] { "请输入 0.05 到 100 之间的数值，例如 1.25", "Enter a value from 0.05 to 100, for example 1.25" } },

            // ---- colour picker
            { "picker.fill", new[] { "填充颜色", "Fill color" } },
            { "picker.outline", new[] { "轮廓颜色", "Outline color" } },
            { "picker.title", new[] { "背景颜色", "Background color" } },
            { "picker.new", new[] { "新颜色", "New" } },
            { "picker.current", new[] { "当前", "Current" } },
            { "picker.hex", new[] { "十六进制", "Hex" } },
            { "picker.palette", new[] { "常用颜色", "Colors" } },
            { "picker.recent", new[] { "最近使用", "Recent" } },
            { "picker.recent.empty", new[] { "用过的颜色会出现在这里", "Colors you choose will appear here" } },
            { "picker.ok", new[] { "确定", "OK" } },
            { "picker.cancel", new[] { "取消", "Cancel" } },
            { "picker.close", new[] { "关闭", "Close" } },

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
