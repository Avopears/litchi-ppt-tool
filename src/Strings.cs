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

            // ---- ribbon: stacking order
            { "grp.order", new[] { "层次", "Order" } },
            { "order.front.label", new[] { "置于顶层", "Bring to Front" } },
            { "order.front.super", new[] { "把选中的对象放到所有对象的最上面。选中多个对象时，它们之间的上下顺序不变。", "Puts the selected objects on top of everything else. Several selected objects keep their order among themselves." } },
            { "order.forward.label", new[] { "上移一层", "Bring Forward" } },
            { "order.forward.super", new[] { "把选中的对象往上移一层。", "Moves the selected objects one step towards the front." } },
            { "order.back.label", new[] { "置于底层", "Send to Back" } },
            { "order.back.super", new[] { "把选中的对象放到所有对象的最下面。选中多个对象时，它们之间的上下顺序不变。", "Puts the selected objects behind everything else. Several selected objects keep their order among themselves." } },
            { "order.backward.label", new[] { "下移一层", "Send Backward" } },
            { "order.backward.super", new[] { "把选中的对象往下移一层。", "Moves the selected objects one step towards the back." } },

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
            { "style.cap.label", new[] { "线端", "Cap" } },
            { "style.cap.tip", new[] { "线端类型", "Line cap type" } },
            { "style.cap.super", new[] {
                "线条两端（以及虚线每一小段两端）的形状，就是“设置形状格式 → 线条 → 线端类型”那一项。\n平：正好在端点处截止；方：两端各多出半个线宽；圆：两端是半圆。\n\n只能用于普通幻灯片上的对象。插件要先临时存一份演示文稿的副本才能改这一项，文件很大时请稍等。",
                "The shape of a line's two ends (and of the ends of every dash): the Cap type setting under Format Shape > Line.\nFlat: stops exactly at the end point. Square: reaches half the line weight past it. Round: ends in a half circle.\n\nFor objects on a slide only. The add-in has to save a temporary copy of the presentation to change this, so very large files take a moment." } },
            { "cap.square", new[] { "方", "Square" } },
            { "cap.round", new[] { "圆", "Round" } },
            { "cap.flat", new[] { "平", "Flat" } },

            // ---- ribbon: slide size
            { "grp.slide", new[] { "幻灯片", "Slide" } },
            { "size.label", new[] { "设置尺寸", "Slide Size" } },
            { "size.tip", new[] { "设置幻灯片尺寸（厘米）", "Set the slide size (in centimetres)" } },
            { "size.super", new[] {
                "按厘米设置幻灯片的宽度和高度。已经做好的内容会按“确保适合”的方式自动等比缩放并保持居中：按宽、高两个比例里较小的那个缩放，原来在页面里的内容不会跑到页面外。\n\n位置、大小、字号、文字边距和段落间距一起缩放；线条粗细和阴影等效果保持不变（和 PowerPoint 一样）。一次 Ctrl+Z 可以撤回。",
                "Sets the width and height of the slides in centimetres. What is already on them is scaled to fit and stays centred, as with PowerPoint's Ensure Fit: by the smaller of the two ratios, so that nothing that was on the slide ends up off it.\n\nPosition, size, font sizes, text margins and paragraph spacing are scaled together; line weights and effects such as shadows keep their size (as in PowerPoint). One Ctrl+Z takes it back." } },

            // ---- ribbon: file tools
            { "grp.file", new[] { "文件", "File" } },
            { "slim.label", new[] { "PPT 瘦身", "Slim Down" } },
            { "slim.tip", new[] { "压缩图片，另存一份更小的演示文稿", "Compress the pictures into a smaller copy of the presentation" } },
            { "slim.super", new[] {
                "把演示文稿里的图片缩小到够用的分辨率并重新压缩，另存成一份更小的新文件，原文件不会被改动。\n\n可以选分辨率上限和图片质量，先看能省多少，满意了再保存。矢量图、动图和音视频不处理。",
                "Shrinks the pictures in the presentation to the resolution they need, compresses them again and saves the result as a new, smaller file. The original is not changed.\n\nChoose a resolution limit and a picture quality, see how much it saves, and save when you like the result. Vector images, animated GIFs, video and audio are left as they are." } },

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
            { "title.order", new[] { "层次", "Order" } },
            { "msg.order.child", new[] {
                "无法调整这个组合内对象的层次。\n可以先取消组合（Ctrl+Shift+G），调整后再重新组合。",
                "The order of this object inside its group could not be changed.\nUngroup first (Ctrl+Shift+G), change the order, then group again." } },
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
            { "msg.cap.needsslide", new[] {
                "线端类型只能设置给普通幻灯片上的对象。\n（母版、版式、备注页里的对象，请用“设置形状格式 → 线条 → 线端类型”。）",
                "The cap type can only be set for objects on a slide.\n(For objects on a master, a layout or a notes page use Format Shape > Line > Cap type.)" } },
            { "err.cap.copy", new[] {
                "读不了演示文稿的临时副本（设置了打开密码的演示文稿不支持这一项）。",
                "The temporary copy of the presentation could not be read (presentations with a password to open are not supported here)." } },

            // ---- slide size
            { "size.title", new[] { "幻灯片尺寸", "Slide size" } },
            { "size.current", new[] { "当前：{0} × {1} 厘米", "Now: {0} × {1} cm" } },
            { "size.width", new[] { "宽度（厘米）", "Width (cm)" } },
            { "size.height", new[] { "高度（厘米）", "Height (cm)" } },
            { "size.swap", new[] { "交换", "Swap" } },
            { "size.presets", new[] { "常用尺寸", "Common sizes" } },
            { "size.preset.wide", new[] { "宽屏 16:9", "Widescreen 16:9" } },
            { "size.preset.standard", new[] { "标准 4:3", "Standard 4:3" } },
            { "size.preset.a4", new[] { "A4 横向", "A4 landscape" } },
            { "size.lock", new[] { "锁定宽高比（改一个，另一个跟着变）", "Lock the proportions (change one, the other follows)" } },
            { "size.hint.scale", new[] { "已有内容会等比缩放到 {0}%，并保持居中。", "What is on the slides is scaled to {0}% and stays centred." } },
            { "size.hint.same", new[] { "已有内容大小不变，保持居中。", "What is on the slides keeps its size and stays centred." } },
            { "size.hint.nochange", new[] { "尺寸没有变化。", "The size is unchanged." } },
            { "size.hint.invalid", new[] { "请输入 2.54 到 142.24 之间的数值（厘米）。", "Enter values from 2.54 to 142.24 (centimetres)." } },
            { "size.note", new[] { "线条粗细和阴影等效果保持不变，和 PowerPoint 的“确保适合”一样。", "Line weights and effects keep their size, as with PowerPoint's Ensure Fit." } },
            { "msg.size.failed", new[] { "设置尺寸失败：\n{0}", "The slide size could not be changed:\n{0}" } },

            // ---- slim down
            { "slim.title", new[] { "PPT 瘦身", "Slim down presentation" } },
            { "slim.overview", new[] { "当前大小：{0}", "Size now: {0}" } },
            { "slim.part.images", new[] { "图片 {0} 张，{1}", "{0} pictures, {1}" } },
            { "slim.part.media", new[] { "音视频 {0}", "Video and audio {0}" } },
            { "slim.part.other", new[] { "其他 {0}", "Everything else {0}" } },
            { "slim.none", new[] { "无", "none" } },
            { "slim.ppi.caption", new[] { "图片分辨率上限（按图片在幻灯片上显示的大小计算）", "Resolution limit (for the size a picture has on its slide)" } },
            { "slim.ppi.keep", new[] { "不缩小", "Keep" } },
            { "slim.ppi.hint.0", new[] { "不减少像素，只重新压缩。", "Keeps every pixel and only compresses again." } },
            { "slim.ppi.hint.330", new[] { "高清：放大看、高分辨率打印都清晰。", "High definition: sharp when zoomed in or printed at high resolution." } },
            { "slim.ppi.hint.220", new[] { "推荐：打印和 4K 屏幕全屏放映都清晰。", "Recommended: sharp in print and full-screen on a 4K display." } },
            { "slim.ppi.hint.150", new[] { "投影仪和普通屏幕放映足够，文件更小。", "Enough for projectors and ordinary screens; smaller files." } },
            { "slim.ppi.hint.96", new[] { "最小，适合发邮件；放大看会模糊。", "Smallest, good for e-mail; blurry when zoomed in." } },
            { "slim.quality.caption", new[] { "图片质量（JPEG）：85 以上几乎看不出差别，60 以下会出现色块", "Picture quality (JPEG): hard to tell apart above 85, blocky below 60" } },
            { "slim.tojpeg", new[] { "把不带透明的照片从 PNG 转成 JPG（截图和图表仍保持 PNG）", "Save photos without transparency as JPG instead of PNG (charts stay PNG)" } },
            { "slim.idle", new[] {
                "点“开始压缩”看看能省多少，满意了再“另存为…”。原文件不会被改动。",
                "Click Compress to see how much it saves, then Save As… when you like the result. The original file is not changed." } },
            { "slim.changed", new[] { "设置改了，再点一次“开始压缩”看新的结果。", "The settings have changed: click Compress again for the new result." } },
            { "slim.preparing", new[] { "正在分析…", "Analysing…" } },
            { "slim.working", new[] { "正在处理图片 {0} / {1}…", "Compressing picture {0} of {1}…" } },
            { "slim.result.detail", new[] { "改动了 {0} 张图片，{1} 张保持原样。", "{0} pictures changed, {1} left as they are." } },
            { "slim.result.none", new[] { "按当前设置，这份演示文稿已经没什么可压的了。", "With these settings there is next to nothing to gain in this presentation." } },
            { "slim.start", new[] { "开始压缩", "Compress" } },
            { "slim.cancel", new[] { "取消", "Cancel" } },
            { "slim.saveas", new[] { "另存为…", "Save As…" } },
            { "slim.saved", new[] { "已保存：{0}", "Saved: {0}" } },
            { "slim.dlg", new[] { "另存瘦身后的演示文稿", "Save the slimmed presentation" } },
            { "filter.pptx", new[] { "PowerPoint 演示文稿 (*.pptx)|*.pptx", "PowerPoint presentation (*.pptx)|*.pptx" } },
            { "filter.pptm", new[] { "启用宏的 PowerPoint 演示文稿 (*.pptm)|*.pptm", "PowerPoint macro-enabled presentation (*.pptm)|*.pptm" } },
            { "slim.suffix", new[] { "_压缩", "_compressed" } },
            { "msg.slim.nowindow", new[] { "请先打开一个演示文稿。\n（受保护的视图下无法使用。）", "Open a presentation first.\n(Not available in Protected View.)" } },
            { "msg.slim.samefile", new[] { "不能覆盖正在打开的这份演示文稿，请换一个文件名。", "The presentation that is open cannot be overwritten. Choose another file name." } },
            { "msg.slim.failed", new[] { "瘦身失败：\n{0}", "Slimming failed:\n{0}" } },
            { "msg.slim.locked", new[] {
                "\n\n如果目标文件正在 PowerPoint 或其他程序里打开，请先关闭它。",
                "\n\nIf the target file is open in PowerPoint or another program, close it first." } },

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
