#nullable enable
using System.Text;
using System.Xml;
using System.Xml.Linq;
using AnnotationModal;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
namespace 附注工具.核心;

public sealed class StylePlan
{
    private readonly XDocument xml;
    private readonly List<StyleSetting> rules;
    private readonly string sourcePath;
    private string sourceHash;

    private StylePlan(string path, byte[] bytes)
    {
        sourcePath = path;
        sourceHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes));
        using var input = new MemoryStream(bytes, false);
        xml = XDocument.Load(input, LoadOptions.PreserveWhitespace);
        rules = ReadStyleXml(xml);
    }

    public static StylePlan Load(string path)
    {
        path = Path.GetFullPath(path);
        try { return new StylePlan(path, File.ReadAllBytes(path)); }
        catch (Exception error) when (error is XmlException or FormatException or NullReferenceException)
        { throw new InvalidDataException("格式配置缺少有效的规则或属性。", error); }
    }

    public SavedFile Save(string path)
    {
        path = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var staged = Path.Combine(Path.GetDirectoryName(path)!, Path.GetFileNameWithoutExtension(path) + ".待校验_" + Guid.NewGuid().ToString("N") + ".xml");
        using (var file = new FileStream(staged, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        using (var writer = XmlWriter.Create(file, new XmlWriterSettings { Encoding = new UTF8Encoding(true), Indent = false, NewLineHandling = NewLineHandling.None }))
            xml.Save(writer);
        if (!XNode.DeepEquals(xml, XDocument.Load(staged, LoadOptions.PreserveWhitespace)))
            throw new InvalidDataException("格式配置保存后未完整保留原节点、选项与顺序。");
        var hash = SafeFile.Hash(staged);
        var backup = SafeFile.BackupExisting(path);
        File.Move(staged, path, true);
        if (SafeFile.Hash(path) != hash) throw new InvalidDataException("格式配置与已校验副本不一致。");
        if (string.Equals(path, sourcePath, StringComparison.OrdinalIgnoreCase)) sourceHash = hash;
        return new SavedFile(path, backup);
    }

    // 原frm_GetTableFromExcel.cs 203—219：保留启用判断、属性和文件顺序；完整XML另行保存。
    private static List<StyleSetting> ReadStyleXml(XDocument document)
    {
        List<StyleSetting> list = new List<StyleSetting>();
        List<XElement> xElementList = document.Root!.Element("StyleSettings")!.Elements("StyleSetting").ToList();
        foreach (XElement item in xElementList)
        {
            StyleSetting styleSetting = new StyleSetting();
            if ((bool)item.Attribute("used")!)
            {
                styleSetting.Name = item.Attribute("Attr")!.Value;
                styleSetting.Value = item.Attribute("value")!.Value;
                styleSetting.Condition = ((item.Attribute("condition") == null) ? "" : item.Attribute("condition")!.Value);
                list.Add(styleSetting);
            }
        }
        return list;
    }
    public SavedFile Apply(string report, IReadOnlyCollection<string>? selectedTableIds, string output,
        bool applyPageSetup = false, CancellationToken token = default)
    {
        token.ThrowIfCancellationRequested();
        var snapshot = ReportReader.Read(report);
        if (selectedTableIds is not null && selectedTableIds.Any(id => !snapshot.Tables.Any(t => t.Id == id)))
            throw new InvalidDataException("选中的表格位置已失效，请重新选择。");
        var chosen = snapshot.Tables.Where(t => selectedTableIds?.Contains(t.Id) ?? true).ToArray();
        var expected = snapshot.Tables.ToDictionary(t => t.Id, t => t.OriginalXml);
        string[]? expectedSections = null;
        return SafeFile.SaveWordCopy(snapshot.FilePath, snapshot.SourceHash, output, document =>
        {
            var body = document.MainDocumentPart!.Document!.Body!;
            if (applyPageSetup) SetDocumentFormat(document);
            foreach (var info in chosen)
            {
                token.ThrowIfCancellationRequested();
                var table = (W.Table)TableSynchronizer.FindNode(body, info.NodePath);
                foreach (var setting in rules)
                {
                    try { SetTableFormat(table, info, setting); }
                    catch (Exception error) when (error is FormatException or OverflowException or ArgumentException)
                    { throw new InvalidDataException($"表格“{info.Title}”的格式“{setting.Name}”数值无效。", error); }
                }
                expected[info.Id] = table.OuterXml;
            }
            expectedSections = body.Descendants<W.SectionProperties>().Select(s => s.OuterXml).ToArray();
        }, file =>
        {
            var reread = ReportReader.Read(file);
            if (reread.Tables.Count != snapshot.Tables.Count || reread.Tables.Any(t => !expected.TryGetValue(t.Id, out var original) || !SameXml(LinkManager.RebaseTableXml(original,report,output), t.OriginalXml)))
                throw new InvalidDataException("格式保存后原表内容或位置核对不一致。");
            using var check = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, false);
            var sections = check.MainDocumentPart!.Document!.Body!.Descendants<W.SectionProperties>().Select(s => s.OuterXml).ToArray();
            if (sections.Length != expectedSections!.Length || sections.Where((s, i) => !SameXml(s, expectedSections[i])).Any())
                throw new InvalidDataException("格式保存后页面设置核对不一致。");
        }, token, new Dictionary<string, string> { [sourcePath] = sourceHash });
    }

    // 原222—265行的动作顺序与分支不变；只有FontUnderline/Alignment使用条件。
    private static void SetTableFormat(W.Table table, TableInfo info, StyleSetting set)
    {
        switch (set.Name)
        {
            case "AutoFitBehavior":
                var width = Get<W.TableWidth>(Get<W.TableProperties>(table));
                width.Type = W.TableWidthUnitValues.Pct;
                width.Width = "5000";
                Get<W.TableLayout>(Get<W.TableProperties>(table)).Type = W.TableLayoutValues.Autofit;
                break;
            case "BackgroundPatternColor":
                SetBackground(Get<W.TableProperties>(table));
                foreach (var cell in Cells(table)) SetBackground(Get<W.TableCellProperties>(cell));
                break;
            case "InnerBorderLineStyle":
            case "TopBottom":
            case "LeftRight":
                SetBorders(table, set);
                break;
            case "FontSpacing":
                var spacing = Units(Convert.ToSingle(set.Value), 20);
                foreach (var properties in FontProperties(table)) Get<W.Spacing>(properties).Val = spacing;
                break;
            case "CharacterUnitLeftRightIndent":
                var indent = Units(Convert.ToSingle(set.Value), 100);
                foreach (var paragraph in table.Descendants<W.Paragraph>())
                {
                    var ind = Get<W.Indentation>(Get<W.ParagraphProperties>(paragraph));
                    ind.LeftChars = indent;
                    ind.RightChars = indent;
                }
                break;
            case "FontSize":
                var size = Units(Convert.ToSingle(set.Value), 2);
                if (size <= 0) throw new FormatException("字号必须大于零。");
                foreach (var properties in FontProperties(table)) Get<W.FontSize>(properties).Val = size.ToString(System.Globalization.CultureInfo.InvariantCulture);
                break;
            case "FontName":
                foreach (var properties in FontProperties(table))
                {
                    var font = Get<W.RunFonts>(properties);
                    font.Ascii = set.Value;
                    font.HighAnsi = set.Value;
                    font.EastAsia = set.Value;
                    font.ComplexScript = set.Value;
                    font.AsciiTheme = null; font.HighAnsiTheme = null; font.EastAsiaTheme = null; font.ComplexScriptTheme = null;
                }
                break;
            case "Height":
                var height = Units(Convert.ToSingle(set.Value), 1440d / 2.54d);
                if (height < 0) throw new FormatException("最小行高不能为负。");
                foreach (var row in table.Elements<W.TableRow>())
                {
                    var rowHeight = Get<W.TableRowHeight>(Get<W.TableRowProperties>(row));
                    rowHeight.HeightType = W.HeightRuleValues.AtLeast;
                    rowHeight.Val = (uint)height;
                }
                break;
            case "FontUnderline":
                foreach (var target in ConditionalTargets(table, info, set.Condition))
                    foreach (var properties in FontProperties(target))
                    {
                        var underline = Get<W.Underline>(properties);
                        underline.Val = W.UnderlineValues.Single;
                        underline.Color = "auto";
                        underline.ThemeColor = null; underline.ThemeShade = null; underline.ThemeTint = null;
                    }
                break;
            case "Alignment":
                var alignment = Convert.ToInt32(set.Value) switch
                {
                    0 => "left", 1 => "center", 2 => "right", 3 => "both", 4 => "distribute",
                    5 => "mediumKashida", 7 => "highKashida", 8 => "lowKashida", 9 => "thaiDistribute",
                    _ => throw new FormatException("对齐值不是原Word支持的枚举。")
                };
                foreach (var target in ConditionalTargets(table, info, set.Condition))
                    foreach (var paragraph in target.Descendants<W.Paragraph>())
                        Get<W.Justification>(Get<W.ParagraphProperties>(paragraph)).Val = new W.JustificationValues(alignment);
                break;
        }
    }

    private static void SetBackground(OpenXmlCompositeElement properties)
    {
        var shading = Get<W.Shading>(properties);
        shading.Val ??= W.ShadingPatternValues.Clear;
        shading.Fill = "auto"; shading.Color = "auto";
        shading.ThemeFill = null; shading.ThemeFillShade = null; shading.ThemeFillTint = null;
        shading.ThemeColor = null; shading.ThemeShade = null; shading.ThemeTint = null;
    }

    private static void SetBorders(W.Table table, StyleSetting set)
    {
        if (!int.TryParse(set.Value, out var value)) return;
        string style;
        uint size = 4;
        string[] sides;
        if (set.Name == "InnerBorderLineStyle")
        {
            style = value == 1 ? "dotted" : "single";
            sides = ["insideH", "insideV"];
        }
        else if (set.Name == "TopBottom")
        {
            style = value == 3 ? "dotted" : "single";
            size = value == 1 ? 8u : 4u;
            sides = ["top", "bottom"];
        }
        else
        {
            style = value switch { 2 => "single", 3 => "dotted", _ => "nil" };
            sides = ["left", "right"];
        }
        var borders = Get<W.TableBorders>(Get<W.TableProperties>(table));
        foreach (var side in sides) SetBorder(borders, side, style, size);
        // Word表级边框会覆盖同一条边的单元格设置；文件版也必须处理已有tcBorders。
        var rows = table.Elements<W.TableRow>().ToArray();
        for (var r = 0; r < rows.Length; r++)
        {
            var cells = rows[r].Elements<W.TableCell>().ToArray();
            for (var c = 0; c < cells.Length; c++)
            {
                var cellBorders = Get<W.TableCellBorders>(Get<W.TableCellProperties>(cells[c]));
                if (set.Name == "TopBottom")
                {
                    if (r == 0) SetBorder(cellBorders, "top", style, size);
                    if (r == rows.Length - 1) SetBorder(cellBorders, "bottom", style, size);
                }
                else if (set.Name == "LeftRight")
                {
                    if (c == 0) SetBorder(cellBorders, "left", style, size);
                    if (c == cells.Length - 1) SetBorder(cellBorders, "right", style, size);
                }
                else
                {
                    if (r > 0) SetBorder(cellBorders, "top", style, size);
                    if (r < rows.Length - 1) SetBorder(cellBorders, "bottom", style, size);
                    if (c > 0) SetBorder(cellBorders, "left", style, size);
                    if (c < cells.Length - 1) SetBorder(cellBorders, "right", style, size);
                }
            }
        }
        if (set.Name == "LeftRight")
            foreach (var border in table.Descendants<W.BorderType>().Where(b =>
                (b.Parent is W.TableBorders or W.TableCellBorders) && ReferenceEquals(b.Ancestors<W.Table>().FirstOrDefault(), table)))
                border.Shadow = false;
    }

    private static void SetBorder(OpenXmlCompositeElement parent, string side, string style, uint size)
    {
        W.BorderType border = side switch
        {
            "top" => Get<W.TopBorder>(parent), "bottom" => Get<W.BottomBorder>(parent),
            "left" => Get<W.LeftBorder>(parent), "right" => Get<W.RightBorder>(parent),
            "insideH" => Get<W.InsideHorizontalBorder>(parent), "insideV" => Get<W.InsideVerticalBorder>(parent),
            _ => throw new ArgumentOutOfRangeException(nameof(side))
        };
        border.Val = new W.BorderValues(style);
        border.Size = style == "nil" ? null : size;
        border.Color = "auto";
        border.ThemeColor = null; border.ThemeShade = null; border.ThemeTint = null;
    }

    private static IEnumerable<W.TableCell> Cells(W.Table table) => table.Elements<W.TableRow>().SelectMany(r => r.Elements<W.TableCell>());
    private static IEnumerable<OpenXmlCompositeElement> FontProperties(OpenXmlElement element)
    {
        foreach (var run in element.Descendants<W.Run>()) yield return Get<W.RunProperties>(run);
        foreach (var paragraph in element.Descendants<W.Paragraph>()) yield return Get<W.ParagraphMarkRunProperties>(Get<W.ParagraphProperties>(paragraph));
    }
    private static T Get<T>(OpenXmlCompositeElement parent) where T : OpenXmlElement, new()
    {
        var child = parent.GetFirstChild<T>();
        if (child is null) parent.AddChild(child = new T(), true);
        return child;
    }
    private static int Units(float value, double factor)
    {
        if (!float.IsFinite(value)) throw new FormatException("格式数值不能为无穷或非数值。");
        return checked((int)Math.Round(value * factor, MidpointRounding.AwayFromZero));
    }

    private static IEnumerable<OpenXmlElement> ConditionalTargets(W.Table table, TableInfo info, string condition)
    {
        if (string.IsNullOrEmpty(condition)) { yield return table; yield break; }
        var spec = ParseCondition(condition);
        var layout = WordTableLayout.Read(info);
        var row = 0;
        if (spec.Row.HasValue)
        {
            row = ResolveRowIndex(spec.Row.Value, layout.WordRowCount);
            if (row < 1 || row > layout.WordRowCount) yield break;
        }
        var rows = table.Elements<W.TableRow>().ToArray();
        foreach (var position in layout.Cells)
        {
            var cell = rows[position.XmlRowIndex].Elements<W.TableCell>().ElementAt(position.XmlCellIndex);
            if ((row <= 0 || position.WordRowIndex == row) && MatchesTypeCondition(spec.Type, cell.InnerText)) yield return cell;
        }
    }

    // 原490—565行的条件、倒数行和文字清理语句；仅去掉Word Cell参数。
    private sealed class ConditionSpec
    {
        public int? Row { get; set; }
        public int? Type { get; set; }
    }
    private static ConditionSpec ParseCondition(string condition)
    {
        ConditionSpec conditionSpec = new ConditionSpec();
        if (string.IsNullOrWhiteSpace(condition)) return conditionSpec;
        string[] array = condition.Split('|');
        string[] array2 = array;
        foreach (string text in array2)
        {
            if (string.IsNullOrWhiteSpace(text)) continue;
            string text2 = text.Trim();
            int result2;
            if (text2.Length >= 4 && text2.StartsWith("row=", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(text2.Substring(4), out var result)) conditionSpec.Row = result;
            }
            else if (text2.Length >= 5 && text2.StartsWith("type=", StringComparison.OrdinalIgnoreCase) && int.TryParse(text2.Substring(5), out result2))
                conditionSpec.Type = result2;
        }
        return conditionSpec;
    }
    private static int ResolveRowIndex(int rowConditionValue, int rowCount)
    {
        if (rowConditionValue > 0) return rowConditionValue;
        if (rowConditionValue < 0) return rowCount + rowConditionValue + 1;
        return 0;
    }
    private static bool MatchesTypeCondition(int? type, string text)
    {
        if (!type.HasValue) return true;
        string s = NormalizeCellText(text);
        decimal result;
        return type.Value switch { 0 => decimal.TryParse(s, out result), 1 => !decimal.TryParse(s, out result), _ => true };
    }
    private static string NormalizeCellText(string text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        StringBuilder stringBuilder = new StringBuilder(text.Length);
        foreach (char c in text)
            if (c != '\a' && !char.IsWhiteSpace(c)) stringBuilder.Append(c);
        return stringBuilder.ToString();
    }

    // 原613—630行；仅由明确的全文排版选项调用。
    private static void SetDocumentFormat(DocumentFormat.OpenXml.Packaging.WordprocessingDocument document)
    {
        var body = document.MainDocumentPart!.Document!.Body!;
        var sections = body.Descendants<W.SectionProperties>().ToArray();
        if (sections.Length == 0) sections = [Get<W.SectionProperties>(body)];
        foreach (var section in sections)
        {
            var size = Get<W.PageSize>(section);
            size.Width = 11906; size.Height = 16838; size.Orient = W.PageOrientationValues.Portrait;
            var margin = Get<W.PageMargin>(section);
            margin.Top = 1440; margin.Bottom = 1440; margin.Left = 1077; margin.Right = 1077;
            margin.Gutter = 0; margin.Header = 850; margin.Footer = 992;
            Get<W.VerticalTextAlignmentOnPage>(section).Val = W.VerticalJustificationValues.Top;
            Get<W.NoEndnote>(section).Val = false;
            Get<W.DocGrid>(section).Type = W.DocGridValues.Lines;
        }
        var part = document.MainDocumentPart.DocumentSettingsPart;
        if (part is not null)
        {
            part.Settings ??= new W.Settings();
            Get<W.MirrorMargins>(part.Settings).Val = false;
        }
    }
    private static bool SameXml(string first, string second) => TableSynchronizer.SameXml(first,second);

}
