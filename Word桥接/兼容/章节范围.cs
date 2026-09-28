#nullable enable
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace 附注工具.核心;

// 章节范围只选择需要更新的真实表格，不改变文档的章节和表格编号。
internal sealed record ChapterScope(string Title, string StartHeading, string? EndHeading,
    int StartIndex, int EndIndex, string[] TableIds)
{
    private static readonly XNamespace Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly Regex NumberPrefix = new(@"^[一二三四五六七八九十百零〇]+[、．.]", RegexOptions.CultureInvariant);
    public object Info => new { chapter_title = Title, start_heading = StartHeading, end_heading = EndHeading,
        start_index = StartIndex, end_index = EndIndex, table_count = TableIds.Length, table_ids = TableIds };
    public bool Contains(TableInfo table) => TableIds.Contains(table.Id, StringComparer.Ordinal);
    public bool Contains(OpenXmlElement element)
    {
        var top = element;
        while (top.Parent is not null && top.Parent is not W.Body) top = top.Parent;
        if (top.Parent is not W.Body body) return false;
        var index = body.ChildElements.ToList().IndexOf(top);
        return index >= StartIndex && index < EndIndex;
    }
    private static string Text(XElement node) => string.Concat(node.Descendants(Wns + "t").Select(n => n.Value)).Trim();
    private static string Normalize(string text) => NumberPrefix.Replace(Regex.Replace(text, @"\s+", ""), "");
    public static ChapterScope? Resolve(JObject request, ReportInfo report)
    {
        if (request["chapter_title"] is null) return null;
        if (request["chapter_title"]?.Type != JTokenType.String || string.IsNullOrWhiteSpace((string?)request["chapter_title"]))
            throw new InvalidDataException("chapter_title 必须是非空章节标题。");
        var title = Normalize((string)request["chapter_title"]!);
        using var archive = ZipFile.OpenRead(report.FilePath);
        using var stream = archive.GetEntry("word/document.xml")!.Open();
        var document = XDocument.Load(stream);
        var body = document.Root!.Element(Wns + "body")!;
        var nodes = body.Elements().ToArray();
        var matches = nodes.Select((node, index) => (node, index))
            .Where(p => p.node.Name == Wns + "p" && Normalize(Text(p.node)) == title).ToArray();
        if (matches.Length == 0) throw new InvalidDataException("未找到指定的大章节：" + title + "。不会改为处理全文。");
        if (matches.Length != 1) throw new InvalidDataException("找到多个同名章节：" + title + "，无法确定处理范围。");
        var styles = new Dictionary<string, XElement>(StringComparer.Ordinal);
        if (archive.GetEntry("word/styles.xml") is { } styleEntry)
        {
            using var styleStream = styleEntry.Open();
            foreach (var style in XDocument.Load(styleStream).Root!.Elements(Wns + "style"))
                if ((string?)style.Attribute(Wns + "styleId") is { } id) styles[id] = style;
        }
        XElement? Property(XElement paragraph, string name)
        {
            var direct = paragraph.Element(Wns + "pPr")?.Element(Wns + name);
            if (direct is not null) return direct;
            var id = (string?)paragraph.Element(Wns + "pPr")?.Element(Wns + "pStyle")?.Attribute(Wns + "val");
            var visited = new HashSet<string>(StringComparer.Ordinal);
            while (id is not null && visited.Add(id) && styles.TryGetValue(id, out var style))
            {
                if (style.Element(Wns + "pPr")?.Element(Wns + name) is { } value) return value;
                id = (string?)style.Element(Wns + "basedOn")?.Attribute(Wns + "val");
            }
            return null;
        }
        int? Outline(XElement p) => (int?)Property(p, "outlineLvl")?.Attribute(Wns + "val");
        var target = matches[0];
        var level = Outline(target.node);
        var number = Property(target.node, "numPr");
        var numberId = (string?)number?.Element(Wns + "numId")?.Attribute(Wns + "val");
        var numberLevel = (int?)number?.Element(Wns + "ilvl")?.Attribute(Wns + "val") ?? 0;
        var manualNumber = NumberPrefix.IsMatch(Regex.Replace(Text(target.node), @"\s+", ""));
        if ((level is null || level >= 9) && numberId is null && !manualNumber)
            throw new InvalidDataException("目标标题未注明章节层级，无法可靠确定结束位置。");
        int end = nodes.Length;
        string? endTitle = null;
        for (int i = target.index + 1; i < nodes.Length; i++)
        {
            if (nodes[i].Name != Wns + "p") continue;
            var nextLevel = Outline(nodes[i]);
            bool ends;
            if (level is < 9) ends = nextLevel is not null && nextLevel <= level;
            else if (numberId is not null)
            {
                var nextNumber = Property(nodes[i], "numPr");
                ends = (string?)nextNumber?.Element(Wns + "numId")?.Attribute(Wns + "val") == numberId
                    && ((int?)nextNumber?.Element(Wns + "ilvl")?.Attribute(Wns + "val") ?? 0) <= numberLevel;
            }
            else ends = NumberPrefix.IsMatch(Regex.Replace(Text(nodes[i]), @"\s+", ""));
            if (!ends) continue;
            end = i; endTitle = Text(nodes[i]); break;
        }
        var ids = report.Tables.Where(t => t.NodePath[0] > target.index && t.NodePath[0] < end).Select(t => t.Id).ToArray();
        return new ChapterScope(title, Text(target.node), endTitle, target.index, end, ids);
    }

    // 上游保存会重定位全篇连接。限定章节时以原包为底稿，仅移植选中的表，避免触碰其他章节。
    public SavedFile SaveSelectedTables(string original, string expectedHash, string prepared, string output)
    {
        original = Path.GetFullPath(original); prepared = Path.GetFullPath(prepared); output = Path.GetFullPath(output);
        if (string.Equals(original, output, StringComparison.OrdinalIgnoreCase) || string.Equals(prepared, output, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("章节结果必须另存，不能覆盖输入或中间成果。");
        if (!string.Equals(Path.GetDirectoryName(prepared), Path.GetDirectoryName(output), StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("章节中间成果必须与最终文档位于同一目录。");
        using var sourceFile = new FileStream(original, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(sourceFile)) != expectedHash)
            throw new InvalidDataException("原始文档在处理期间发生变化，停止生成章节结果。");
        sourceFile.Position = 0;
        using var donorFile = new FileStream(prepared, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var source = new ZipArchive(sourceFile, ZipArchiveMode.Read, true);
        using var donor = new ZipArchive(donorFile, ZipArchiveMode.Read, true);
        XDocument ReadMain(ZipArchive zip)
        { using var stream = zip.GetEntry("word/document.xml")!.Open(); return XDocument.Load(stream, LoadOptions.PreserveWhitespace); }
        var document = ReadMain(source);
        var changed = ReadMain(donor);
        var body = document.Root!.Element(Wns + "body")!;
        var donorBody = changed.Root!.Element(Wns + "body")!;
        XElement Locate(XElement node, string id)
        {
            foreach (var index in id.Split("#body/", StringSplitOptions.None)[1].Split('/').Select(int.Parse))
                node = node.Elements().ElementAt(index);
            if (node.Name != Wns + "tbl") throw new InvalidDataException("章节表格位置已变化，停止保存。");
            return node;
        }
        foreach (var id in TableIds) Locate(body, id).ReplaceWith(new XElement(Locate(donorBody, id)));
        using var xmlMemory = new MemoryStream();
        using (var writer = XmlWriter.Create(xmlMemory, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false,
            OmitXmlDeclaration = document.Declaration is null, CloseOutput = false })) document.Save(writer);
        var xmlBytes = xmlMemory.ToArray();
        var stage = Path.Combine(Path.GetDirectoryName(output)!, Path.GetFileNameWithoutExtension(output) + ".章节待校验_" + Guid.NewGuid().ToString("N") + ".docx");
        using (var file = new FileStream(stage, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        using (var result = new ZipArchive(file, ZipArchiveMode.Create))
            foreach (var entry in source.Entries)
            {
                var copy = result.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                copy.LastWriteTime = entry.LastWriteTime; copy.ExternalAttributes = entry.ExternalAttributes;
                using var destination = copy.Open();
                if (entry.FullName == "word/document.xml") destination.Write(xmlBytes);
                else { using var input = entry.Open(); input.CopyTo(destination); }
            }
        using (var verify = ZipFile.OpenRead(stage))
        {
            if (!source.Entries.Select(e => e.FullName).Order().SequenceEqual(verify.Entries.Select(e => e.FullName).Order()))
                throw new InvalidDataException("章节保存改变了文档部件列表。");
            foreach (var entry in source.Entries.Where(e => e.FullName != "word/document.xml"))
            {
                using var before = entry.Open(); using var after = verify.GetEntry(entry.FullName)!.Open();
                if (!System.Security.Cryptography.SHA256.HashData(before).SequenceEqual(System.Security.Cryptography.SHA256.HashData(after)))
                    throw new InvalidDataException("章节保存改变了范围外文档部件：" + entry.FullName);
            }
            if (!TableSynchronizer.SameXml(document.Root!.ToString(SaveOptions.DisableFormatting), ReadMain(verify).Root!.ToString(SaveOptions.DisableFormatting)))
                throw new InvalidDataException("章节保存后的正文校验不一致。");
        }
        var actual = ReportReader.Read(stage);
        var expected = ReportReader.Read(prepared);
        foreach (var id in TableIds)
            if (!TableSynchronizer.SameXml(actual.Tables.Single(t => t.Id == id).OriginalXml, expected.Tables.Single(t => t.Id == id).OriginalXml))
                throw new InvalidDataException("章节表格保存后校验不一致。");
        var stagedHash = SafeFile.Hash(stage);
        var backup = SafeFile.BackupExisting(output);
        File.Move(stage, output, true);
        if (SafeFile.Hash(output) != stagedHash) throw new InvalidDataException("最终章节结果与校验文件不一致。");
        return new SavedFile(output, backup);
    }
}
