#nullable enable
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using AnnotationModal;
namespace 附注工具.核心;

public static class ReportReader
{
    public static ReportInfo Read(string reportPath)
    {
        var path = Path.GetFullPath(reportPath);
        var bytes = ModernFile.ReadSnapshot(path, ".docx");
        using var stream = new MemoryStream(bytes, false);
        using var document = WordprocessingDocument.Open(stream, false);
        var body = document.MainDocumentPart?.Document?.Body
            ?? throw new InvalidDataException("报告中没有可读取的正文。");
        var tables = new List<TableInfo>();
        foreach (var table in body.Descendants<Table>().Where(t => !t.Ancestors<Table>().Any()))
        {
            var route = GetNodePath(table, body);
            var description = table.GetFirstChild<TableProperties>()?.GetFirstChild<TableDescription>()?.Val?.Value;
            WorkTableInfo? link = null;
            string? error = null;
            if (LinkCodec.IsLink(description))
            {
                try { link = LinkCodec.Parse(description); _=LinkCodec.ResolveSource(path,link); }
                catch (Exception e) when(e is InvalidDataException or ArgumentException or NotSupportedException) { error = e.Message; }
            }
            var title = table.PreviousSibling<Paragraph>()?.InnerText.Trim();
            if (string.IsNullOrWhiteSpace(title)) title = $"表格 {tables.Count + 1}";
            tables.Add(new TableInfo("/word/document.xml#body/" + string.Join('/', route), title, route,
                table.Elements<TableRow>().Count(), table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0,
                link, error, table.OuterXml));
        }
        return new ReportInfo(path, Convert.ToHexString(SHA256.HashData(bytes)), tables);
    }

    private static int[] GetNodePath(OpenXmlElement element, OpenXmlElement root)
    {
        var indices = new List<int>();
        while (!ReferenceEquals(element, root))
        {
            var parent = element.Parent ?? throw new InvalidDataException("表格位置不在报告正文中。");
            indices.Add(parent.ChildElements.ToList().IndexOf(element));
            element = parent;
        }
        indices.Reverse();
        return indices.ToArray();
    }
}

internal static class ModernFile
{
    internal static byte[] ReadSnapshot(string path, params string[] extensions)
    {
        var extension=string.Join("、",extensions);
        if (!extensions.Contains(Path.GetExtension(path),StringComparer.OrdinalIgnoreCase))
            throw new NotSupportedException($"目前支持 {extension}；旧 .doc/.xls 请先另存为支持的格式。");
        var bytes = File.ReadAllBytes(path);
        ReadOnlySpan<byte> oldHeader = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        if (bytes.AsSpan().StartsWith(oldHeader))
            throw new NotSupportedException($"此文件是旧 .doc/.xls 或加密文档。目前支持未加密的 {extension}，修改后缀不能转换文件。");
        return bytes;
    }
}
