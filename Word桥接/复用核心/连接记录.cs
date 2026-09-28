#nullable enable
using AnnotationModal;
using Newtonsoft.Json;
using ResourceFactory;
namespace 附注工具.核心;

public static class LinkCodec
{
    public static bool IsLink(string? description) => CommonFunction.IsLinkTable(description);

    public static WorkTableInfo Parse(string? description)
    {
        if (!IsLink(description)) throw new InvalidDataException("这张表没有有效的附注连接标识。");
        WorkTableInfo? link;
        try { link = CommonFunction.GetWorkTableInfo(description!); }
        catch (JsonException error) { throw new InvalidDataException("这张表的连接记录已损坏。", error); }
        if (link is null || string.IsNullOrWhiteSpace(link.LinkName))
            throw new InvalidDataException("这张表的连接记录缺少 Excel 区域名称。");
        return link;
    }

    public static string ResolveSource(string reportPath, WorkTableInfo link)
    {
        // 沿用原相对路径优先规则；用 Windows 路径处理规范化上级目录。
        var path = !string.IsNullOrEmpty(link.RelativePath) ? link.RelativePath : link.SavePath;
        if (string.IsNullOrWhiteSpace(path)) throw new InvalidDataException("这张表没有设置 Excel 来源路径。");
        if (path.Contains('\0')) throw new ArgumentException("这张表的 Excel 来源路径含有无效字符，请重新设置表格链接。");
        // 原版把报告目录从完整路径移除，留下单个开头反斜杠；它仍是报告目录内的相对路径。
        if (!string.IsNullOrEmpty(link.RelativePath) && path.StartsWith('\\') && !path.StartsWith("\\\\", StringComparison.Ordinal))
            path = path[1..];
        return Path.GetFullPath(path, Path.GetDirectoryName(Path.GetFullPath(reportPath))!);
    }
}
public sealed record ReportInfo(string FilePath, string SourceHash, IReadOnlyList<TableInfo> Tables);
public sealed record TableInfo(string Id, string Title, int[] NodePath, int RowCount, int GridColumnCount,
    WorkTableInfo? Link, string? LinkError, string OriginalXml);
