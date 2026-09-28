#nullable enable
using System.Xml.Linq;
using AnnotationModal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
namespace 附注工具.核心;

public sealed record TableUpdate(TableInfo Table, TableMappingResult? Mapping, string? UpdatedXml, string? Error)
{
    public bool CanApply => Error is null && UpdatedXml is not null;
}
public sealed record SyncPreview(ReportInfo Report, IReadOnlyList<TableUpdate> Tables,
    IReadOnlyDictionary<string, string> SourceHashes);
public sealed record SyncResult(string OutputPath, int SuccessfulTables, int FailedTables,
    int WrittenCells, bool OutputVerified, IReadOnlyList<string> Errors, string? BackupPath);
public static class TableSynchronizer
{
    public static SyncPreview Preview(string report, IReadOnlyCollection<string>? selectedIds = null, CancellationToken token = default, DalSetting? setting = null)
    {
        var snapshot = ReportReader.Read(report);
        if (selectedIds is not null && selectedIds.Any(id => !snapshot.Tables.Any(t => t.Id == id)))
            throw new InvalidDataException("选中的表格位置已失效，请重新打开报告后选择。");
        var chosen = snapshot.Tables.Where(t => selectedIds?.Contains(t.Id) ?? (t.Link is not null || t.LinkError is not null)).ToArray();
        var updates = new Dictionary<string, TableUpdate>();
        var hashes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var resolved = new List<(TableInfo Table, string Path)>();
        foreach (var table in chosen)
        {
            token.ThrowIfCancellationRequested();
            try
            {
                if (table.LinkError is not null || table.Link is null)
                    throw new InvalidDataException(table.LinkError ?? "这张表尚未建立连接。");
                resolved.Add((table, LinkCodec.ResolveSource(snapshot.FilePath, table.Link)));
            }
            catch (Exception error) when (error is not OperationCanceledException)
            { updates[table.Id] = new(table, null, null, error.Message); }
        }
        // 沿用原UpdateWordTable按来源分组、逐表失败不影响其他表的顺序；同名区域只读一次。
        foreach (var group in resolved.GroupBy(x => x.Path, StringComparer.OrdinalIgnoreCase))
        {
            var regions = new Dictionary<string, ExcelRegion>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in group)
            {
                token.ThrowIfCancellationRequested();
                try
                {
                    var name = item.Table.Link!.LinkName;
                    if (!regions.TryGetValue(name, out var region)) regions[name] = region = ExcelReader.ReadNamedRegion(item.Path, name, setting: setting);
                    if (hashes.TryGetValue(item.Path, out var previous) && previous != region.SourceHash)
                        throw new InvalidDataException("同一底稿在读取过程中发生变化，请重新预览。");
                    hashes[item.Path] = region.SourceHash;
                    var mapping = TableMapper.Build(item.Table, region);
                    if (!mapping.CanApply) throw new InvalidDataException(string.Join("；", mapping.Errors));
                    var xml = XElement.Parse(mapping.UpdatedTable!.OriginalXml);
                    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    var rows = xml.Elements(w + "tr").ToArray();
                    foreach (var pair in mapping.Cells)
                    {
                        var cell = rows[pair.Target.XmlRowIndex].Elements(w + "tc").ElementAt(pair.Target.XmlCellIndex);
                        ReplaceCellText(cell, pair.Source.Text);
                    }
                    updates[item.Table.Id] = new(item.Table, mapping, xml.ToString(SaveOptions.DisableFormatting), null);
                }
                catch (Exception error) when (error is not OperationCanceledException)
                { updates[item.Table.Id] = new(item.Table, null, null, error.Message); }
            }
        }
        return new SyncPreview(snapshot, chosen.Select(t => updates[t.Id]).ToArray(), hashes);
    }
    public static SyncResult Save(SyncPreview preview, string output, CancellationToken token = default)
    {
        var valid = preview.Tables.Where(t => t.CanApply).ToArray();
        if (valid.Length == 0) throw new InvalidOperationException("没有可更新的表格，请先处理预览中的问题。");
        var saved = SafeFile.SaveWordCopy(preview.Report.FilePath, preview.Report.SourceHash, output, document =>
        {
            var body = document.MainDocumentPart!.Document!.Body!;
            foreach (var update in valid)
            {
                token.ThrowIfCancellationRequested();
                var node = FindNode(body, update.Table.NodePath);
                if (node is not Table table || !SameXml(table.OuterXml, update.Table.OriginalXml))
                    throw new InvalidDataException($"表格位置或内容已变化：{update.Table.Title}");
                table.Parent!.ReplaceChild(ParseTableXml(update.UpdatedXml!), table);
            }
        }, file =>
        {
            var reread = ReportReader.Read(file);
            foreach (var before in preview.Report.Tables)
            {
                var after = reread.Tables.SingleOrDefault(t => t.Id == before.Id)
                    ?? throw new InvalidDataException($"保存后找不到原表：{before.Title}");
                var expected = valid.FirstOrDefault(t => t.Table.Id == before.Id)?.UpdatedXml ?? before.OriginalXml;
                if (!SameXml(LinkManager.RebaseTableXml(expected,preview.Report.FilePath,output), after.OriginalXml))
                    throw new InvalidDataException($"保存后表格内容核对不一致：{before.Title}");
            }
        }, token, preview.SourceHashes);
        var errors = preview.Tables.Where(t => !t.CanApply).Select(t => $"{t.Table.Title}：{t.Error}").ToArray();
        return new(saved.OutputPath, valid.Length, errors.Length, valid.Sum(t => t.Mapping!.Cells.Count), true, errors, saved.BackupPath);
    }

    internal static OpenXmlElement FindNode(OpenXmlElement root, int[] path)
    {
        foreach (var index in path)
        {
            if (index < 0 || index >= root.ChildElements.Count) throw new InvalidDataException("报告中的目标位置已失效。");
            root = root.ChildElements[index];
        }
        return root;
    }

    // Open XML SDK 的脱离文档解析会将部分显式格式属性转成等效简写。
    // 在节点树不变的前提下逐项恢复原属性，既不改实际格式，也不放宽最终校验。
    internal static Table ParseTableXml(string xml)
    {
        var original = XElement.Parse(xml);
        var table = new Table(xml);
        RestoreAttributes(table, original);
        if (!SameXml(xml, table.OuterXml))
            throw new InvalidDataException("Word表格XML解析改变了原内容或格式，停止回写。");
        return table;
    }

    private static void RestoreAttributes(OpenXmlElement node, XElement original)
    {
        if (node.LocalName != original.Name.LocalName || node.NamespaceUri != original.Name.NamespaceName)
            throw new InvalidDataException("Word表格XML解析改变了元素名称。");
        var children = node.ChildElements.ToArray();
        var originalChildren = original.Elements().ToArray();
        if (children.Length != originalChildren.Length)
            throw new InvalidDataException("Word表格XML解析改变了原结构。");
        // 先完成该节点的延迟属性解析，再恢复所有真实属性与原词法值。
        _ = node.GetAttributes();
        node.ClearAllAttributes();
        foreach (var attribute in original.Attributes().Where(a => !a.IsNamespaceDeclaration))
        {
            var prefix = original.GetPrefixOfNamespace(attribute.Name.Namespace) ?? "";
            node.SetAttribute(new OpenXmlAttribute(prefix, attribute.Name.LocalName,
                attribute.Name.NamespaceName, attribute.Value));
        }
        foreach (var attribute in original.Attributes().Where(a => a.IsNamespaceDeclaration))
        {
            var prefix = attribute.Name.LocalName == "xmlns" ? "" : attribute.Name.LocalName;
            if (!node.NamespaceDeclarations.Any(d => d.Key == prefix))
                node.AddNamespaceDeclaration(prefix, attribute.Value);
        }
        for (int i = 0; i < children.Length; i++) RestoreAttributes(children[i], originalChildren[i]);
    }
    internal static bool SameXml(string left, string right)
    {
        var a = XElement.Parse(left);
        var b = XElement.Parse(right);
        // XML已把名称解析为“命名空间URI＋名称”；前缀及重复xmlns声明不改变内容。
        // 保留所有真实属性及文字，避免SDK正常调整前缀时误判为数据丢失。
        foreach(var declaration in a.DescendantsAndSelf().Concat(b.DescendantsAndSelf()).Attributes().Where(x=>x.IsNamespaceDeclaration).ToArray()) declaration.Remove();
        // SDK会将<t></t>写成<t/>；仅消除两种空文字写法的差异，保留属性、格式及非空内容。
        foreach (var empty in a.DescendantsAndSelf().Concat(b.DescendantsAndSelf()).Where(e => !e.Nodes().Any()).ToArray())
            empty.RemoveNodes();
        // 属性顺序不影响XML含义，但XNode.DeepEquals按序比较，故只统一顺序。
        // 保留全部属性名称、属性值、文字与元素次序，真实格式变化仍会被拒绝。
        foreach (var element in a.DescendantsAndSelf().Concat(b.DescendantsAndSelf()))
        {
            var attributes = element.Attributes().OrderBy(x => x.Name.ToString(), StringComparer.Ordinal).ToArray();
            element.RemoveAttributes();
            element.Add(attributes);
        }
        return XNode.DeepEquals(a, b);
    }

    internal static void ReplaceCellText(XElement cell, string value)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        if (cell.Descendants(w + "txbxContent").Any())
            throw new InvalidDataException("取数目标格包含文本框，不能把正文金额写入框内文字。");
        if (cell.Descendants().Any(e => e.Name == w + "fldChar" || e.Name == w + "fldSimple" || e.Name == w + "instrText"))
            throw new InvalidDataException("取数目标格含Word域，不能把域结果当作普通文字覆盖。");
        value = value.Replace("\r\n", "\n").Replace('\r', '\n');
        var paragraphs = cell.Elements(w + "p").ToArray();
        var original = string.Join("\n", paragraphs.Select(p => string.Concat(p.Descendants().Select(e =>
            e.Name == w + "t" ? e.Value : e.Name == w + "br" || e.Name == w + "cr" ? "\n" : e.Name == w + "tab" ? "\t" : ""))));
        // 没有数值变化时保留整个原单元格，尤其不能给原空段落补造文字节点。
        if (original == value) return;
        var texts = cell.Descendants(w + "t").ToArray();
        if (texts.Length == 0)
        {
            var paragraph = paragraphs.FirstOrDefault();
            if (paragraph is null) { paragraph = new XElement(w + "p"); cell.Add(paragraph); }
            var run = paragraph.Elements(w + "r").FirstOrDefault();
            if (run is null)
            {
                run = new XElement(w + "r");
                // 空白段落的文字属性记录在段落标记中，新文字须明确沿用本段属性。
                if (paragraph.Element(w + "pPr")?.Element(w + "rPr") is { } properties)
                    run.Add(new XElement(properties));
                paragraph.Add(run);
            }
            var text = new XElement(w + "t");
            run.Add(text);
            texts = [text];
            paragraphs = cell.Elements(w + "p").ToArray();
        }
        var fragments = new List<(XElement Node, XElement? Font, string Text)>();
        var last = Array.FindLastIndex(texts, t => t.Value.Length > 0);
        if (last < 0) last = 0;
        var offset = 0;
        for (var i = 0; i < texts.Length; i++)
        {
            var length = i == last ? value.Length - offset : Math.Min(texts[i].Value.Length, value.Length - offset);
            if (offset + length < value.Length && length > 0 && char.IsHighSurrogate(value[offset + length - 1]) && char.IsLowSurrogate(value[offset + length])) length++;
            fragments.Add((texts[i], texts[i].Ancestors(w + "r").FirstOrDefault()?.Element(w + "rPr"), value.Substring(offset, length)));
            offset += length;
        }
        if (paragraphs.Length == 1 && !value.Contains('\n'))
        {
            // 保留每个原文字片段的格式，将多出的字符交给最后有效片段；旧值的分隔不应残留。
            foreach (var separator in cell.Descendants().Where(e => e.Name == w + "br" || e.Name == w + "cr" || e.Name == w + "tab").ToArray()) separator.Remove();
            foreach (var fragment in fragments) fragment.Node.ReplaceWith(TextNodes(fragment.Text, w, fragment.Node));
            return;
        }
        // 改变段数时只处理普通文字；不能为合并段落顺带丢掉书签、图片或控件。
        string[] protectedNames = ["bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "drawing", "pict", "sdt", "hyperlink", "ins", "del"];
        if (cell.Descendants().Any(e => protectedNames.Any(name => e.Name == w + name))
            || cell.Descendants(w + "p").Count() != paragraphs.Length)
            throw new InvalidDataException("目标格的多段落含链接、图片或其他文字对象，不能直接改变段落数量。");
        var lineCount = value.Count(c => c == '\n') + 1;
        var newParagraphs = new XElement[lineCount];
        for (var i = 0; i < lineCount; i++)
        {
            var paragraph = new XElement(w + "p");
            // Word实测：段数改变时沿用原最后段的设置；段数相同则保留原各段设置。
            var source = paragraphs[paragraphs.Length == lineCount ? i : paragraphs.Length - 1];
            if (source.Element(w + "pPr") is { } properties) paragraph.Add(new XElement(properties));
            newParagraphs[i] = paragraph;
        }
        var line = 0;
        foreach (var fragment in fragments)
        {
            var pieces = fragment.Text.Split('\n');
            for (var i = 0; i < pieces.Length; i++)
            {
                if (i > 0) line++;
                if (pieces[i].Length == 0) continue;
                var run = new XElement(w + "r");
                if (fragment.Font is not null) run.Add(new XElement(fragment.Font));
                run.Add(TextNodes(pieces[i], w, fragment.Node));
                newParagraphs[line].Add(run);
            }
        }
        paragraphs[0].AddBeforeSelf(newParagraphs);
        foreach (var paragraph in paragraphs) paragraph.Remove();
    }

    private static IEnumerable<XElement> TextNodes(string value, XNamespace w, XElement source)
    {
        var pieces = value.Split('\t');
        for (var i = 0; i < pieces.Length; i++)
        {
            if (i > 0) yield return new XElement(w + "tab");
            var text = new XElement(w + "t", source.Attributes().Select(a => new XAttribute(a)), pieces[i]);
            if (pieces[i].Length > 0 && (char.IsWhiteSpace(pieces[i][0]) || char.IsWhiteSpace(pieces[i][^1])))
                text.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            yield return text;
        }
    }
}
