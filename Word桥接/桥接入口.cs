#nullable enable
using System.Text;
using System.Xml.Linq;
using AnnotationModal;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using 附注工具.核心;
using W = DocumentFormat.OpenXml.Wordprocessing;

internal static class Program
{
    private static readonly XNamespace Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        try
        {
            if (args.Length != 1) throw new InvalidDataException("请提供一个 UTF-8 JSON 请求文件路径。");
            var request = JObject.Parse(File.ReadAllText(args[0], Encoding.UTF8).TrimStart('\uFEFF'));
            var action = Required(request, "action");
            object result = action switch
            {
                "read" => Read(request),
                "extract" => Extract(request),
                "update" => Update(request),
                _ => throw new InvalidDataException("未知操作：" + action)
            };
            Console.WriteLine(JsonConvert.SerializeObject(result, Formatting.None));
            return 0;
        }
        catch (Exception error)
        {
            Console.WriteLine(JsonConvert.SerializeObject(new { ok = false, error = error.Message }));
            return 1;
        }
    }

    static string Required(JObject request, string field) =>
        (string?)request[field] is { Length: > 0 } value ? value : throw new InvalidDataException("缺少参数：" + field);
    static string Full(JObject request, string field) => Path.GetFullPath(Required(request, field));

    static void ProtectOutputs(IEnumerable<string> inputs, params string[] outputs)
    {
        var used = inputs.Select(Path.GetFullPath).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var path in outputs)
        {
            if (!used.Add(Path.GetFullPath(path))) throw new InvalidDataException("输出不能覆盖输入，多个输出也不能使用同一路径。");
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        }
    }

    static void Backup(string path)
    {
        if (!File.Exists(path)) return;
        var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "BackUp",
            "附注自动更新系统_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + "_" + Guid.NewGuid().ToString("N")[..6]);
        Directory.CreateDirectory(root);
        var target = Path.Combine(root, Path.GetFileName(path));
        File.Copy(path, target, false);
        if (SafeFile.Hash(path) != SafeFile.Hash(target)) throw new IOException("备份校验失败，停止写入。");
    }

    static object Read(JObject request)
    {
        var word = Full(request, "word");
        var snapshot = ReportReader.Read(word);
        var scope = ChapterScope.Resolve(request, snapshot);
        return new { ok = true, action = "read", word, source_hash = snapshot.SourceHash, scope = scope?.Info, tables = Context(word, scope) };
    }

    static object Extract(JObject request)
    {
        var word = Full(request, "word");
        var excel = Full(request, "excel");
        var linked = Full(request, "linked_word");
        var context = Full(request, "context");
        ProtectOutputs([word], excel, linked, context);
        var snapshot = ReportReader.Read(word);
        var scope = ChapterScope.Resolve(request, snapshot);
        var selected = snapshot.Tables.Where(t => scope?.Contains(t) ?? true).ToArray();
        if (selected.Length == 0) throw new InvalidDataException("指定范围内没有可提取的正文表格。");
        // 每张表都必须通过上游逐格回刷校验；不静默遗漏可能包含财务数据的表。
        foreach (var path in new[] { excel, linked, context }) Backup(path);
        var prepared = scope is null ? linked : Path.Combine(Path.GetDirectoryName(linked)!, "章节提取中间成果_" + Guid.NewGuid().ToString("N") + ".docx");
        var result = TableExtractor.WordToExcel(word, selected.Select(t => t.Id).ToArray(), excel, prepared);
        if (scope is not null) scope.SaveSelectedTables(word, snapshot.SourceHash, prepared, linked);
        var afterScope = ChapterScope.Resolve(request, ReportReader.Read(linked));
        var content = new { word, excel, linked_word = linked, scope = afterScope?.Info, tables = Context(linked, afterScope) };
        File.WriteAllText(context, JsonConvert.SerializeObject(content, Formatting.Indented), new UTF8Encoding(true));
        _ = JObject.Parse(File.ReadAllText(context, Encoding.UTF8).TrimStart('\uFEFF'));
        return new { ok = true, action = "extract", excel, linked_word = linked, context, table_count = result.TableCount, scope = afterScope?.Info, warnings = result.Errors };
    }

    static object[] Context(string word, ChapterScope? scope = null)
    {
        var report = ReportReader.Read(word);
        using var document = WordprocessingDocument.Open(word, false);
        var previous = new List<string>();
        var blocks = new List<object>();
        var result = new List<object>();
        int index = 0;
        foreach (var node in document.MainDocumentPart!.Document!.Body!.Descendants())
        {
            if (node is W.Paragraph p && !node.Ancestors<W.Table>().Any() && (scope?.Contains(node) ?? true))
            {
                if (!string.IsNullOrWhiteSpace(p.InnerText))
                {
                    var text = p.InnerText.Trim();
                    previous.Add(text);
                    blocks.Add(new { kind = "paragraph", text });
                }
            }
            if (node is not W.Table || node.Ancestors<W.Table>().Any()) continue;
            var table = report.Tables[index++];
            if (scope is not null && !scope.Contains(table)) continue;
            ExcelRegion? region = null;
            string? source = null;
            if (table.Link is not null && table.LinkError is null)
            {
                source = LinkCodec.ResolveSource(word, table.Link);
                if (File.Exists(source)) region = ExcelReader.ReadNamedRegion(source, table.Link.LinkName);
            }
            result.Add(new
            {
                table_id = table.Id, range_name = table.Link?.LinkName, title = table.Title,
                chapter_context = previous.ToArray(),
                preceding_blocks = blocks.ToArray(),
                unit_text_candidates = previous.Where(p => p.Contains("单位")).ToArray(),
                context_note = "preceding_blocks只记录当前表前原文与表格的先后顺序，不表示各段适用于当前表。table标记表示原文中实际经过的表格；chapter_context是其中的正文段落。unit_text_candidates只是出现单位字样的候选原文，须按说明含义及其所属表判断，不能按距离指定单位。当前表的sheet及行列范围仅用于关联原文，不证明前述所有说明都适用。",
                sheet = region?.Sheet, first_row = region?.FirstRow, first_column = region?.FirstColumn,
                row_count = region?.RowCount ?? table.RowCount, column_count = region?.ColumnCount ?? table.GridColumnCount,
                source, link_error = table.LinkError,
                structural_omissions = StructuralOmissions(table, region)
            });
            blocks.Add(new { kind = "table", table_id = table.Id });
        }
        return result.ToArray();
    }

    static object[] StructuralOmissions(TableInfo table, ExcelRegion? region)
    {
        if (region is null) return [];
        var xml = XElement.Parse(table.OriginalXml);
        var description = (string?)xml.Element(Wns + "tblPr")?.Element(Wns + "tblDescription")?.Attribute(Wns + "val");
        if (!LinkCodec.IsLink(description) || (bool?)JObject.Parse(description![2..^2])["UseGridColumns"] != true) return [];
        var layout = WordTableLayout.Read(table);
        var edges = Edges(layout);
        var omitted = new List<object>();
        // 只投射含独立主格的保留行；整行纵向续接不能投射到上一逻辑行。
        foreach (var row in layout.Cells.GroupBy(c => c.XmlRowIndex).OrderBy(g => g.Key))
        {
            var first = row.First();
            if (first.GridBefore == 0) continue;
            var columns = Array.IndexOf(edges, first.GridBefore);
            if (columns <= 0) throw new InvalidDataException("行首省略范围不能对应到提取后的网格。");
            var number = region.FirstRow + first.WordRowIndex - 1;
            var start = ClosedXML.Excel.XLHelper.GetColumnLetterFromNumber(region.FirstColumn) + number;
            var end = ClosedXML.Excel.XLHelper.GetColumnLetterFromNumber(region.FirstColumn + columns - 1) + number;
            omitted.Add(new { sheet = region.Sheet, range = start == end ? start : start + ":" + end,
                reason = "Word 原表该位置没有单元格" });
        }
        return omitted.ToArray();
    }

    static object Update(JObject request)
    {
        var word = Full(request, "word");
        var excel = Full(request, "excel");
        var output = Full(request, "output");
        ProtectOutputs([word, excel], output);
        var operations = request["operations"] is null ? new JArray() : request["operations"] as JArray ?? throw new InvalidDataException("operations 必须是数组。");
        var snapshot = ReportReader.Read(word);
        var scope = ChapterScope.Resolve(request, snapshot);
        var selected = snapshot.Tables.Where(t => scope?.Contains(t) ?? true).ToArray();
        var linked = selected.Where(t => t.Link is not null && t.LinkError is null).ToArray();
        if (linked.Length == 0) throw new InvalidDataException("文档没有可用的表格连接，请先提取表格。");
        if (selected.Any(t => t.LinkError is not null)) throw new InvalidDataException("文档存在损坏的连接，停止更新。");
        var current = linked.ToDictionary(t => t.Link!.LinkName, StringComparer.OrdinalIgnoreCase);
        foreach (var token in operations)
        {
            if (token is not JObject operation) throw new InvalidDataException("行列变更必须是对象。");
            var range = Required(operation, "range_name");
            if (!current.TryGetValue(range, out var table)) throw new InvalidDataException("变更表格不存在：" + range);
            current[range] = ChangeStructure(table, operation);
        }
        // 验证所有区域后再创建结果，禁止按行列差额自动猜测新增或删除。
        foreach (var table in current.Values)
        {
            var region = ExcelReader.ReadNamedRegion(excel, table.Link!.LinkName);
            var layout = WordTableLayout.Read(table);
            var edges = Edges(layout);
            if (region.RowCount != layout.WordRowCount || region.ColumnCount != edges.Length - 1)
                throw new InvalidDataException(table.Link.LinkName + " 的行列与更新后附注表格不一致，请提供完整的行列变更清单。");
        }
        var intermediate = Path.Combine(Path.GetDirectoryName(output)!, "Word桥接中间文件_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(intermediate);
        var structured = Path.Combine(intermediate, "结构调整.docx");
        SafeFile.SaveWordCopy(word, snapshot.SourceHash, structured, doc =>
        {
            foreach (var table in current.Values)
            {
                var node = TableSynchronizer.FindNode(doc.MainDocumentPart!.Document!.Body!, table.NodePath);
                node.Parent!.ReplaceChild(TableSynchronizer.ParseTableXml(table.OriginalXml), node);
            }
        }, path =>
        {
            var actual = ReportReader.Read(path);
            foreach (var table in current.Values)
            {
                var found = actual.Tables.Single(t => t.Id == table.Id);
                if (!TableSynchronizer.SameXml(LinkManager.RebaseTableXml(table.OriginalXml, word, structured), found.OriginalXml))
                    throw new InvalidDataException("结构保存后校验不一致。");
            }
        });
        var settings = new Dictionary<string, WorkTableInfo>();
        foreach (var table in ReportReader.Read(structured).Tables.Where(t => current.ContainsKey(t.Link?.LinkName ?? "") && (scope?.Contains(t) ?? true)))
        {
            var layout = WordTableLayout.Read(table);
            var clone = JsonConvert.DeserializeObject<WorkTableInfo>(JsonConvert.SerializeObject(table.Link))!;
            clone.SavePath = excel;
            clone.RelativePath = Path.GetRelativePath(Path.GetDirectoryName(structured)!, excel);
            clone.autoInsert = 0;
            clone.RowResize = 0;
            clone.ColResize = Edges(layout).Length - 1 - layout.WordColumnCount;
            settings.Add(table.Id, clone);
        }
        var sourced = Path.Combine(intermediate, "连接换源.docx");
        LinkManager.SetTableLinks(structured, settings, sourced);
        var preview = TableSynchronizer.Preview(sourced, current.Values.Select(t => t.Id).ToArray());
        if (preview.Tables.Any(t => !t.CanApply))
            throw new InvalidDataException(string.Join("；", preview.Tables.Where(t => !t.CanApply).Select(t => t.Table.Title + "：" + t.Error)));
        Backup(output);
        var prepared = scope is null ? output : Path.Combine(Path.GetDirectoryName(output)!, "章节更新中间成果_" + Guid.NewGuid().ToString("N") + ".docx");
        var saved = TableSynchronizer.Save(preview, prepared);
        if (scope is not null) scope.SaveSelectedTables(word, snapshot.SourceHash, prepared, output);
        var afterScope = ChapterScope.Resolve(request, ReportReader.Read(output));
        return new { ok = true, action = "update", output, successful_tables = saved.SuccessfulTables,
            written_cells = saved.WrittenCells, output_verified = saved.OutputVerified, scope = afterScope?.Info, errors = saved.Errors };
    }

    static int[] Edges(WordTableLayout layout) => layout.Cells
        .SelectMany(c => new[] { c.GridStartColumn - 1, c.GridStartColumn + c.GridSpan - 1 }).Distinct().Order().ToArray();

    static TableInfo ChangeStructure(TableInfo table, JObject operation)
    {
        var axis = Required(operation, "axis");
        var action = Required(operation, "action");
        foreach (var field in new[] { "index", "count", "template_index" })
            if (operation[field] is { } value && value.Type != JTokenType.Integer)
                throw new InvalidDataException("行列变更的 " + field + " 必须是整数。");
        int index = (int?)operation["index"] ?? -1;
        int count = (int?)operation["count"] ?? 1;
        int template = (int?)operation["template_index"] ?? -1;
        if (axis is not ("row" or "column") || action is not ("insert" or "delete") || index < 0 || count < 1)
            throw new InvalidDataException("行列变更参数无效。");
        var layout = WordTableLayout.Read(table);
        if (layout.Cells.Any(c => c.GridBefore != 0)) throw new InvalidDataException("行首省略网格的表格暂不能调整结构。");
        var xml = XElement.Parse(table.OriginalXml);
        var rows = xml.Elements(Wns + "tr").ToArray();
        if (rows.Length != layout.WordRowCount) throw new InvalidDataException("含整行纵向合并续接的表格暂不能调整结构。");
        var edges = Edges(layout);
        var size = axis == "row" ? rows.Length : edges.Length - 1;
        if (index > size || (action == "delete" && (index + count > size || count >= size))
            || (action == "insert" && (template < 0 || template >= size)))
            throw new InvalidDataException("行列变更位置、数量或样式来源越界。");
        if (axis == "row")
        {
            GuardVerticalMerges(rows, index, count, action, template);
            if (action == "delete") foreach (var row in rows.Skip(index).Take(count)) row.Remove();
            else for (int i = 0; i < count; i++)
            {
                var clone = new XElement(rows[template]);
                foreach (var cell in clone.Elements(Wns + "tc")) Blank(cell);
                if (index == rows.Length) xml.Add(clone); else rows[index].AddBeforeSelf(clone);
            }
        }
        else
        {
            var start = edges[index];
            var end = action == "delete" ? edges[index + count] : start;
            var grid = xml.Element(Wns + "tblGrid") ?? throw new InvalidDataException("表格缺少列宽网格。");
            var columns = grid.Elements(Wns + "gridCol").ToArray();
            foreach (var row in rows)
            {
                var cells = Cells(row).ToArray();
                if (action == "delete")
                {
                    if (cells.Any(c => c.Start < end && c.End > start && (c.Start < start || c.End > end)))
                        throw new InvalidDataException("删除列切入合并单元格，无法保持原表结构。");
                    foreach (var c in cells.Where(c => c.Start >= start && c.End <= end)) c.Node.Remove();
                }
                else
                {
                    if (cells.Any(c => c.Start < start && c.End > start))
                        throw new InvalidDataException("插入列切入合并单元格，无法保持原表结构。");
                    var source = cells.FirstOrDefault(c => c.Start == edges[template] && c.End == edges[template + 1]);
                    if (source.Node is null) throw new InvalidDataException("样式来源列跨合并单元格，无法复制。");
                    var anchor = cells.FirstOrDefault(c => c.Start == start).Node;
                    for (int i = 0; i < count; i++)
                    {
                        var clone = new XElement(source.Node); Blank(clone);
                        if (anchor is null) row.Add(clone); else anchor.AddBeforeSelf(clone);
                    }
                }
            }
            if (action == "delete") foreach (var col in columns.Skip(start).Take(end - start)) col.Remove();
            else
            {
                var source = columns.Skip(edges[template]).Take(edges[template + 1] - edges[template]).ToArray();
                for (int i = 0; i < count; i++) foreach (var col in source)
                {
                    var clone = new XElement(col);
                    if (start == columns.Length) grid.Add(clone); else columns[start].AddBeforeSelf(clone);
                }
            }
            // 固定表宽随显式增删列调整；原有各列宽度保持，新增列复制来源列宽。
            if (xml.Element(Wns + "tblPr")?.Element(Wns + "tblW") is { } width
                && (string?)width.Attribute(Wns + "type") == "dxa")
                width.SetAttributeValue(Wns + "w", grid.Elements(Wns + "gridCol").Sum(c => (int?)c.Attribute(Wns + "w") ?? 0));
        }
        var changed = table with { OriginalXml = xml.ToString(SaveOptions.DisableFormatting), RowCount = xml.Elements(Wns + "tr").Count() };
        _ = WordTableLayout.Read(changed);
        return changed;
    }

    static IEnumerable<(XElement Node, int Start, int End)> Cells(XElement row)
    {
        int start = 0;
        foreach (var cell in row.Elements(Wns + "tc"))
        {
            int span = (int?)cell.Element(Wns + "tcPr")?.Element(Wns + "gridSpan")?.Attribute(Wns + "val") ?? 1;
            yield return (cell, start, start + span); start += span;
        }
    }

    static void GuardVerticalMerges(XElement[] rows, int index, int count, string action, int template)
    {
        for (int r = 0; r < rows.Length; r++) foreach (var cell in Cells(rows[r]))
        {
            if ((string?)cell.Node.Element(Wns + "tcPr")?.Element(Wns + "vMerge")?.Attribute(Wns + "val") != "restart") continue;
            int end = r + 1;
            while (end < rows.Length)
            {
                var next = Cells(rows[end]).FirstOrDefault(c => c.Start == cell.Start && c.End == cell.End).Node;
                var merge = next?.Element(Wns + "tcPr")?.Element(Wns + "vMerge");
                if (merge is null || (string?)merge.Attribute(Wns + "val") == "restart") break;
                end++;
            }
            if (action == "insert" && ((r < index && index < end) || (r <= template && template < end)))
                throw new InvalidDataException("新增行切入纵向合并，或样式来源属于纵向合并，无法直接复制。");
            if (action == "delete" && r < index + count && end > index && (r < index || end > index + count))
                throw new InvalidDataException("删除行切入纵向合并单元格，无法保持原表结构。");
        }
    }

    static void Blank(XElement cell)
    {
        // 新格只继承表格、段落和字体样式，不复制书签、图片或原值。
        var properties = cell.Element(Wns + "tcPr") is { } cp ? new XElement(cp) : null;
        var paragraph = cell.Element(Wns + "p");
        var p = new XElement(Wns + "p");
        if (paragraph?.Element(Wns + "pPr") is { } pp) p.Add(new XElement(pp));
        var run = new XElement(Wns + "r");
        if (paragraph?.Descendants(Wns + "rPr").FirstOrDefault() is { } rp) run.Add(new XElement(rp));
        run.Add(new XElement(Wns + "t", "")); p.Add(run);
        cell.ReplaceNodes(properties, p);
    }
}
