#nullable enable
using System.Xml;
using System.Xml.Linq;
using AnnotationDal;
using Newtonsoft.Json.Linq;

namespace 附注工具.核心;

// XML位置从0开始；Word编号和逻辑网格位置从1开始，三者不能混用。
public sealed record WordCellPosition(string TargetId, int XmlRowIndex, int XmlCellIndex,
    int WordRowIndex, int WordColumnIndex, int GridStartColumn, int GridSpan,
    int GridBefore, string? VerticalMerge, string Text);

public sealed record WordTableLayout(int WordRowCount, int WordColumnCount, IReadOnlyList<WordCellPosition> Cells)
{
    public static WordTableLayout Read(TableInfo table)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var xml = XElement.Parse(table.OriginalXml);
        if (xml.Name != w + "tbl") throw new InvalidDataException("对应位置不是Word表格。");
        if (xml.Descendants(w + "tbl").Any()) throw new InvalidDataException("表内包含嵌套表，不能把内表内容混入外表对应。");
        if (xml.Descendants(w + "hMerge").Any()) throw new InvalidDataException("这张表使用尚未验证的旧式横向合并标记。");
        var rows = xml.Elements(w + "tr").ToArray();
        if (xml.Descendants(w + "tr").Count() != rows.Length || rows.Any(r => r.Descendants(w + "tc").Count() != r.Elements(w + "tc").Count()))
            throw new InvalidDataException("表格包含行或单元格级包装结构，不能略过其中的行格。");
        if (xml.Descendants().Any(e => e.Name == w + "trPrChange" || e.Name == w + "tcPrChange"
            || e.Name == w + "cellIns" || e.Name == w + "cellDel" || e.Name == w + "cellMerge")
            || rows.SelectMany(r => r.Elements(w + "trPr").Descendants()).Any(e => e.Name == w + "ins" || e.Name == w + "del"))
            throw new InvalidDataException("表格含有行或单元格结构修订，需先确认采用的表格结构。");
        if (rows.Any(r => ((int?)r.Element(w + "trPr")?.Element(w + "gridAfter")?.Attribute(w + "val") ?? 0) != 0))
            throw new InvalidDataException("这张表省略了行尾网格，当前尚未取得该结构的Word编号对照。");
        var cells = new List<WordCellPosition>();
        var activeMerges = new HashSet<(int Start, int Span)>();
        int wordRow = 0, wordColumns = 0;
        for (var r = 0; r < rows.Length; r++)
        {
            var rowCells = rows[r].Elements(w + "tc").ToArray();
            var before = (int?)rows[r].Element(w + "trPr")?.Element(w + "gridBefore")?.Attribute(w + "val") ?? 0;
            var gridColumn = before + 1;
            var nextMerges = new HashSet<(int Start, int Span)>();
            wordColumns = Math.Max(wordColumns, rowCells.Length);
            // 实测：全续接行仍留在XML中，但不占Word的行编号；空白独立格仍占行。
            if (rowCells.Any(c => c.Element(w + "tcPr")?.Element(w + "vMerge") is not { } m || (string?)m.Attribute(w + "val") == "restart"))
                wordRow++;
            for (var c = 0; c < rowCells.Length; c++)
            {
                var properties = rowCells[c].Element(w + "tcPr");
                var span = (int?)properties?.Element(w + "gridSpan")?.Attribute(w + "val") ?? 1;
                var merge = properties?.Element(w + "vMerge");
                var mergeValue = merge is null ? null : (string?)merge.Attribute(w + "val") ?? "continue";
                var continued = mergeValue == "continue";
                if (before < 0 || span < 1 || mergeValue is not (null or "restart" or "continue"))
                    throw new InvalidDataException($"物理第{r + 1}行第{c + 1}格的合并标记无效。");
                if (continued && !activeMerges.Contains((gridColumn, span)))
                    throw new InvalidDataException($"物理第{r + 1}行第{c + 1}格没有宽度一致的上方纵向主格。");
                if (merge is not null) nextMerges.Add((gridColumn, span));
                if (!continued)
                    cells.Add(new WordCellPosition($"{table.Id}/tr/{r}/tc/{c}", r, c, wordRow, c + 1,
                        gridColumn, span, before, mergeValue, string.Concat(rowCells[c].Descendants(w + "t").Select(t => t.Value))));
                gridColumn += span;
            }
            activeMerges = nextMerges;
        }
        return new WordTableLayout(wordRow, wordColumns, cells);
    }
}

public sealed record CellMapping(string TargetId, string SourceAddress, WordCellPosition Target, ExcelCellInfo Source);
public sealed record TableInsertion(int BeforeXmlRowIndex, int Count);
public sealed record TableMappingResult(IReadOnlyList<CellMapping> Cells, IReadOnlyList<string> Errors,
    IReadOnlyList<TableInsertion> Insertions, TableInfo? UpdatedTable = null)
{
    public bool CanApply => Errors.Count == 0;
}

public static class TableMapper
{
    public static TableMappingResult Build(TableInfo table, ExcelRegion region)
    {
        TableMappingResult Failure(string message) => new([], [message], []);
        if (table.Link is null || table.LinkError is not null) return Failure(table.LinkError ?? "这张表没有连接记录。");
        WordTableLayout layout;
        try { layout = WordTableLayout.Read(table); }
        catch (Exception e) when (e is InvalidDataException or XmlException or FormatException or OverflowException)
        { return Failure(e.Message); }

        // 新版直接由Excel生成的表，整行纵向续接仍对应Excel的真实行。
        // 原qXmIGG2tP固定RowResize=0，无法补偿Word省略的行号；旧连接继续使用旧编号。
        var physicalRows=false;
        var gridColumns=false;
        XNamespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var description=(string?)XElement.Parse(table.OriginalXml).Element(w+"tblPr")?.Element(w+"tblDescription")?.Attribute(w+"val");
        if(LinkCodec.IsLink(description))
        {
            JToken? flag;JToken? gridFlag;
            try { var record=JObject.Parse(description![2..^2]);flag=record["UsePhysicalRows"];gridFlag=record["UseGridColumns"]; }
            catch(Newtonsoft.Json.JsonException) { return Failure("表格连接记录已损坏。"); }
            if(flag is not null && flag.Type!=JTokenType.Boolean) return Failure("新表行对应标记无效，必须是布尔值。");
            physicalRows=(bool?)flag??false;
            if(gridFlag is not null && gridFlag.Type!=JTokenType.Boolean) return Failure("新表列对应标记无效，必须是布尔值。");
            gridColumns=(bool?)gridFlag??false;
        }
        if(gridColumns && (table.Link.autoInsert!=0 || table.Link.RowResize!=0))
            return Failure("使用网格位置的连接需提供明确行列变更，不能使用隐式插行或行调整。");
        if(physicalRows && table.Link.autoInsert!=0) return Failure("这张表含整行纵向续接，当前不支持自动插行，请将插入点保持0。");

        // 原刷新使用Word接口的数量加Resize校验；Resize不参与取数起点计算。
        if (region.ColumnCount != layout.WordColumnCount + table.Link.ColResize)
            return Failure($"Excel区域有{region.ColumnCount}列，与Word的{layout.WordColumnCount}列及列调整设置不符。");
        if (region.Merges.Any(m => m.FirstRow < region.FirstRow || m.LastRow > region.LastRow
            || m.FirstColumn < region.FirstColumn || m.LastColumn > region.LastColumn))
            return Failure("Excel命名区域的边界切入合并块，不能确定来源主格。");

        var updatedTable = table;
        var insertions = new List<TableInsertion>();
        // 原AnnotationDal.cs 5809～5868行：按Word逻辑行号找首个主格，再在其行上方补差额行。
        var expectedRows=physicalRows ? table.RowCount : layout.WordRowCount;
        var extraRows = region.RowCount - (expectedRows + table.Link.RowResize);
        if (extraRows != 0)
        {
            if (extraRows < 0 || table.Link.autoInsert == 0)
                return Failure($"Excel区域有{region.RowCount}行，与Word的{layout.WordRowCount}行及行调整设置不符；不能自动删行。");
            var at = table.Link.autoInsert > 0 ? table.Link.autoInsert : layout.WordRowCount + table.Link.autoInsert + 1;
            var anchor = layout.Cells.FirstOrDefault(c => c.WordRowIndex == at);
            if (anchor is null) return Failure($"插入点{table.Link.autoInsert}没有对应的Word主格。");
            updatedTable = InsertRows(table, anchor.XmlRowIndex, extraRows);
            layout = WordTableLayout.Read(updatedTable);
            insertions.Add(new(anchor.XmlRowIndex, extraRows));
        }

        // 只有显式标记的新提取连接使用网格坐标，旧连接继续沿用Word编号和MergeCalc。
        var edges=layout.Cells.SelectMany(c=>new[]{c.GridStartColumn,c.GridStartColumn+c.GridSpan}).Distinct().Order().ToArray();
        if(gridColumns)
        {
            if(edges.Length<2 || edges[0]!=1 || region.ColumnCount!=edges.Length-1)
                return Failure("Excel列数与原Word网格边界不一致。");
            var error=ValidateGridSource(updatedTable,layout,region,edges,physicalRows);
            if(error is not null) return Failure(error);
        }
        var calc = new MergeCalc();
        var map = calc.GetMapArray(new(region.FirstRow, region.LastRow, region.FirstColumn, region.LastColumn), calc.GetMergeDic(region.Merges));
        var sources = region.Cells.ToDictionary(c => (c.Row, c.Column));
        var mappings = new List<CellMapping>();
        foreach (var target in layout.Cells)
        {
            // 原5903～5905行：负ColResize允许保留Word中超过来源列数的格。
            if (!gridColumns && target.WordColumnIndex > region.ColumnCount) continue;
            var sourceRow = region.FirstRow + (physicalRows ? target.XmlRowIndex : target.WordRowIndex - 1);
            var sourceColumn = gridColumns ? region.FirstColumn+Array.IndexOf(edges,target.GridStartColumn)
                : region.FirstColumn + target.WordColumnIndex - 1;
            if (!gridColumns && map.TryGetValue(sourceRow, out var columns))
            {
                var index = target.WordColumnIndex - 1;
                if (index < 0 || index >= columns.Length)
                    return Failure($"Word第{target.WordRowIndex}行第{target.WordColumnIndex}格超过Excel该行的有效来源列。");
                sourceColumn = columns[index];
            }
            if (!sources.TryGetValue((sourceRow, sourceColumn), out var source))
                return Failure($"Word第{target.WordRowIndex}行第{target.WordColumnIndex}格对应的来源不在当前区域内。");
            if (!source.CanUpdateReport)
                return Failure($"{region.Sheet}!{source.Address}：{source.Error}");
            mappings.Add(new(target.TargetId, source.Address, target, source));
        }
        return new TableMappingResult(mappings, [], insertions, updatedTable);
    }

    private static string? ValidateGridSource(TableInfo table,WordTableLayout layout,ExcelRegion region,int[] edges,bool physicalRows)
    {
        XNamespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var rows=XElement.Parse(table.OriginalXml).Elements(w+"tr").ToArray();
        var keptRows=layout.Cells.Select(c=>c.XmlRowIndex).Distinct().Order().ToArray();
        foreach(var group in layout.Cells.GroupBy(c=>c.XmlRowIndex))
        {
            var first=group.First();
            if(first.GridBefore==0) continue;
            int end=Array.IndexOf(edges,first.GridBefore+1);
            if(end<0) return "行首省略位置没有对应的网格边界。";
            int row=region.FirstRow+(physicalRows?first.XmlRowIndex:first.WordRowIndex-1);
            int lastColumn=region.FirstColumn+end-1;
            if(region.Cells.Any(c=>c.Row==row && c.Column<=lastColumn && (c.Text.Length>0 || c.IsFormula)))
                return "Excel在Word原本省略的位置填写了数据，不能忽略或补造该单元格。";
            if(region.Merges.Any(m=>m.FirstRow<=row && m.LastRow>=row && m.FirstColumn<=lastColumn && m.LastColumn>=region.FirstColumn))
                return "Excel合并范围跨越Word原本省略的位置，不能更改原表格结构。";
        }
        var expectedMerges=new HashSet<MergeRegion>();
        foreach(var cell in layout.Cells)
        {
            int lastPhysical=cell.XmlRowIndex;
            if(cell.VerticalMerge=="restart")
                for(int row=cell.XmlRowIndex+1;row<rows.Length;row++)
                {
                    int start=((int?)rows[row].Element(w+"trPr")?.Element(w+"gridBefore")?.Attribute(w+"val")??0)+1;
                    XElement? continuation=null;
                    foreach(var candidate in rows[row].Elements(w+"tc"))
                    {
                        int span=(int?)candidate.Element(w+"tcPr")?.Element(w+"gridSpan")?.Attribute(w+"val")??1;
                        if(start==cell.GridStartColumn && span==cell.GridSpan) { continuation=candidate;break; }
                        start+=span;
                    }
                    var merge=continuation?.Element(w+"tcPr")?.Element(w+"vMerge");
                    if(merge is null || ((string?)merge.Attribute(w+"val")??"continue")!="continue") break;
                    lastPhysical=row;
                }
            int firstRow=region.FirstRow+(physicalRows?cell.XmlRowIndex:cell.WordRowIndex-1);
            int lastRow=region.FirstRow+(physicalRows?lastPhysical:keptRows.Count(r=>r<=lastPhysical)-1);
            int firstColumn=region.FirstColumn+Array.IndexOf(edges,cell.GridStartColumn);
            int lastColumn=region.FirstColumn+Array.IndexOf(edges,cell.GridStartColumn+cell.GridSpan)-1;
            if(lastRow>firstRow || lastColumn>firstColumn)
                expectedMerges.Add(new(firstRow,lastRow,firstColumn,lastColumn));
        }
        return expectedMerges.SetEquals(region.Merges) ? null : "Excel合并范围与Word实际单元格的格形不一致，请提供一致的结构变更。";
    }

    private static TableInfo InsertRows(TableInfo table, int beforeXmlRowIndex, int count)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var xml = XElement.Parse(table.OriginalXml);
        var rows = xml.Elements(w + "tr").ToArray();
        var target = rows[beforeXmlRowIndex];
        // Word的25例实测：复制目标行的格形；续接格延长上方合并，restart则成为新的独立空格。
        for (var i = 0; i < count; i++)
        {
            var row = new XElement(w + "tr");
            if (target.Element(w + "trPr") is { } rowProperties) row.Add(new XElement(rowProperties));
            foreach (var sourceCell in target.Elements(w + "tc"))
            {
                var cell = new XElement(w + "tc");
                if (sourceCell.Element(w + "tcPr") is { } cellProperties)
                {
                    var properties = new XElement(cellProperties);
                    if ((string?)properties.Element(w + "vMerge")?.Attribute(w + "val") == "restart")
                        properties.Element(w + "vMerge")!.Remove();
                    cell.Add(properties);
                }
                var paragraph = new XElement(w + "p");
                var sourceParagraph = sourceCell.Elements(w + "p").FirstOrDefault();
                if (sourceParagraph?.Element(w + "pPr") is { } paragraphProperties)
                    paragraph.Add(new XElement(paragraphProperties));
                if (sourceParagraph?.Descendants(w + "r").FirstOrDefault()?.Element(w + "rPr") is { } runProperties)
                    paragraph.Add(new XElement(w + "r", new XElement(runProperties)));
                cell.Add(paragraph);
                row.Add(cell);
            }
            target.AddBeforeSelf(row);
        }
        return table with { OriginalXml = xml.ToString(SaveOptions.DisableFormatting), RowCount = rows.Length + count };
    }
}
