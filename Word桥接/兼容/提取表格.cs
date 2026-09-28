#nullable enable
using AnnotationModal;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using W=DocumentFormat.OpenXml.Wordprocessing;
namespace 附注工具.核心;
public sealed record ReportInsertionPoint(int BodyIndex,string ReportHash);
public sealed record InsertionChoice(string Title,ReportInsertionPoint Point);
public sealed record ExtractionResult(string ReportPath,string ExcelPath,int TableCount,IReadOnlyList<string> Errors);
public static class TableExtractor
{
    private static readonly XNamespace w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    public static ExtractionResult WordToExcel(string report,IReadOnlyCollection<string> selectedIds,string outputExcel,string outputReport,DalSetting? setting=null,CancellationToken token=default)
    {
        outputExcel=Path.GetFullPath(outputExcel);outputReport=Path.GetFullPath(outputReport);
        if(!Path.GetExtension(outputExcel).Equals(".xlsx",StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("新提取的底稿请保存为 .xlsx。");
        if(Path.GetFullPath(report).Equals(outputReport,StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("请选择新的报告副本，不能覆盖原报告。");
        using var sourceLock=new FileStream(report,FileMode.Open,FileAccess.Read,FileShare.Read);
        var snapshot=ReportReader.Read(report);var selected=snapshot.Tables.Where(t=>selectedIds.Contains(t.Id)).ToArray();
        if(selected.Length==0 || selectedIds.Any(id=>!selected.Any(t=>t.Id==id))) throw new InvalidDataException("请选择报告中实际存在的表格。");
        using var doc=WordprocessingDocument.Open(report,false);using var book=new XLWorkbook();var sheet=book.AddWorksheet("sheet1");
        var links=new Dictionary<string,WorkTableInfo>();var originalTexts=new Dictionary<string,string[]>();var warnings=new HashSet<string>();int startRow=2;
        foreach(var info in selected)
        {
            token.ThrowIfCancellationRequested();
            var xml=XElement.Parse(info.OriginalXml);var layout=WordTableLayout.Read(info);
            // 新提取连接显式使用网格边界，保留行首省略位置，不能把Word列号当网格列。
            var useGridColumns=layout.Cells.Any(c=>c.GridBefore!=0);
            var rows=xml.Elements(w+"tr").ToArray();var grid=xml.Element(w+"tblGrid")?.Elements(w+"gridCol").Select(c=>(double?)c.Attribute(w+"w")??0).ToArray()??[];
            var edges=layout.Cells.SelectMany(c=>new[]{c.GridStartColumn,c.GridStartColumn+c.GridSpan}).Distinct().Order().ToArray();
            if(edges.Length<2 || edges[0]!=1 || edges[^1]-1>grid.Length || grid.Any(n=>n<=0)) throw new InvalidDataException($"{info.Title}：表格没有完整且明确的列宽，不能猜测合并边界。");
            var keptRows=layout.Cells.Select(c=>c.XmlRowIndex).Distinct().Order().ToArray();int columns=edges.Length-1;
            NoteWordStyles(xml,doc,warnings);
            for(int r=0;r<keptRows.Length;r++)
            {
                int end=r+1<keptRows.Length?keptRows[r+1]:rows.Length;
                var heights=rows.Skip(keptRows[r]).Take(end-keptRows[r]).Select(row=>row.Element(w+"trPr")?.Element(w+"trHeight"))
                    .Where(h=>h is not null && (string?)h.Attribute(w+"hRule")!="auto").Select(h=>(double?)h!.Attribute(w+"val")??0).ToArray();
                if(heights.Any(h=>h>0)) sheet.Row(startRow+r).Height=heights.Sum()/20d;
            }
            for(int c=0;c<columns;c++) sheet.Column(c+2).Width=Math.Max(sheet.Column(c+2).Width,grid.Skip(edges[c]-1).Take(edges[c+1]-edges[c]).Sum()/120d);
            // 原rWEZoJ3Go：按旧列号取最小格宽（点）再除以6，只增宽不缩窄。
            foreach(var group in layout.Cells.Where(_=>!useGridColumns).GroupBy(c=>c.WordColumnIndex))
            {
                var width=group.Select(c=>grid.Skip(c.GridStartColumn-1).Take(c.GridSpan).Sum()/20d).Where(n=>n<9999).DefaultIfEmpty(0).Min()/6d;
                sheet.Column(group.Key+1).Width=Math.Max(sheet.Column(group.Key+1).Width,width);
            }
            var values=new List<string>();
            foreach(var position in layout.Cells)
            {
                var cell=rows[position.XmlRowIndex].Elements(w+"tc").ElementAt(position.XmlCellIndex);var text=CellText(cell);values.Add(text);
                TableSynchronizer.ReplaceCellText(new XElement(cell),text);
                int firstRow=startRow+position.WordRowIndex-1,lastPhysical=position.XmlRowIndex;
                if(position.VerticalMerge=="restart")
                    for(int row=position.XmlRowIndex+1;row<rows.Length;row++)
                    {
                        var continuation=RowCells(rows[row]).FirstOrDefault(c=>c.Start==position.GridStartColumn && c.Span==position.GridSpan);
                        if(continuation.Node is null || continuation.Merge!="continue") break;lastPhysical=row;
                    }
                int lastRow=startRow+keptRows.Count(r=>r<=lastPhysical)-1;
                int firstColumn=Array.IndexOf(edges,position.GridStartColumn)+2,lastColumn=Array.IndexOf(edges,position.GridStartColumn+position.GridSpan)+1;
                var target=sheet.Range(firstRow,firstColumn,lastRow,lastColumn);
                if(lastRow>firstRow || lastColumn>firstColumn) target.Merge();
                var bottom=RowCells(rows[lastPhysical]).First(c=>c.Start==position.GridStartColumn && c.Span==position.GridSpan).Node;
                SetExcelValue(target.FirstCell(),text);CopyWordStyle(cell,bottom,target,doc.MainDocumentPart?.ThemePart,warnings);
            }
            var node=(W.Table)TableSynchronizer.FindNode(doc.MainDocumentPart!.Document!.Body!,info.NodePath);
            if(node.PreviousSibling() is W.Paragraph heading) SetExcelValue(sheet.Cell(startRow-1,2),heading.InnerText);
            var range=sheet.Range(startRow,2,startRow+layout.WordRowCount-1,columns+1);
            var name="TA_"+(snapshot.Tables.ToList().FindIndex(t=>t.Id==info.Id)+1).ToString("000",CultureInfo.InvariantCulture);
            if(book.DefinedNames.Any(n=>n.Name.Equals(name,StringComparison.OrdinalIgnoreCase))) throw new InvalidDataException($"命名区域重复：{name}。");
            range.AddToNamed(name,XLScope.Workbook);
            if(setting?.ShowColor==true) range.Style.Fill.BackgroundColor=XLColor.FromArgb(238,236,225);
            links[info.Id]=new WorkTableInfo{SavePath=outputExcel,RelativePath=Path.GetRelativePath(Path.GetDirectoryName(snapshot.FilePath)!,outputExcel),LinkName=name,RowResize=0,ColResize=columns-layout.WordColumnCount,autoInsert=0};
            originalTexts[info.Id]=values.ToArray();startRow+=layout.WordRowCount+3;
        }
        var stagedExcel=Temporary(outputExcel);book.SaveAs(stagedExcel);
        foreach(var info in selected)
        {
            var region=ExcelReader.ReadNamedRegion(stagedExcel,links[info.Id].LinkName);
            var mappingTable=new W.Table(info.OriginalXml);
            var properties=mappingTable.GetFirstChild<W.TableProperties>() ?? mappingTable.PrependChild(new W.TableProperties());
            var description=properties.GetFirstChild<W.TableDescription>();
            var linkRecord=JObject.FromObject(links[info.Id]);
            if(WordTableLayout.Read(info).Cells.Any(c=>c.GridBefore!=0)) linkRecord["UseGridColumns"]=true;
            var linkDescription="<#"+linkRecord.ToString(Formatting.None)+"#>";
            if(description is null) properties.AddChild(new W.TableDescription{Val=linkDescription},true);
            else description.Val=linkDescription;
            var mapping=TableMapper.Build(info with{Link=links[info.Id],LinkError=null,OriginalXml=mappingTable.OuterXml},region);
            if(!mapping.CanApply || !mapping.Cells.Select(c=>c.Source.Text).SequenceEqual(originalTexts[info.Id]))
                throw new InvalidDataException($"{info.Title}：提取后逐格回刷核对不一致，未发布新连接。"+string.Join("；",mapping.Errors));
        }
        var stagedReport=Temporary(outputReport);
        LinkManager.SetTableLinks(snapshot.FilePath,links,stagedReport,clearPhysicalRowMode:true);
        var reread=ReportReader.Read(stagedReport);
        foreach(var info in selected)
        {
            var after=reread.Tables.Single(t=>t.Id==info.Id);
            if(!LinkCodec.ResolveSource(stagedReport,after.Link!).Equals(outputExcel,StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("新报告连接没有指向将交付的底稿。");
        }
        token.ThrowIfCancellationRequested();SafeFile.CommitPrepared((stagedExcel,outputExcel),(stagedReport,outputReport));
        return new(outputReport,outputExcel,selected.Length,warnings.ToArray());
    }
    public static ExtractionResult ExcelToWord(string excel,IReadOnlyCollection<string> rangeNames,string report,ReportInsertionPoint insertionPoint,string outputReport,bool hasHeader=true,CancellationToken token=default)
    {
        excel=Path.GetFullPath(excel);outputReport=Path.GetFullPath(outputReport);
        using var sourceLock=new FileStream(excel,FileMode.Open,FileAccess.Read,FileShare.Read);
        var snapshot=ReportReader.Read(report);
        if(snapshot.SourceHash!=insertionPoint.ReportHash) throw new InvalidDataException("报告在选择插入点后发生变化，请重新选择位置。");
        var index=ExcelReader.GetWorkbookInfo(excel);var selected=index.NamedRegions.Where(n=>rangeNames.Contains(n.DisplayName)&&n.IsVisible).ToArray();
        if(selected.Length==0 || rangeNames.Any(n=>!selected.Any(s=>s.DisplayName==n))) throw new InvalidDataException("请选择实际存在且可见的命名区域。");
        using var memory=new MemoryStream(File.ReadAllBytes(excel),false);using var book=new XLWorkbook(memory);
        using var sourceXml=SpreadsheetDocument.Open(excel,false);
        var inserted=new List<OpenXmlElement>();
        foreach(var name in selected)
        {
            token.ThrowIfCancellationRequested();var region=ExcelReader.ReadNamedRegion(excel,name.DisplayName);
            if(region.SourceHash!=index.SourceHash) throw new InvalidDataException("Excel在读取过程中发生变化，请重新选择。");
            if(region.Cells.Any(c=>!c.CanUpdateReport)) throw new InvalidDataException($"{name.DisplayName}："+string.Join("；",region.Cells.Where(c=>!c.CanUpdateReport).Select(c=>c.Address+" "+c.Error)));
            if(region.Merges.Any(m=>m.FirstRow<region.FirstRow || m.LastRow>region.LastRow || m.FirstColumn<region.FirstColumn || m.LastColumn>region.LastColumn)) throw new InvalidDataException("所选名称切入合并块，不能提取不完整的合并表。");
            if(hasHeader && region.FirstRow>1)
            {
                var title=ExcelReader.ReadCell(excel,region.Sheet,XLHelper.GetColumnLetterFromNumber(region.FirstColumn)+(region.FirstRow-1));
                if(!title.CanUpdateReport) throw new InvalidDataException("表前标题无法读取："+title.Error);
                if(title.Text.Length>0) inserted.Add(Paragraph(title.Text));
            }
            var table=CreateWordTable(region,book.Worksheet(region.Sheet),sourceXml.WorkbookPart!);
            var temporary=new TableInfo("新表",name.DisplayName,[],region.RowCount,region.ColumnCount,null,null,table.OuterXml);
            var layout=WordTableLayout.Read(temporary);
            var link=new WorkTableInfo{SavePath=excel,RelativePath=Path.GetRelativePath(Path.GetDirectoryName(snapshot.FilePath)!,excel),LinkName=name.DisplayName,RowResize=0,ColResize=region.ColumnCount-layout.WordColumnCount,autoInsert=0};
            var record=JObject.FromObject(link);
            if(layout.WordRowCount!=region.RowCount) record["UsePhysicalRows"]=true;
            table.GetFirstChild<W.TableProperties>()!.AddChild(new W.TableDescription{Val="<#"+record.ToString(Formatting.None)+"#>"},true);
            var mapped=TableMapper.Build(temporary with{OriginalXml=table.OuterXml,Link=link},region);
            if(!mapped.CanApply) throw new InvalidDataException("提入Word后的合并对应不成立："+string.Join("；",mapped.Errors));
            inserted.Add(table);inserted.Add(new W.Paragraph());
        }
        string? expected=null;
        SafeFile.SaveWordCopy(snapshot.FilePath,snapshot.SourceHash,outputReport,doc=>
        {
            var body=doc.MainDocumentPart!.Document!.Body!;int last=body.ChildElements.Count;
            if(body.LastChild is W.SectionProperties) last--;
            if(insertionPoint.BodyIndex<0 || insertionPoint.BodyIndex>last) throw new InvalidDataException("报告插入位置已失效。");
            var anchor=insertionPoint.BodyIndex<body.ChildElements.Count ? body.ChildElements[insertionPoint.BodyIndex] : null;
            foreach(var node in inserted) {if(anchor is null) body.Append(node);else body.InsertBefore(node,anchor);}
            expected=body.OuterXml;
        },file=>
        {
            using var actual=WordprocessingDocument.Open(file,false);
            var expectedBody=new W.Body(expected!);
            foreach(var table in expectedBody.Descendants<W.Table>().ToArray())
                table.Parent!.ReplaceChild(new W.Table(LinkManager.RebaseTableXml(table.OuterXml,snapshot.FilePath,outputReport)),table);
            if(!TableSynchronizer.SameXml(expectedBody.OuterXml,actual.MainDocumentPart!.Document!.Body!.OuterXml)) throw new InvalidDataException("提入表格后报告内容校验不一致。");
            var result=ReportReader.Read(file);
            if(result.Tables.Count!=snapshot.Tables.Count+selected.Length) throw new InvalidDataException("提入后的表格数量不一致。");
        },token,new Dictionary<string,string>{{excel,index.SourceHash}});
        return new(outputReport,excel,selected.Length,[]);
    }
    public static IReadOnlyList<InsertionChoice> ReadInsertionPoints(string report)
    {
        var snapshot=ReportReader.Read(report);using var doc=WordprocessingDocument.Open(report,false);var body=doc.MainDocumentPart!.Document!.Body!;
        var choices=body.ChildElements.TakeWhile(e=>e is not W.SectionProperties).Select((e,i)=>new InsertionChoice("插在前面："+(e.InnerText.Length>60?e.InnerText[..60]+"…":e.InnerText),new(i,snapshot.SourceHash))).ToList();
        choices.Add(new("文末",new(choices.Count,snapshot.SourceHash)));return choices;
    }
    private static string Temporary(string output)
    {
        var folder=Path.GetDirectoryName(output)!;Directory.CreateDirectory(folder);
        return Path.Combine(folder,Path.GetFileNameWithoutExtension(output)+".提表待校验_"+Guid.NewGuid().ToString("N")+Path.GetExtension(output));
    }
    private static IEnumerable<(XElement Node,int Start,int Span,string? Merge)> RowCells(XElement row)
    {
        int column=((int?)row.Element(w+"trPr")?.Element(w+"gridBefore")?.Attribute(w+"val")??0)+1;
        foreach(var cell in row.Elements(w+"tc"))
        {
            int span=(int?)cell.Element(w+"tcPr")?.Element(w+"gridSpan")?.Attribute(w+"val")??1;
            var merge=cell.Element(w+"tcPr")?.Element(w+"vMerge");
            yield return(cell,column,span,merge is null?null:(string?)merge.Attribute(w+"val")??"continue");column+=span;
        }
    }
    private static string CellText(XElement cell)=>string.Join("\n",cell.Elements(w+"p").Select(p=>string.Concat(p.Descendants().Select(e=>e.Name==w+"t"?e.Value:e.Name==w+"br"||e.Name==w+"cr"?"\n":e.Name==w+"tab"?"\t":""))));
    private static void SetExcelValue(IXLCell cell,string text)
    {
        // 剪贴板替代：仅将显示可以逐字复原的明确金额、百分比写成数值；编号与其余文字保持字面。
        var numeric=text;bool percent=numeric.EndsWith('%'),brackets=numeric.StartsWith('(')&&numeric.EndsWith(')');
        if(percent) numeric=numeric[..^1];if(brackets) numeric="-"+numeric[1..^1];
        if(Regex.IsMatch(numeric,@"\A-?(?:0|[1-9][0-9]{0,14}|[1-9][0-9]{0,2}(?:,[0-9]{3})+)(?:\.[0-9]+)?\z")
            && numeric.Count(char.IsDigit)<=15 && double.TryParse(numeric,NumberStyles.Number,CultureInfo.InvariantCulture,out var value))
        {
            var decimals=numeric.Contains('.')?numeric.Length-numeric.IndexOf('.')-1:0;
            var format=(numeric.Contains(',')?"#,##0":"0")+(decimals>0?"."+new string('0',decimals):"")+(percent?"%":"");
            if(brackets) format+=";("+format+")";
            cell.Value=percent?value/100d:value;cell.Style.NumberFormat.Format=format;
            if(cell.GetFormattedString(CultureInfo.InvariantCulture)==text) return;
        }
        cell.Style.NumberFormat.Format="@";cell.Value=text.StartsWith('\'')?"'"+text:text;
    }
    private static void CopyWordStyle(XElement cell,XElement bottom,IXLRange target,ThemePart? theme,HashSet<string> warnings)
    {
        var font=cell.Descendants(w+"rPr").FirstOrDefault();
        if(font is not null)
        {
            if(font.Element(w+"rFonts") is { } fonts) target.Style.Font.FontName=(string?)fonts.Attribute(w+"eastAsia")??(string?)fonts.Attribute(w+"ascii")??target.Style.Font.FontName;
            if((double?)font.Element(w+"sz")?.Attribute(w+"val") is { } size && size>0) target.Style.Font.FontSize=size/2;
            target.Style.Font.Bold=font.Element(w+"b") is { } bold && (string?)bold.Attribute(w+"val") is not ("0" or "false");
            target.Style.Font.Italic=font.Element(w+"i") is { } italic && (string?)italic.Attribute(w+"val") is not ("0" or "false");
            if(WordColor(font.Element(w+"color"),"val","themeColor",theme) is { } color) target.Style.Font.FontColor=color;
            target.Style.Font.Underline=(string?)font.Element(w+"u")?.Attribute(w+"val") switch
            {
                "double"=>XLFontUnderlineValues.Double,"none" or null=>XLFontUnderlineValues.None,
                "single" or "words"=>XLFontUnderlineValues.Single,
                _=>throw new InvalidDataException("此单元格使用尚不能等效提取的下划线样式。")
            };
        }
        target.Style.Alignment.WrapText=true;
        var alignment=(string?)cell.Element(w+"p")?.Element(w+"pPr")?.Element(w+"jc")?.Attribute(w+"val");
        target.Style.Alignment.Horizontal=alignment switch {"center"=>XLAlignmentHorizontalValues.Center,"right" or "end"=>XLAlignmentHorizontalValues.Right,"both"=>XLAlignmentHorizontalValues.Justify,"distribute"=>XLAlignmentHorizontalValues.Distributed,_=>XLAlignmentHorizontalValues.Left};
        target.Style.Alignment.Vertical=(string?)cell.Element(w+"tcPr")?.Element(w+"vAlign")?.Attribute(w+"val") switch
        {"center"=>XLAlignmentVerticalValues.Center,"bottom"=>XLAlignmentVerticalValues.Bottom,_=>XLAlignmentVerticalValues.Top};
        if(WordColor(cell.Element(w+"tcPr")?.Element(w+"shd"),"fill","themeFill",theme) is { } fill) target.Style.Fill.BackgroundColor=fill;
        var table=cell.Ancestors(w+"tbl").First();var rows=table.Elements(w+"tr").ToArray();var shared=table.Element(w+"tblPr")?.Element(w+"tblBorders");
        XElement? Edge(XElement at,string side,string fallback)=>at.Element(w+"tcPr")?.Element(w+"tcBorders")?.Element(w+side)
            ??cell.Element(w+"tcPr")?.Element(w+"tcBorders")?.Element(w+side)??shared?.Element(w+fallback);
        ApplyWordBorder(target.FirstRow(),"top",Edge(cell,"top",cell.Parent==rows[0]?"top":"insideH"),theme);
        ApplyWordBorder(target.LastRow(),"bottom",Edge(bottom,"bottom",bottom.Parent==rows[^1]?"bottom":"insideH"),theme);
        var anchor=RowCells(cell.Parent!).Single(c=>c.Node==cell);int first=Array.IndexOf(rows,cell.Parent),last=Array.IndexOf(rows,bottom.Parent),line=0;
        var left=cell==cell.Parent!.Elements(w+"tc").First()?"left":"insideV";
        var right=cell==cell.Parent!.Elements(w+"tc").Last()?"right":"insideV";var previous=cell;
        foreach(var row in rows.Skip(first).Take(last-first+1))
        {
            var parts=RowCells(row).ToArray();var at=parts.Single(c=>c.Start==anchor.Start&&c.Span==anchor.Span).Node;
            if(parts.All(c=>c.Merge=="continue"))
            {
                if(!XElement.DeepEquals(Edge(at,"left",left),Edge(previous,"left",left)) || !XElement.DeepEquals(Edge(at,"right",right),Edge(previous,"right",right)))
                    warnings.Add("整行合并续接按原编号压成一行，其中分段侧边框无法全部保留，需要复核。");
                continue;
            }
            // 纵向合并的每段侧边可有不同颜色，按实际保留下来的行逐段复制。
            var destination=target.Row(++line);
            ApplyWordBorder(destination.FirstCell().AsRange(),"left",Edge(at,"left",left),theme);
            ApplyWordBorder(destination.LastCell().AsRange(),"right",Edge(at,"right",right),theme);previous=at;
        }
    }
    private static W.Table CreateWordTable(ExcelRegion region,IXLWorksheet sheet,WorkbookPart sourcePart)
    {
        var table=new W.Table(new W.TableProperties(new W.TableWidth{Type=W.TableWidthUnitValues.Auto,Width="0"}),new W.TableGrid());
        var grid=table.GetFirstChild<W.TableGrid>()!;
        for(int c=region.FirstColumn;c<=region.LastColumn;c++) grid.Append(new W.GridColumn{Width=Math.Max(120,(int)Math.Round((sheet.Column(c).Width*7+5)*15)).ToString(CultureInfo.InvariantCulture)});
        var values=region.Cells.ToDictionary(c=>(c.Row,c.Column));
        for(int r=region.FirstRow;r<=region.LastRow;r++)
        {
            var row=new W.TableRow(new W.TableRowProperties(new W.TableRowHeight{Val=(uint)Math.Round(sheet.Row(r).Height*20),HeightType=W.HeightRuleValues.AtLeast}));
            for(int c=region.FirstColumn;c<=region.LastColumn;)
            {
                var merge=region.Merges.FirstOrDefault(m=>m.FirstRow<=r&&m.LastRow>=r&&m.FirstColumn<=c&&m.LastColumn>=c);
                if(merge is not null && c!=merge.FirstColumn) {c++;continue;}
                int span=merge is null?1:merge.LastColumn-merge.FirstColumn+1;bool continued=merge is not null&&r>merge.FirstRow;
                var properties=new W.TableCellProperties();if(span>1) properties.Append(new W.GridSpan{Val=span});
                if(merge is not null&&merge.LastRow>merge.FirstRow) properties.Append(new W.VerticalMerge{Val=continued?W.MergedCellValues.Continue:W.MergedCellValues.Restart});
                var cell=new W.TableCell(properties);var source=sheet.Cell(merge?.FirstRow??r,c);var text=continued?"":values[(r,c)].Text;
                if(source.Style.Fill.PatternType==XLFillPatternValues.Solid)
                    properties.AddChild(new W.Shading{Val=W.ShadingPatternValues.Clear,Fill=ExcelColor(source.Style.Fill.BackgroundColor,sourcePart),Color="auto"},true);
                properties.AddChild(new W.TableCellVerticalAlignment{Val=source.Style.Alignment.Vertical switch
                {XLAlignmentVerticalValues.Center=>W.TableVerticalAlignmentValues.Center,XLAlignmentVerticalValues.Bottom=>W.TableVerticalAlignmentValues.Bottom,_=>W.TableVerticalAlignmentValues.Top}},true);
                // 横向合并的右框在最右格；纵向合并的底框在最后续接行，不能一律取左上格。
                var border=new W.TableCellBorders();
                border.Append(ExcelBorder("top",sheet.Range(r,c,r,c+span-1),sourcePart),ExcelBorder("left",sheet.Range(r,c,r,c),sourcePart),
                    ExcelBorder("bottom",sheet.Range(r,c,r,c+span-1),sourcePart),ExcelBorder("right",sheet.Range(r,c+span-1,r,c+span-1),sourcePart));
                properties.AddChild(border,true);
                foreach(var line in text.Replace("\r\n","\n").Replace('\r','\n').Split('\n'))
                {
                    var paragraph=Paragraph(line);var run=paragraph.GetFirstChild<W.Run>()!;var font=source.Style.Font;
                    var runProperties=new W.RunProperties(new W.RunFonts{Ascii=font.FontName,HighAnsi=font.FontName,EastAsia=font.FontName},new W.FontSize{Val=(font.FontSize*2).ToString(CultureInfo.InvariantCulture)});
                    if(font.Bold) runProperties.AddChild(new W.Bold(),true);if(font.Italic) runProperties.AddChild(new W.Italic(),true);
                    runProperties.AddChild(new W.Color{Val=ExcelColor(font.FontColor,sourcePart)},true);
                    if(font.Underline!=XLFontUnderlineValues.None) runProperties.AddChild(new W.Underline{Val=font.Underline is XLFontUnderlineValues.Double or XLFontUnderlineValues.DoubleAccounting?W.UnderlineValues.Double:W.UnderlineValues.Single},true);
                    run.PrependChild(runProperties);
                    var alignment=source.Style.Alignment.Horizontal switch
                    {
                        XLAlignmentHorizontalValues.Center or XLAlignmentHorizontalValues.CenterContinuous=>W.JustificationValues.Center,
                        XLAlignmentHorizontalValues.Right=>W.JustificationValues.Right,
                        XLAlignmentHorizontalValues.Justify=>W.JustificationValues.Both,
                        XLAlignmentHorizontalValues.Distributed=>W.JustificationValues.Distribute,
                        XLAlignmentHorizontalValues.General when source.CachedValue.IsNumber||source.CachedValue.IsDateTime=>W.JustificationValues.Right,
                        XLAlignmentHorizontalValues.General when source.CachedValue.IsBoolean=>W.JustificationValues.Center,
                        _=>W.JustificationValues.Left
                    };
                    paragraph.PrependChild(new W.ParagraphProperties(new W.Justification{Val=alignment}));
                    cell.Append(paragraph);
                }
                row.Append(cell);c+=span;
            }
            table.Append(row);
        }
        return table;
    }
    private static void NoteWordStyles(XElement table,WordprocessingDocument doc,HashSet<string> warnings)
    {
        var styles=doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if(styles is null) return;
        var source=XElement.Parse(styles.OuterXml);
        if(source.Element(w+"docDefaults")?.Descendants().Any(n=>n.HasElements && (n.Name==w+"rPr" || n.Name==w+"pPr"))==true)
            warnings.Add("原表涉及文档默认的继承格式；本次仅复制明确的直接格式，继承部分需要复核。");
        var defaults=source.Elements(w+"style").Where(n=>(string?)n.Attribute(w+"default") is "1" or "true" or "on")
            .Select(n=>(string?)n.Attribute(w+"styleId"));
        var referenced=table.Descendants().Where(n=>n.Name==w+"tblStyle"||n.Name==w+"pStyle"||n.Name==w+"rStyle")
            .Select(n=>(string?)n.Attribute(w+"val"));
        foreach(var id in referenced.Concat(defaults).Distinct())
        {
            var seen=new HashSet<string>();var current=id;
            while(current is not null && seen.Add(current))
            {
                var style=source.Elements(w+"style").FirstOrDefault(s=>(string?)s.Attribute(w+"styleId")==current);
                if(style?.Elements().Any(n=>n.HasElements && n.Name.LocalName is "tblPr" or "tcPr" or "pPr" or "rPr" or "tblStylePr")==true)
                {
                    warnings.Add($"原表使用样式“{id}”的继承格式；本次仅复制明确的直接格式，继承部分需要复核。");break;
                }
                current=(string?)style?.Element(w+"basedOn")?.Attribute(w+"val");
            }
        }
    }
    private static void ApplyWordBorder(IXLRangeBase edge,string side,XElement? node,ThemePart? theme)
    {
        if(node is null) return;
        var size=(int?)node.Attribute(w+"sz")??4;
        var style=(string?)node.Attribute(w+"val") switch
        {
            null or "nil" or "none"=>XLBorderStyleValues.None,"double"=>XLBorderStyleValues.Double,
            "single"=>size>=16?XLBorderStyleValues.Thick:size>=12?XLBorderStyleValues.Medium:XLBorderStyleValues.Thin,
            "dotted"=>XLBorderStyleValues.Dotted,"dashed" or "dashSmallGap"=>size>=12?XLBorderStyleValues.MediumDashed:XLBorderStyleValues.Dashed,
            "dotDash"=>size>=12?XLBorderStyleValues.MediumDashDot:XLBorderStyleValues.DashDot,
            "dotDotDash"=>size>=12?XLBorderStyleValues.MediumDashDotDot:XLBorderStyleValues.DashDotDot,
            "dashDotStroked"=>XLBorderStyleValues.SlantDashDot,
            _=>throw new InvalidDataException("表格存在尚不能等效提取的装饰边框样式。")
        };
        var color=WordColor(node,"color","themeColor",theme)??XLColor.Black;var border=edge.Style.Border;
        switch(side)
        {
            case "top":border.TopBorder=style;border.TopBorderColor=color;break;
            case "bottom":border.BottomBorder=style;border.BottomBorderColor=color;break;
            case "left":border.LeftBorder=style;border.LeftBorderColor=color;break;
            case "right":border.RightBorder=style;border.RightBorderColor=color;break;
        }
    }
    private static W.BorderType ExcelBorder(string side,IXLRange edge,WorkbookPart source)
    {
        var values=edge.Cells().Select(c=>side switch
        {
            "top"=>(c.Style.Border.TopBorder,c.Style.Border.TopBorderColor),"bottom"=>(c.Style.Border.BottomBorder,c.Style.Border.BottomBorderColor),
            "left"=>(c.Style.Border.LeftBorder,c.Style.Border.LeftBorderColor),_=>(c.Style.Border.RightBorder,c.Style.Border.RightBorderColor)
        }).Distinct().ToArray();
        if(values.Length!=1) throw new InvalidDataException($"合并格 {edge.RangeAddress} 的同一外边使用了多段不同样式，目前无法准确提入Word。");
        var (style,color)=values[0];
        var (val,size)=style switch
        {
            XLBorderStyleValues.None=>("nil",0u),XLBorderStyleValues.Hair=>("single",2u),XLBorderStyleValues.Thin=>("single",4u),
            XLBorderStyleValues.Medium=>("single",12u),XLBorderStyleValues.Thick=>("single",18u),XLBorderStyleValues.Double=>("double",6u),
            XLBorderStyleValues.Dotted=>("dotted",4u),XLBorderStyleValues.Dashed=>("dashed",8u),XLBorderStyleValues.MediumDashed=>("dashed",12u),
            XLBorderStyleValues.DashDot=>("dotDash",8u),XLBorderStyleValues.MediumDashDot=>("dotDash",12u),
            XLBorderStyleValues.DashDotDot=>("dotDotDash",8u),XLBorderStyleValues.MediumDashDotDot=>("dotDotDash",12u),
            XLBorderStyleValues.SlantDashDot=>("dashDotStroked",12u),_=>throw new InvalidDataException("底稿存在尚未支持的边框样式。")
        };
        W.BorderType result=side switch {"top"=>new W.TopBorder(),"bottom"=>new W.BottomBorder(),"left"=>new W.LeftBorder(),_=>new W.RightBorder()};
        result.Val=new W.BorderValues(val);result.Size=size;result.Color=style==XLBorderStyleValues.None?"auto":ExcelColor(color,source);return result;
    }
    private static XLColor? WordColor(XElement? node,string valueAttribute,string themeAttribute,ThemePart? theme)
    {
        if(node is null) return null;
        var name=(string?)node.Attribute(w+themeAttribute);
        var rgb=name is null?(string?)node.Attribute(w+valueAttribute):ThemeRgb(theme,name);
        if(rgb is null or "auto") return null;
        if(!Regex.IsMatch(rgb,"\\A[0-9a-fA-F]{6}\\z")) throw new InvalidDataException("原表颜色不是有效的RGB值。");
        if(name is not null)
        {
            var tint=(string?)node.Attribute(w+(themeAttribute=="themeFill"?"themeFillTint":"themeTint"));
            var shade=(string?)node.Attribute(w+(themeAttribute=="themeFill"?"themeFillShade":"themeShade"));
            if(shade is not null) rgb=TintRgb(rgb,Convert.ToInt32(shade,16)/255d-1);
            if(tint is not null) rgb=TintRgb(rgb,1-Convert.ToInt32(tint,16)/255d);
        }
        return XLColor.FromHtml("#"+rgb);
    }
    private static string ExcelColor(XLColor color,WorkbookPart source)
    {
        if(!color.HasValue || color.ColorType==XLColorType.Indexed && color.Indexed is 64 or 65) return "auto";
        if(color.ColorType==XLColorType.Theme)
        {
            // 默认字体未显式指定颜色且主题部件缺省时使用自动色；其他主题色必须有真实定义。
            if(source.ThemePart is null && color.ThemeColor is XLThemeColor.Text1 or XLThemeColor.Background1) return "auto";
            return TintRgb(ThemeRgb(source.ThemePart,color.ThemeColor.ToString()),color.ThemeTint);
        }
        if(color.ColorType==XLColorType.Indexed && source.WorkbookStylesPart?.Stylesheet?.Colors?.IndexedColors is { } palette)
        {
            var rgb=palette.ChildElements.ElementAtOrDefault(color.Indexed)?.GetAttribute("rgb","").Value;
            if(rgb is null || !Regex.IsMatch(rgb,"\\A(?:[0-9a-fA-F]{2})?[0-9a-fA-F]{6}\\z")) throw new InvalidDataException("底稿自定义调色板缺少所用颜色。");
            return rgb[^6..].ToUpperInvariant();
        }
        return (color.Color.ToArgb()&0xFFFFFF).ToString("X6",CultureInfo.InvariantCulture);
    }
    private static string ThemeRgb(ThemePart? theme,string name)
    {
        var key=name.ToLowerInvariant() switch
        {
            "background1" or "light1"=>"lt1","text1" or "dark1"=>"dk1","background2" or "light2"=>"lt2","text2" or "dark2"=>"dk2",
            "hyperlink"=>"hlink","followedhyperlink"=>"folHlink",var other=>other
        };
        var node=theme?.Theme?.ThemeElements?.ColorScheme?.ChildElements.FirstOrDefault(n=>n.LocalName==key)?.FirstChild;
        var rgb=node?.GetAttribute(node.LocalName=="sysClr"?"lastClr":"val","").Value;
        if(rgb is null || !Regex.IsMatch(rgb,"\\A[0-9a-fA-F]{6}\\z")) throw new InvalidDataException($"来源文件缺少主题颜色 {name} 的明确RGB值。");
        return rgb.ToUpperInvariant();
    }
    private static string TintRgb(string rgb,double tint)
    {
        if(tint==0) return rgb.ToUpperInvariant();
        if(tint is < -1 or > 1) throw new InvalidDataException("主题颜色亮暗比例无效。");
        // 按OOXML的HLS亮度规则换色，不把tint误当作逐通道混色。
        var color=System.Drawing.ColorTranslator.FromHtml("#"+rgb);double h=color.GetHue()/60d,s=color.GetSaturation(),l=color.GetBrightness();
        l=tint<0?l*(1+tint):l*(1-tint)+tint;
        double c=(1-Math.Abs(2*l-1))*s,x=c*(1-Math.Abs(h%2-1)),m=l-c/2;
        var (red,green,blue)=h switch {<1=>(c,x,0d),<2=>(x,c,0d),<3=>(0d,c,x),<4=>(0d,x,c),<5=>(x,0d,c),_=>(c,0d,x)};
        int Channel(double value)=>(int)Math.Clamp(Math.Round((value+m)*255,MidpointRounding.AwayFromZero),0,255);
        return $"{Channel(red):X2}{Channel(green):X2}{Channel(blue):X2}";
    }
    private static W.Paragraph Paragraph(string text)
    {
        var paragraph=new W.Paragraph();var run=new W.Run();var parts=text.Split('\t');
        for(int i=0;i<parts.Length;i++){if(i>0) run.Append(new W.TabChar());run.Append(new W.Text(parts[i]){Space=SpaceProcessingModeValues.Preserve});}
        paragraph.Append(run);return paragraph;
    }
}
