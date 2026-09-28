#nullable enable
using AnnotationModal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.VariantTypes;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
namespace 附注工具.核心;
public sealed record SourceInfo(string Path, int TableCount, int BodyCount, bool Exists);
public enum SourceLinkKind { All,Tables,Body }
public static class LinkManager
{
    public static SavedFile SetTableLink(string report,string tableId,WorkTableInfo link,string output) => SetTableLinks(report,new Dictionary<string,WorkTableInfo>{[tableId]=link},output);
    public static SavedFile SetTableLinks(string report,IReadOnlyDictionary<string,WorkTableInfo> links,string output,bool clearPhysicalRowMode=false)
    {
        var snapshot=ReportReader.Read(report);
        if(links.Count==0) throw new InvalidDataException("请先选择要设置的表格。");
        var changes=new List<(TableInfo Info,string Description)>();
        foreach(var pair in links)
        {
            var info=snapshot.Tables.SingleOrDefault(t=>t.Id==pair.Key) ?? throw new InvalidDataException("选中的表格位置已失效。");
            var link=pair.Value;
            // 沿用原保存连接的名称、路径与插入点校验；全部检查通过后只保存一次。
            if(string.IsNullOrWhiteSpace(link.LinkName)) throw new InvalidDataException("链接名称不能为空。");
            _=LinkCodec.ResolveSource(report,link);
            var rows=WordTableLayout.Read(info).WordRowCount;
            if(link.autoInsert>rows || link.autoInsert< -rows) throw new InvalidDataException("插入点必须在表格之内。");
            var description=GetDescription(new Table(info.OriginalXml));
            if(!string.IsNullOrEmpty(description) && !LinkCodec.IsLink(description))
                throw new InvalidDataException("此表已有普通说明，不能用连接覆盖它。请保留原说明后再建立连接。");
            var record=ReadTableRecord(description) ?? new JObject();
            if(clearPhysicalRowMode)
            {
                record.Remove("UsePhysicalRows");
                // 仅新提取表使用新网格对应；换源继续保留标记，不改变旧连接的编号含义。
                if(WordTableLayout.Read(info).Cells.Any(c=>c.GridBefore!=0)) record["UseGridColumns"]=true;
                else record.Remove("UseGridColumns");
            }
            if((bool?)record["UsePhysicalRows"]==true && link.autoInsert!=0)
                throw new InvalidDataException("这张新提取表含整行纵向续接，暂不支持自动插行；插入点请保持0。");
            foreach(var property in JObject.FromObject(link).Properties()) record[property.Name]=property.Value;
            changes.Add((info,"<#"+record.ToString(Formatting.None)+"#>"));
        }
        return Mutate(snapshot,output,doc=>
        {
            foreach(var change in changes)
                SetDescription((Table)TableSynchronizer.FindNode(doc.MainDocumentPart!.Document!.Body!,change.Info.NodePath),change.Description);
        });
    }
    public static SavedFile ChangeSource(string report,string oldPath,string newPath,string output,SourceLinkKind kind=SourceLinkKind.All)
    {
        var snapshot=ReportReader.Read(report); oldPath=Path.GetFullPath(oldPath,Folder(report));newPath=Path.GetFullPath(newPath,Folder(report));
        if(!Enum.IsDefined(kind)) throw new InvalidDataException("来源替换范围无效。");
        if(!new[]{".xlsx",".xlsm"}.Contains(Path.GetExtension(newPath),StringComparer.OrdinalIgnoreCase))
            throw new InvalidDataException("请选择 .xlsx 或 .xlsm 底稿。");
        return Mutate(snapshot,output,doc=>
        {
            int changed=0;
            foreach(var table in doc.MainDocumentPart!.Document!.Descendants<Table>().Where(_=>kind!=SourceLinkKind.Body))
            {
                var record=ReadTableRecord(GetDescription(table)); if(record is null) continue;
                var link=record.ToObject<WorkTableInfo>()!;
                if(!SamePath(LinkCodec.ResolveSource(report,link),oldPath)) continue;
                record["SavePath"]=newPath;
                if(!string.IsNullOrEmpty(link.RelativePath)) record["RelativePath"]=Path.GetRelativePath(Folder(report),newPath);
                SetDescription(table,"<#"+record.ToString(Formatting.None)+"#>");changed++;
            }
            foreach(var pair in BodyProperties(doc).Where(_=>kind!=SourceLinkKind.Tables))
            {
                if(!SamePath(ResolveBodySource(report,pair.Record),oldPath)) continue;
                pair.Record["FilePath"]=(bool?)pair.Record["IsRelative"]==true ? Path.GetRelativePath(Folder(report),newPath) : newPath;
                pair.Property.GetFirstChild<VTLPWSTR>()!.Text=pair.Record.ToString(Formatting.None);changed++;
            }
            if(changed==0) throw new InvalidDataException("报告中没有使用所选来源的连接，请重新读取来源清单。");
        });
    }
    public static SavedFile Unlink(string report,IReadOnlyCollection<string> selectedIds,string output)
    {
        var snapshot=ReportReader.Read(report);
        if(selectedIds.Count==0) throw new InvalidDataException("请先选择要断开的连接。");
        return Mutate(snapshot,output,doc=>
        {
            var matched=new HashSet<string>();
            foreach(var info in snapshot.Tables.Where(t=>selectedIds.Contains(t.Id)))
            {
                var table=(Table)TableSynchronizer.FindNode(doc.MainDocumentPart!.Document!.Body!,info.NodePath);
                if(!LinkCodec.IsLink(GetDescription(table))) throw new InvalidDataException("选中的表格没有本软件的连接。");
                table.GetFirstChild<TableProperties>()?.GetFirstChild<TableDescription>()?.Remove(); matched.Add(info.Id);
            }
            foreach(var name in selectedIds.Where(IsBodyLinkName).Distinct(StringComparer.Ordinal))
            {
                var starts=DocumentRoots(doc).SelectMany(root=>root.Descendants<BookmarkStart>()
                    .Where(b=>b.Name?.Value==name).Select(start=>(Root:root,Start:start))).ToArray();
                var properties=doc.CustomFilePropertiesPart?.Properties?.Elements<CustomDocumentProperty>()
                    .Where(p=>p.Name?.Value==name).ToArray() ?? [];
                // 用户明确选中了实际自有书签，记录缺失或损坏也应能断开；无书签的未知属性仍保留。
                if(starts.Length==0 && !properties.Any(p=>ReadBodyRecord(p.GetFirstChild<VTLPWSTR>()?.Text) is not null))continue;
                foreach(var item in starts)
                {
                    foreach(var end in item.Root.Descendants<BookmarkEnd>().Where(b=>b.Id?.Value==item.Start.Id?.Value).ToArray())end.Remove();
                    item.Start.Remove();
                }
                foreach(var property in properties)property.Remove();
                matched.Add(name);
            }
            if(selectedIds.Any(id=>!matched.Contains(id))) throw new InvalidDataException("选中的连接位置已失效，已停止保存。");
        });
    }
    public static SavedFile CleanInvalid(string report,string output) => Mutate(ReportReader.Read(report),output,doc=>
    {
        var names=DocumentRoots(doc).SelectMany(r=>r.Descendants<BookmarkStart>()).Select(b=>b.Name?.Value).ToHashSet();
        // 原ClearInvalidLink只清理无同名书签的属性；来源临时失联并不表示应删连接。
        foreach(var pair in BodyProperties(doc).Where(p=>!names.Contains(p.Property.Name?.Value)).ToArray()) pair.Property.Remove();
    });
    public static IReadOnlyList<SourceInfo> GetSources(string report)
    {
        var snapshot=ReportReader.Read(report);var sources=new List<(string Path,bool IsBody)>();
        foreach(var table in snapshot.Tables.Where(t=>t.Link is not null && t.LinkError is null))
        { try { sources.Add((LinkCodec.ResolveSource(report,table.Link!),false)); } catch(ArgumentException) {} }
        using var doc=WordprocessingDocument.Open(report,false);
        foreach(var pair in BodyProperties(doc))
        { try { sources.Add((ResolveBodySource(report,pair.Record),true)); } catch(ArgumentException) {} }
        return sources.GroupBy(s=>s.Path,StringComparer.OrdinalIgnoreCase)
            .Select(g=>new SourceInfo(g.Key,g.Count(x=>!x.IsBody),g.Count(x=>x.IsBody),File.Exists(g.Key))).ToArray();
    }
    public static bool IsBodyLinkName(string? name) => name is not null && Regex.IsMatch(name,@"\A【B[0-9a-fA-F]{32}B】\z");
    public static JObject? ReadBodyRecord(string? json)
    {
        if(string.IsNullOrWhiteSpace(json)) return null;
        try
        {
            var record=JObject.Parse(json);
            if(string.IsNullOrWhiteSpace((string?)record["FilePath"]) || string.IsNullOrWhiteSpace((string?)record["Name"])
                || !Regex.IsMatch(((string?)record["Address"] ?? "").Trim(),@"\A\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}\z")) return null;
            if(record["IsRelative"] is { Type:not JTokenType.Boolean and not JTokenType.Null }) return null;
            if(record["LinkMode"] is { Type:not JTokenType.Integer and not JTokenType.Null }) return null;
            return record;
        }
        catch(Exception error) when(error is JsonException or ArgumentException or InvalidCastException) { return null; }
    }
    public static string ResolveBodySource(string report,JObject record)
    {
        var path=(string?)record["FilePath"] ?? throw new InvalidDataException("正文连接缺少来源路径。");
        return LinkCodec.ResolveSource(report,new WorkTableInfo {SavePath=path,RelativePath=(bool?)record["IsRelative"]==true ? path : ""});
    }
    internal static IEnumerable<(CustomDocumentProperty Property,JObject Record)> BodyProperties(WordprocessingDocument doc)
    {
        foreach(var property in doc.CustomFilePropertiesPart?.Properties?.Elements<CustomDocumentProperty>() ?? [])
        {
            if(!IsBodyLinkName(property.Name?.Value) || property.GetFirstChild<VTLPWSTR>() is not { } value) continue;
            if(ReadBodyRecord(value.Text) is { } record) yield return (property,record);
        }
    }
    internal static IEnumerable<OpenXmlPartRootElement> DocumentRoots(WordprocessingDocument doc)
    {
        if(doc.MainDocumentPart?.Document is { } main) yield return main;
        foreach(var part in doc.MainDocumentPart?.HeaderParts ?? []) if(part.Header is { } root) yield return root;
        foreach(var part in doc.MainDocumentPart?.FooterParts ?? []) if(part.Footer is { } root) yield return root;
        if(doc.MainDocumentPart?.FootnotesPart?.Footnotes is { } footnotes) yield return footnotes;
        if(doc.MainDocumentPart?.EndnotesPart?.Endnotes is { } endnotes) yield return endnotes;
    }
    internal static void RebaseLinks(WordprocessingDocument doc,string input,string output)
    {
        if(SamePath(Folder(input),Folder(output))) return;
        foreach(var root in DocumentRoots(doc)) RebaseTables(root,input,output);
        foreach(var pair in BodyProperties(doc))
        {
            var json=RebaseBodyRecord(pair.Record,input,output);
            pair.Property.GetFirstChild<VTLPWSTR>()!.Text=json.ToString(Formatting.None);
        }
    }
    public static string RebaseTableXml(string xml,string input,string output)
    {
        if(SamePath(Folder(input),Folder(output))) return xml;
        var table=TableSynchronizer.ParseTableXml(xml);
        var record=ReadTableRecord(GetDescription(table));
        if(record is null || string.IsNullOrEmpty((string?)record["RelativePath"])) return xml;
        RebaseOneTable(table,input,output); return table.OuterXml;
    }
    internal static JObject RebaseBodyRecord(JObject record,string input,string output)
    {
        var result=(JObject)record.DeepClone();
        if((bool?)record["IsRelative"]==true && !SamePath(Folder(input),Folder(output)))
        {
            try { result["FilePath"]=Path.GetRelativePath(Folder(output),ResolveBodySource(input,record)); }
            catch(Exception error) when(error is InvalidDataException or ArgumentException or NotSupportedException)
            { /* 无效的非目标连接原样保留，不能因另存而误处理其他内容。 */ }
        }
        return result;
    }
    private static void RebaseTables(OpenXmlElement root,string input,string output)
    {
        if(SamePath(Folder(input),Folder(output))) return;
        foreach(var table in root.Descendants<Table>()) RebaseOneTable(table,input,output);
    }
    private static void RebaseOneTable(Table table,string input,string output)
    {
        var record=ReadTableRecord(GetDescription(table)); if(record is null) return;
        var link=record.ToObject<WorkTableInfo>()!; if(string.IsNullOrEmpty(link.RelativePath)) return;
        try
        {
            var source=LinkCodec.ResolveSource(input,link);record["SavePath"]=source;record["RelativePath"]=Path.GetRelativePath(Folder(output),source);
            SetDescription(table,"<#"+record.ToString(Formatting.None)+"#>");
        }
        catch(Exception error) when(error is InvalidDataException or ArgumentException or NotSupportedException)
        { /* 目标表仍由入口和预览拒绝；其他损坏连接不能阻断限定章节。 */ }
    }
    private static JObject? ReadTableRecord(string? value)
    {
        if(!LinkCodec.IsLink(value)) return null;
        try
        {
            var link=LinkCodec.Parse(value);
            if(string.IsNullOrWhiteSpace(link.SavePath) && string.IsNullOrWhiteSpace(link.RelativePath)) return null;
            return JObject.Parse(value![2..^2]);
        }
        catch(InvalidDataException) { return null; }
    }
    private static string? GetDescription(Table table) => table.GetFirstChild<TableProperties>()?.GetFirstChild<TableDescription>()?.Val?.Value;
    private static void SetDescription(Table table,string value)
    {
        var properties=table.GetFirstChild<TableProperties>() ?? table.PrependChild(new TableProperties());
        var description=properties.GetFirstChild<TableDescription>();
        if(description is null) properties.AddChild(new TableDescription{Val=value},true); else description.Val=value;
    }
    private static string Folder(string file)=>Path.GetDirectoryName(Path.GetFullPath(file))!;
    private static bool SamePath(string left,string right)=>string.Equals(left,right,StringComparison.OrdinalIgnoreCase);
    private static SavedFile Mutate(ReportInfo snapshot,string output,Action<WordprocessingDocument> edit)
    {
        string? expected=null;string? expectedProperties=null;
        return SafeFile.SaveWordCopy(snapshot.FilePath,snapshot.SourceHash,output,doc=>
        {
            edit(doc);
            var copy=(Document)doc.MainDocumentPart!.Document!.CloneNode(true);
            RebaseTables(copy,snapshot.FilePath,output);expected=copy.OuterXml;
            if(doc.CustomFilePropertiesPart?.Properties is { } properties)
            {
                var copied=(Properties)properties.CloneNode(true);
                if(!SamePath(Folder(snapshot.FilePath),Folder(output)))
                    foreach(var property in copied.Elements<CustomDocumentProperty>())
                        if(IsBodyLinkName(property.Name?.Value) && property.GetFirstChild<VTLPWSTR>() is { } value && ReadBodyRecord(value.Text) is { } record)
                            value.Text=RebaseBodyRecord(record,snapshot.FilePath,output).ToString(Formatting.None);
                expectedProperties=copied.OuterXml;
            }
        },file=>
        {
            using var reread=WordprocessingDocument.Open(file,false);
            if(!TableSynchronizer.SameXml(expected!,reread.MainDocumentPart!.Document!.OuterXml))
                throw new InvalidDataException("连接保存后的报告内容核对不一致。");
            var actual=reread.CustomFilePropertiesPart?.Properties?.OuterXml;
            if(expectedProperties is null ? actual is not null : actual is null || !TableSynchronizer.SameXml(expectedProperties,actual))
                throw new InvalidDataException("连接保存后的自定义属性核对不一致。");
        });
    }
}
