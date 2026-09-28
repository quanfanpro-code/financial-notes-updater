#nullable enable
using AnnotationDal;
using AnnotationModal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.VariantTypes;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
namespace 附注工具.核心;

public sealed record TextSelection(string PartUri, int[] ParagraphPath, int Start, int Length, string ExpectedText, string ReportHash);
public sealed record BodyParagraphInfo(string Id, string PartUri, int[] ParagraphPath, string Text);
public sealed record BodyLinkEntry(string BookmarkName, string PartUri, string Text, DocLinkInfo? Link, string? SourcePath, TextSelection? Selection, string? Error, string? StartParagraphId = null);
public sealed record BodyReportInfo(string FilePath, string SourceHash, IReadOnlyList<BodyParagraphInfo> Paragraphs, IReadOnlyList<BodyLinkEntry> Links);
public sealed record BodyLinkResult(string OutputPath, string? CreatedBookmarkName, int SuccessfulLinks, int FailedLinks, int ChangedLinks, IReadOnlyList<string> Errors, bool OutputVerified, string? BackupPath);

public sealed record BodyLinkChange(BodyLinkEntry Entry,string? NewText,string? Error)
{ public bool CanApply=>Error is null; public bool IsChanged=>CanApply && Entry.Text!=NewText; }
public sealed record BodyRefreshPreview(BodyReportInfo Report,IReadOnlyList<BodyLinkChange> Links,IReadOnlyDictionary<string,string> SourceHashes);

public static class BodyLinkService
{

    public static BodyReportInfo Read(string report)
    {
        var path=Path.GetFullPath(report);var bytes=ModernFile.ReadSnapshot(path,".docx");
        var hash=Convert.ToHexString(SHA256.HashData(bytes));
        using var stream=new MemoryStream(bytes,false);using var doc=WordprocessingDocument.Open(stream,false);
        var paragraphs=new List<BodyParagraphInfo>();var links=new List<BodyLinkEntry>();
        var properties=doc.CustomFilePropertiesPart?.Properties?.Elements<CustomDocumentProperty>().ToArray() ?? [];
        foreach(var root in LinkManager.DocumentRoots(doc))
        {
            var uri=root.OpenXmlPart!.Uri.ToString();
            foreach(var p in root.Descendants<Paragraph>().Where(p=>!p.Ancestors<Paragraph>().Any()))
            {
                var nodePath=NodePath(root,p);paragraphs.Add(new(uri+"#/"+string.Join("/",nodePath),uri,nodePath,Plain(p)));
            }
            foreach(var group in root.Descendants<BookmarkStart>().Where(b=>LinkManager.IsBodyLinkName(b.Name?.Value)).GroupBy(b=>b.Name!.Value!))
            {
                var name=group.Key;var start=group.First();DocLinkInfo? info=null;string? source=null;string? error=null;string text="";TextSelection? selection=null;
                try
                {
                    if(group.Count()!=1) throw new InvalidDataException("存在重复的同名书签。");
                    var matched=properties.Where(p=>p.Name?.Value==name).ToArray();
                    if(matched.Length!=1) throw new InvalidDataException("没有唯一的同名连接记录。");
                    var json=LinkManager.ReadBodyRecord(matched[0].GetFirstChild<VTLPWSTR>()?.Text) ?? throw new InvalidDataException("连接记录不完整或格式错误。");
                    info=Info(json);source=LinkManager.ResolveBodySource(path,json);
                }
                catch(Exception e) when(IsRecordError(e)){error=e.Message;}
                try
                {
                    var range=Locate(root,start);text=range.Text;EnsureEditable(range);
                    if(range.Paragraphs.Count==1)
                    {
                        var p=range.Paragraphs[0];var offset=p.ChildElements.TakeWhile(e=>e!=range.Start).Sum(e=>Plain(e).Length);
                        selection=new(uri,NodePath(root,p),offset,text.Length,text,hash);
                    }
                }
                catch(Exception e) when(IsRecordError(e)){error=Join(error,e.Message);}
                var anchor=start.Ancestors<Paragraph>().LastOrDefault();
                links.Add(new(name,uri,text,info,source,selection,error,anchor is null?null:uri+"#/"+string.Join("/",NodePath(root,anchor))));
            }
        }
        var existing=links.Select(l=>l.BookmarkName).ToHashSet(StringComparer.Ordinal);
        foreach(var property in properties.Where(p=>LinkManager.IsBodyLinkName(p.Name?.Value) && !existing.Contains(p.Name!.Value!)))
            links.Add(new(property.Name!.Value!,"","",null,null,null,"连接记录没有对应书签。"));
        return new(path,hash,paragraphs,links);
    }

    public static IReadOnlyList<string> GetSelectedNames(string report,TextSelection selection)
    {
        var bytes=ModernFile.ReadSnapshot(Path.GetFullPath(report),".docx");
        var hash=Convert.ToHexString(SHA256.HashData(bytes));
        if(!string.Equals(hash,selection.ReportHash,StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("报告在选取文字后已变化，请重新选择。");
        using var stream=new MemoryStream(bytes,false);
        using var doc=WordprocessingDocument.Open(stream,false);
        var root=Root(doc,selection.PartUri);
        var paragraph=TableSynchronizer.FindNode(root,selection.ParagraphPath) as Paragraph
            ?? throw new InvalidDataException("选中的段落位置已失效。");
        var text=Plain(paragraph);
        if(selection.Start<0 || selection.Length<0 || selection.Start>text.Length-selection.Length
            || text.Substring(selection.Start,selection.Length)!=selection.ExpectedText
            || CutsCharacter(text,selection.Start) || CutsCharacter(text,selection.Start+selection.Length))
            throw new InvalidDataException("选中的文字与报告不一致，请重新选择。");
        var names=new List<string>();
        foreach(var start in root.Descendants<BookmarkStart>().Where(b=>LinkManager.IsBodyLinkName(b.Name?.Value)))
        {
            var range=Locate(root,start);
            if(!range.Paragraphs.Contains(paragraph))continue;
            int first=range.Paragraphs[0]==paragraph
                ?paragraph.ChildElements.TakeWhile(e=>e!=range.Start).Sum(e=>Plain(e).Length):0;
            int last=range.Paragraphs[^1]==paragraph
                ?paragraph.ChildElements.TakeWhile(e=>e!=range.End).Sum(e=>Plain(e).Length):text.Length;
            // 独立窗口选区规则：文字必须正长度交叠；光标包含开头、排除末尾；空连接仅接受精确光标。
            // 原版枚举 Selection.Range.Bookmarks，未实测其零光标边界，不将本规则冒称旧Word的完全仿真。
            bool selected=selection.Length>0
                ?Math.Max(selection.Start,first)<Math.Min(selection.Start+selection.Length,last)
                :range.Text.Length==0 ?selection.Start==first :selection.Start>=first && selection.Start<last;
            var name=start.Name!.Value!;
            if(selected && !names.Contains(name,StringComparer.Ordinal))names.Add(name);
        }
        return names;
    }

    public static BodyLinkResult Create(string report,TextSelection selection,DocLinkInfo link,string output)
    {
        var snapshot=Read(report);var json=Validated(link);
        if(!string.Equals(snapshot.SourceHash,selection.ReportHash,StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("报告在选取文字后已变化，请重新选择。");
        var sources=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
        string? inserted=null;
        if(selection.Length==0) inserted=ReadSource(report,json,sources);
        var name="【B"+Guid.NewGuid().ToString("N")+"B】";
        var saved=Save(snapshot,output,doc=>
        {
            var root=Root(doc,selection.PartUri);
            var p=TableSynchronizer.FindNode(root,selection.ParagraphPath) as Paragraph ?? throw new InvalidDataException("选中的段落位置已失效。");
            var text=Plain(p);
            if(selection.Start<0 || selection.Length<0 || selection.Start>text.Length-selection.Length
                || text.Substring(selection.Start,selection.Length)!=selection.ExpectedText
                || CutsCharacter(text,selection.Start) || CutsCharacter(text,selection.Start+selection.Length))
                throw new InvalidDataException("选中的文字与报告不一致，请重新选择。");
            var id=NextBookmarkId(doc);
            var end=InsertMarker(p,selection.Start+selection.Length,new BookmarkEnd{Id=id});
            var start=InsertMarker(p,selection.Start,new BookmarkStart{Name=name,Id=id},end);
            EnsureNoOverlap(root,start,end);
            var range=Locate(root,start);EnsureEditable(range);
            if(inserted is not null) Replace(range,inserted,null);
            var part=doc.CustomFilePropertiesPart ?? doc.AddCustomFilePropertiesPart();
            part.Properties ??=new Properties();
            if(!part.Properties.NamespaceDeclarations.Any(n=>n.Key=="vt"))part.Properties.AddNamespaceDeclaration("vt","http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            var propertyId=Math.Max(1,part.Properties.Elements<CustomDocumentProperty>().Select(p=>p.PropertyId?.Value ?? 1).DefaultIfEmpty(1).Max())+1;
            part.Properties.Append(new CustomDocumentProperty(new VTLPWSTR(json.ToString(Formatting.None))){Name=name,PropertyId=propertyId,FormatId="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"});
        },sources);
        return new(saved.OutputPath,name,1,0,inserted is null?0:1,[],true,saved.BackupPath);
    }

    public static SavedFile Change(string report,string bookmarkName,DocLinkInfo link,string output)
    {
        var snapshot=Read(report);
        if(!LinkManager.IsBodyLinkName(bookmarkName) || snapshot.Links.Count(l=>l.BookmarkName==bookmarkName)!=1)
            throw new InvalidDataException("选中的正文连接不存在或重复。");
        var changes=Validated(link);var sources=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
        _=ReadSource(report,changes,sources);
        return Save(snapshot,output,doc=>
        {
            var property=doc.CustomFilePropertiesPart?.Properties?.Elements<CustomDocumentProperty>().SingleOrDefault(p=>p.Name?.Value==bookmarkName)
                ?? throw new InvalidDataException("选中的正文连接记录不存在。");
            var value=property.GetFirstChild<VTLPWSTR>() ?? throw new InvalidDataException("正文连接记录不是文字属性。");
            var record=JObject.Parse(value.Text);
            foreach(var field in changes.Properties()) record[field.Name]=field.Value;
            value.Text=record.ToString(Formatting.None);
        },sources);
    }



    public static BodyRefreshPreview Preview(string report,IReadOnlyCollection<string>? selectedBookmarkNames=null,CancellationToken token=default)
    {
        token.ThrowIfCancellationRequested();
        var snapshot=Read(report);
        if(selectedBookmarkNames is {Count:0}) throw new InvalidDataException("请先选择要刷新的正文连接。");
        if(selectedBookmarkNames is not null && selectedBookmarkNames.Any(n=>!snapshot.Links.Any(l=>l.BookmarkName==n)))
            throw new InvalidDataException("选中的正文连接已失效，请重新读取报告。");
        var selected=snapshot.Links.Where(l=>selectedBookmarkNames is null || selectedBookmarkNames.Contains(l.BookmarkName)).ToArray();
        if(selected.Length==0) throw new InvalidDataException("报告中没有可刷新的正文连接。");
        var sources=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
        var values=new Dictionary<string,(string? Value,string? Error)>(StringComparer.Ordinal);
        foreach(var entry in selected)
        {
            token.ThrowIfCancellationRequested();
            if(values.ContainsKey(entry.BookmarkName)){values[entry.BookmarkName]=(null,"存在重复的同名书签。");continue;}
            try
            {
                if(entry.Error is not null) throw new InvalidDataException(entry.Error);
                values[entry.BookmarkName]=(ReadSource(report,Validated(entry.Link!),sources),null);
                token.ThrowIfCancellationRequested();
            }
            catch(Exception e) when(IsRecordError(e)){values[entry.BookmarkName]=(null,e.Message);}
        }
        return new(snapshot,selected.DistinctBy(e=>e.BookmarkName).Select(e=>new BodyLinkChange(e,values[e.BookmarkName].Value,values[e.BookmarkName].Error)).ToArray(),sources);
    }

    public static BodyLinkResult Refresh(string report,IReadOnlyCollection<string>? selectedBookmarkNames,string output,CancellationToken token=default)=>Save(Preview(report,selectedBookmarkNames,token),output,token);

    public static BodyLinkResult Save(BodyRefreshPreview preview,string output,CancellationToken token=default)
    {
        token.ThrowIfCancellationRequested();
        var snapshot=preview.Report;var sources=preview.SourceHashes;
        int success=0,failed=0,changed=0;var errors=new List<string>();
        var saved=Save(snapshot,output,doc=>
        {
            foreach(var change in preview.Links)
            {
                token.ThrowIfCancellationRequested();
                var entry=change.Entry;var result=(Value:change.NewText,Error:change.Error);BoundRange? range=null;
                try
                {
                    var root=Root(doc,entry.PartUri);
                    var starts=root.Descendants<BookmarkStart>().Where(b=>b.Name?.Value==entry.BookmarkName).ToArray();
                    if(starts.Length!=1) throw new InvalidDataException("没有唯一的书签位置。");
                    range=Locate(root,starts[0]);
                    if(result.Error is not null) throw new InvalidDataException(result.Error);
                    EnsureEditable(range);
                    var different=range.Text!=result.Value;
                    Replace(range,result.Value!,different && result.Value!.Length>0 ? HighlightColorValues.Yellow : null);
                    success++;if(different)changed++;
                }
                catch(Exception e) when(IsRecordError(e))
                {
                    failed++;errors.Add(entry.BookmarkName+"："+e.Message);
                    if(range is not null) Mark(range,HighlightColorValues.Red);
                }
            }
        },sources,token);
        return new(saved.OutputPath,null,success,failed,changed,errors,true,saved.BackupPath);
    }
    public static SavedFile Unlink(string report,IReadOnlyCollection<string> names,string output)=>LinkManager.Unlink(report,names,output);

    private static string ReadSource(string report,JObject json,Dictionary<string,string> sources)
    {
        var link=Info(json);var file=LinkManager.ResolveBodySource(report,json);ExcelRegion region;ExcelCellInfo cell;
        if(link.LinkMode==DocLinkMode.Name)
        {
            region=ExcelReader.ReadNamedRegion(file,link.Name);
            Remember(region,sources);
            var (row,column)=Offset(link.Address);
            if(row>region.RowCount || column>region.ColumnCount) throw new InvalidDataException("正文地址偏移超出名称区域。");
            cell=region.Cells.Single(c=>c.Row==region.FirstRow+row-1 && c.Column==region.FirstColumn+column-1);
        }
        else
        {
            region=ExcelReader.ReadCellSnapshot(file,link.Name,link.Address.Trim());
            Remember(region,sources);cell=region.Cells.Single();
        }
        if(!cell.CanUpdateReport) throw new InvalidDataException(cell.Error ?? "底稿单元格没有有效结果。");
        // 原 uybuS19sY 的两个取数入口均在写入正文前 Trim。
        return cell.Text.Trim();
    }
    private static void Remember(ExcelRegion region,Dictionary<string,string> sources)
    {
        if(sources.TryGetValue(region.FilePath,out var hash) && hash!=region.SourceHash)
            throw new InvalidDataException("同一底稿在读取过程中发生变化，请重新刷新。");
        sources[region.FilePath]=region.SourceHash;
    }
    private static (int Row,int Column) Offset(string address)
    {
        var match=Regex.Match(address.Trim(),@"\A\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})\z");
        if(!match.Success) throw new InvalidDataException("正文单元格地址无效。");
        int column=0;foreach(var c in match.Groups[1].Value.ToUpperInvariant())column=column*26+c-'A'+1;
        var row=int.Parse(match.Groups[2].Value);
        if(column>16384 || row>1048576) throw new InvalidDataException("正文单元格地址超出底稿范围。");
        return(row,column);
    }
    private static JObject Validated(DocLinkInfo link)
    {
        var json=JObject.FromObject(link);if(LinkManager.ReadBodyRecord(json.ToString(Formatting.None)) is null) throw new InvalidDataException("正文连接需要来源文件、工作表或名称以及单元格地址。");
        _=Offset(link.Address);return json;
    }
    private static DocLinkInfo Info(JObject json)=>new()
    {
        Address=(string)json["Address"]!,Name=(string)json["Name"]!,FilePath=(string)json["FilePath"]!,
        IsRelative=(bool?)json["IsRelative"]==true,Descr=(string?)json["Descr"] ?? "",
        // 原检查对缺省和未定义枚举均退回 Address。
        LinkMode=(int?)json["LinkMode"]==1 ? DocLinkMode.Name : DocLinkMode.Address
    };

    private sealed record BoundRange(BookmarkStart Start,BookmarkEnd End,IReadOnlyList<Paragraph> Paragraphs)
    {
        public IEnumerable<OpenXmlElement> Selected(Paragraph p)
        {
            bool active=p!=Paragraphs[0];
            foreach(var child in p.ChildElements)
            {
                if(child==Start){active=true;continue;}
                if(child==End)yield break;
                if(active && child is not ParagraphProperties)yield return child;
            }
        }
        public string Text=>string.Join("\n",Paragraphs.Select(p=>string.Concat(Selected(p).Select(Plain))));
        public IEnumerable<Run> Runs=>Paragraphs.SelectMany(Selected).SelectMany(n=>n is Run r ? new[]{r} : n.Descendants<Run>());
    }
    private static BoundRange Locate(OpenXmlPartRootElement root,BookmarkStart start)
    {
        var ends=root.Descendants<BookmarkEnd>().Where(e=>e.Id?.Value==start.Id?.Value).ToArray();
        if(ends.Length!=1 || start.Parent is not Paragraph first || ends[0].Parent is not Paragraph last)
            throw new InvalidDataException("书签首尾位置不完整或位于尚未支持的文字结构中。");
        if(first.Parent!=last.Parent)throw new InvalidDataException("书签跨越不同表格或正文容器，已停止替换。");
        var siblings=first.Parent!.ChildElements.ToArray();int from=Array.IndexOf(siblings,first),to=Array.IndexOf(siblings,last);
        if(from>to || siblings.Skip(from).Take(to-from+1).Any(n=>n is not Paragraph))
            throw new InvalidDataException("书签跨越非段落结构，已停止替换。");
        if(first==last && first.ChildElements.ToList().IndexOf(start)>=first.ChildElements.ToList().IndexOf(ends[0]))
            throw new InvalidDataException("书签首尾顺序错误。");
        return new(start,ends[0],siblings.Skip(from).Take(to-from+1).Cast<Paragraph>().ToArray());
    }
    private static void EnsureEditable(BoundRange range)
    {
        foreach(var p in range.Paragraphs)
        {
            if(p.Ancestors<Paragraph>().Any() || p.ParagraphProperties?.GetFirstChild<SectionProperties>() is not null
                || p.Ancestors().Any(a=>a.LocalName is "ins" or "del" or "moveFrom" or "moveTo"))
                throw new InvalidDataException("选区包含文本框、修订或分节结构，已停止替换。");
            foreach(var child in range.Selected(p))
                if(child is not Run run || run.ChildElements.Any(e=>e is not RunProperties and not Text and not TabChar and not Break and not CarriageReturn))
                    throw new InvalidDataException("选区包含域、图片、内容控件或其他书签，已保留原内容。");
        }
    }
    private static void Replace(BoundRange range,string value,HighlightColorValues? color)
    {
        value=value.Replace("\r\n","\n").Replace("\r","\n");
        var oldRuns=range.Runs.ToArray();
        var fragments=oldRuns.Select(r=>(Length:Plain(r).Length,Properties:r.RunProperties is null?null:(RunProperties)r.RunProperties.CloneNode(true))).Where(f=>f.Length>0).ToArray();
        if(fragments.Length==0)
        {
            var neighbor=oldRuns.FirstOrDefault() ?? range.Start.PreviousSibling<Run>() ?? range.End.NextSibling<Run>();
            fragments=[(0,neighbor?.RunProperties is null?null:(RunProperties)neighbor.RunProperties.CloneNode(true))];
        }
        var lines=new List<List<Run>>{new()};int position=0;
        for(int i=0;i<fragments.Length;i++)
        {
            int length=i==fragments.Length-1 ? value.Length-position : Math.Min(fragments[i].Length,value.Length-position);
            if(CutsCharacter(value,position+length))length--;
            var part=value.Substring(position,length);position+=length;
            var pieces=part.Split('\n');
            for(int j=0;j<pieces.Length;j++)
            {
                if(j>0)lines.Add(new());
                if(pieces[j].Length>0)
                {
                    var run=new Run();if(fragments[i].Properties is not null)run.RunProperties=(RunProperties)fragments[i].Properties!.CloneNode(true);
                    if(color is not null) SetColor(run,color.Value);
                    AppendText(run,pieces[j]);lines[^1].Add(run);
                }
            }
        }
        if(value.Length==0)
        {
            var anchor=new Run();if(fragments[0].Properties is not null)anchor.RunProperties=(RunProperties)fragments[0].Properties!.CloneNode(true);
            anchor.Append(new Text(""){Space=SpaceProcessingModeValues.Preserve});lines[0].Add(anchor);
        }
        var first=range.Paragraphs[0];var last=range.Paragraphs[^1];
        var styles=range.Paragraphs.Select(p=>p.ParagraphProperties?.CloneNode(true)).ToArray();
        // 局部正文合段时，未选前缀继续沿用首段排版；整格规则留在表格更新中处理。
        var hasPrefix=first.ChildElements.TakeWhile(e=>e!=range.Start).Any(e=>e is not ParagraphProperties and not BookmarkStart and not BookmarkEnd);
        var suffix=last.ChildElements.SkipWhile(e=>e!=range.End).ToArray();
        foreach(var p in range.Paragraphs)foreach(var node in range.Selected(p).ToArray())node.Remove();
        foreach(var node in suffix)node.Remove();
        foreach(var p in range.Paragraphs.Skip(1))p.Remove();
        Paragraph current=first;
        for(int i=0;i<lines.Count;i++)
        {
            var style=styles[lines.Count==styles.Length?i:hasPrefix?0:styles.Length-1];
            if(i>0)
            {
                var next=new Paragraph();current.InsertAfterSelf(next);current=next;
            }
            if(lines.Count!=1 || range.Paragraphs.Count!=1)
            {
                current.ParagraphProperties?.Remove();if(style is not null)current.PrependChild(style.CloneNode(true));
            }
            foreach(var run in lines[i])current.Append(run);
        }
        foreach(var node in suffix)current.Append(node);
    }
    private static void Mark(BoundRange range,HighlightColorValues color)
    {
        var runs=range.Runs.ToArray();foreach(var run in runs)SetColor(run,color);
        if(runs.Length==0)
        {
            var run=new Run(new Text(""));SetColor(run,color);range.End.InsertBeforeSelf(run);
        }
    }
    private static void SetColor(Run run,HighlightColorValues color)
    {
        run.RunProperties ??=new RunProperties();var highlight=run.RunProperties.GetFirstChild<Highlight>();
        if(highlight is null)run.RunProperties.AddChild(new Highlight{Val=color},true);else highlight.Val=color;
    }
    private static void AppendText(Run run,string value)
    {
        var pieces=value.Split('\t');
        for(int i=0;i<pieces.Length;i++){if(i>0)run.Append(new TabChar());if(pieces[i].Length>0)run.Append(new Text(pieces[i]){Space=SpaceProcessingModeValues.Preserve});}
    }
    private static T InsertMarker<T>(Paragraph p,int offset,T marker,OpenXmlElement? samePositionBefore=null) where T:OpenXmlElement
    {
        int position=0;
        foreach(var node in p.ChildElements.ToArray())
        {
            if(node is ParagraphProperties)continue;
            int length=Plain(node).Length;
            if(position==offset)
            {
                if(node is BookmarkEnd && node!=samePositionBefore)continue;
                if(samePositionBefore is not null && node==samePositionBefore){p.InsertBefore(marker,node);return marker;}
                p.InsertBefore(marker,node);return marker;
            }
            if(position+length>offset)
            {
                if(node is not Run run)throw new InvalidDataException("选中的文字位于内容控件、超链接或其他复杂结构内。");
                var (left,right)=SplitRun(run,offset-position);
                if(left is not null)p.InsertBefore(left,run);
                p.InsertBefore(marker,run);if(right is not null)p.InsertBefore(right,run);run.Remove();return marker;
            }
            position+=length;
        }
        p.Append(marker);return marker;
    }
    private static (Run? Left,Run? Right) SplitRun(Run run,int offset)
    {
        if(run.ChildElements.Any(e=>e is not RunProperties and not Text and not TabChar and not Break and not CarriageReturn))
            throw new InvalidDataException("选中的文字与域或图片共用文字节点，已停止拆分。");
        var left=new Run();var right=new Run();
        if(run.RunProperties is not null){left.RunProperties=(RunProperties)run.RunProperties.CloneNode(true);right.RunProperties=(RunProperties)run.RunProperties.CloneNode(true);}
        int position=0;
        foreach(var child in run.ChildElements.Where(c=>c is not RunProperties))
        {
            int length=Plain(child).Length;
            if(position+length<=offset)left.Append(child.CloneNode(true));
            else if(position>=offset)right.Append(child.CloneNode(true));
            else if(child is Text t)
            {
                int split=offset-position;
                left.Append(new Text(t.Text[..split]){Space=SpaceProcessingModeValues.Preserve});right.Append(new Text(t.Text[split..]){Space=SpaceProcessingModeValues.Preserve});
            }
            else throw new InvalidDataException("不能从换行或制表符中间拆分。");
            position+=length;
        }
        return(left.ChildElements.Any(e=>e is not RunProperties)?left:null,right.ChildElements.Any(e=>e is not RunProperties)?right:null);
    }
    private static void EnsureNoOverlap(OpenXmlPartRootElement root,BookmarkStart start,BookmarkEnd end)
    {
        var elements=root.Descendants().ToArray();int a=Array.IndexOf(elements,start),b=Array.IndexOf(elements,end);
        foreach(var other in root.Descendants<BookmarkStart>().Where(s=>s!=start))
        {
            var matching=root.Descendants<BookmarkEnd>().Where(e=>e.Id?.Value==other.Id?.Value).ToArray();
            if(matching.Length!=1)continue;
            int c=Array.IndexOf(elements,other),d=Array.IndexOf(elements,matching[0]);
            if(c<b && d>a)throw new InvalidDataException("选区与已有书签重叠，请选择独立的文字。");
        }
    }
    private static string Plain(OpenXmlElement node)
    {
        if(node is Text t)return t.Text;
        if(node is TabChar)return "\t";if(node is Break or CarriageReturn)return "\n";
        if(node is ParagraphProperties or RunProperties)return "";
        return string.Concat(node.ChildElements.Where(e=>e is not Paragraph).Select(Plain));
    }
    private static bool CutsCharacter(string text,int index)=>index>0 && index<text.Length && char.IsHighSurrogate(text[index-1]) && char.IsLowSurrogate(text[index]);
    private static int[] NodePath(OpenXmlElement root,OpenXmlElement node)
    {
        var path=new List<int>();
        while(node!=root){var parent=node.Parent ?? throw new InvalidDataException("段落位置无效。");path.Add(parent.ChildElements.ToList().IndexOf(node));node=parent;}
        path.Reverse();return path.ToArray();
    }
    private static OpenXmlPartRootElement Root(WordprocessingDocument doc,string uri)=>LinkManager.DocumentRoots(doc).SingleOrDefault(r=>r.OpenXmlPart!.Uri.ToString()==uri)
        ?? throw new InvalidDataException("连接所在的报告部件已失效。");
    private static string NextBookmarkId(WordprocessingDocument doc)
    {
        var ids=LinkManager.DocumentRoots(doc).SelectMany(r=>r.Descendants<BookmarkStart>()).Select(b=>b.Id?.Value).ToHashSet();
        int id=0;while(ids.Contains(id.ToString()))id++;return id.ToString();
    }
    private static bool IsRecordError(Exception e)=>e is InvalidDataException or NotSupportedException or IOException or UnauthorizedAccessException or ArgumentException or InvalidOperationException or FormatException or JsonException or OverflowException;
    private static string Join(string? a,string b)=>a is null?b:a+"；"+b;

    private static SavedFile Save(BodyReportInfo snapshot,string output,Action<WordprocessingDocument> edit,IReadOnlyDictionary<string,string> sources,CancellationToken token=default)
    {
        Dictionary<string,string> expected=[];string? expectedProperties=null;
        return SafeFile.SaveWordCopy(snapshot.FilePath,snapshot.SourceHash,output,doc=>
        {
            edit(doc);
            foreach(var root in LinkManager.DocumentRoots(doc))
            {
                var copy=root.CloneNode(true);
                foreach(var table in copy.Descendants<Table>().Reverse().ToArray())
                    table.Parent!.ReplaceChild(new Table(LinkManager.RebaseTableXml(table.OuterXml,snapshot.FilePath,output)),table);
                expected[root.OpenXmlPart!.Uri.ToString()]=copy.OuterXml;
            }
            if(doc.CustomFilePropertiesPart?.Properties is { } properties)
            {
                var copy=(Properties)properties.CloneNode(true);
                if(!string.Equals(Path.GetDirectoryName(Path.GetFullPath(snapshot.FilePath)),Path.GetDirectoryName(Path.GetFullPath(output)),StringComparison.OrdinalIgnoreCase))
                foreach(var property in copy.Elements<CustomDocumentProperty>())
                    if(LinkManager.IsBodyLinkName(property.Name?.Value) && property.GetFirstChild<VTLPWSTR>() is { } value && LinkManager.ReadBodyRecord(value.Text) is { } record)
                        value.Text=LinkManager.RebaseBodyRecord(record,snapshot.FilePath,output).ToString(Formatting.None);
                expectedProperties=copy.OuterXml;
            }
        },file=>
        {
            using var actual=WordprocessingDocument.Open(file,false);
            var roots=LinkManager.DocumentRoots(actual).ToDictionary(r=>r.OpenXmlPart!.Uri.ToString(),r=>r.OuterXml);
            if(expected.Count!=roots.Count || expected.Any(p=>!roots.TryGetValue(p.Key,out var xml) || !TableSynchronizer.SameXml(p.Value,xml)))
                throw new InvalidDataException("正文连接保存后的内容核对不一致。");
            var props=actual.CustomFilePropertiesPart?.Properties?.OuterXml;
            if(expectedProperties is null?props is not null:props is null || !TableSynchronizer.SameXml(expectedProperties,props))
                throw new InvalidDataException("正文连接保存后的属性核对不一致。");
        },token:token,sources:sources);
    }
}


