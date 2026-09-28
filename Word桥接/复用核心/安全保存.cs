#nullable enable
using DocumentFormat.OpenXml.Packaging;
using System.Security.Cryptography;
namespace 附注工具.核心;

public sealed record SavedFile(string OutputPath, string? BackupPath);
public static class SafeFile
{
    public static SavedFile SaveExcelCopy(string file,string expectedHash,string output,Action<MemoryStream> edit,Action<string> validator,CancellationToken token=default)
    {
        var extension=Path.GetExtension(file);
        if ((!extension.Equals(".xlsx",StringComparison.OrdinalIgnoreCase) && !extension.Equals(".xlsm",StringComparison.OrdinalIgnoreCase))
            || !extension.Equals(Path.GetExtension(output),StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("底稿副本应保留原来的 .xlsx 或 .xlsm 格式。");
        return SaveCopy(file,expectedHash,output,edit,validator,token,null);
    }
    public static string Hash(string file) => Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(file)));
    public static SavedFile SaveWordCopy(string report, string expectedHash, string output,
        Action<WordprocessingDocument> edit, Action<string> validator,
        CancellationToken token = default, IReadOnlyDictionary<string, string>? sources = null)
    {
        if (!string.Equals(Path.GetExtension(output), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("报告副本必须保存为 .docx 文件。");
        return SaveCopy(report,expectedHash,output,memory =>
        {
            using var document=WordprocessingDocument.Open(memory,true);
            edit(document);
            LinkManager.RebaseLinks(document,report,output);
        },validator,token,sources);
    }
    private static SavedFile SaveCopy(string report,string expectedHash,string output,Action<MemoryStream> edit,
        Action<string> validator,CancellationToken token,IReadOnlyDictionary<string,string>? sources)
    {
        report=Path.GetFullPath(report); output=Path.GetFullPath(output);
        if(string.Equals(report,output,StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("请另存为副本，不能覆盖本次输入文件。");
        token.ThrowIfCancellationRequested();
        var hashes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { [report] = expectedHash };
        foreach (var source in sources ?? new Dictionary<string, string>()) hashes[Path.GetFullPath(source.Key)] = source.Value;
        var locked = new Dictionary<string, FileStream>(StringComparer.OrdinalIgnoreCase);
        try
        {
            // 持有只读句柄直到提交，阻止其他程序在核对后改写或换掉报告及底稿。
            foreach (var item in hashes)
            {
                var stream = new FileStream(item.Key, FileMode.Open, FileAccess.Read, FileShare.Read);
                locked.Add(item.Key, stream);
                if (!string.Equals(Convert.ToHexString(SHA256.HashData(stream)), item.Value, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException($"文件在预览后发生变化，请重新预览：{item.Key}");
                stream.Position = 0;
            }
            using var memory = new MemoryStream();
            locked[report].CopyTo(memory);
            memory.Position = 0;
            edit(memory);
            token.ThrowIfCancellationRequested();
            var folder = Path.GetDirectoryName(output)!;
            Directory.CreateDirectory(folder);
            var staged = Path.Combine(folder, Path.GetFileNameWithoutExtension(output) + ".待校验_" + Guid.NewGuid().ToString("N") + Path.GetExtension(output));
            using (var file = new FileStream(staged, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                memory.Position = 0;
                memory.CopyTo(file);
                file.Flush(true);
            }
            validator(staged);
            token.ThrowIfCancellationRequested();
            var stagedHash = Hash(staged);
            var backup = BackupExisting(output);
            token.ThrowIfCancellationRequested();
            File.Move(staged, output, true);
            if (Hash(output) != stagedHash) throw new InvalidDataException("保存后的文件与已校验副本不一致。");
            return new SavedFile(output, backup);
        }
        finally
        {
            foreach (var stream in locked.Values) stream.Dispose();
            // 失败时保留待校验文件，便于排查；不自动删除用户目录中的文件。
        }
    }
    public static string? BackupExisting(string file)
    {
        if (!File.Exists(file)) return null;
        var folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "BackUp",
            "独立附注工具_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + "_" + Guid.NewGuid().ToString("N")[..6]);
        Directory.CreateDirectory(folder);
        var backup = Path.Combine(folder, Path.GetFileName(file));
        File.Copy(file, backup, false);
        if (Hash(file) != Hash(backup)) throw new IOException("已有文件备份校验失败，已停止写入。");
        return backup;
    }
    internal static void CommitPrepared(params (string Staged,string Output)[] files)
    {
        var prepared=files.Select(f=>(Staged:Path.GetFullPath(f.Staged),Output:Path.GetFullPath(f.Output))).ToArray();
        if(prepared.Select(f=>f.Output).Distinct(StringComparer.OrdinalIgnoreCase).Count()!=prepared.Length)
            throw new InvalidDataException("两个输出文件不能使用同一路径。");
        var held=new List<FileStream>();var backups=new Dictionary<string,string?>();var hashes=new Dictionary<string,string>();var written=new List<string>();
        try
        {
            // 两份文件均已校验；先检查两个目标均可替换，再备份，不能先覆盖Excel才发现报告被占用。
            foreach(var file in prepared)
            {
                if(!string.Equals(Path.GetDirectoryName(file.Staged),Path.GetDirectoryName(file.Output),StringComparison.OrdinalIgnoreCase)
                    || string.Equals(file.Staged,file.Output,StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("待提交副本必须位于对应输出文件夹，并使用独立名称。");
                hashes[file.Output]=Hash(file.Staged);
                if(File.Exists(file.Output))
                {
                    using(var probe=new FileStream(file.Output,FileMode.Open,FileAccess.ReadWrite,FileShare.None)) {}
                    held.Add(new FileStream(file.Output,FileMode.Open,FileAccess.Read,FileShare.Read|FileShare.Delete));
                }
            }
            foreach(var file in prepared) backups[file.Output]=BackupExisting(file.Output);
            foreach(var file in prepared)
            {
                File.Move(file.Staged,file.Output,true);written.Add(file.Output);
                if(Hash(file.Output)!=hashes[file.Output]) throw new InvalidDataException("输出文件与已校验副本不一致。");
            }
        }
        catch(Exception error) when(written.Count>0)
        {
            foreach(var stream in held) stream.Dispose();held.Clear();
            var restored=new List<string>();
            foreach(var path in written)
                if(backups[path] is { } backup)
                {
                    try
                    {
                        BackupExisting(path);File.Copy(backup,path,true);
                        if(Hash(path)==Hash(backup)) restored.Add(path);
                    }
                    catch(IOException) { }
                }
            var retained=written.Except(restored).ToArray();
            throw new IOException("两份输出未全部提交。"+(retained.Length>0 ? "已生成、需保留核查的文件："+string.Join("；",retained) : "已恢复先前被替换的输出。"),error);
        }
        finally { foreach(var stream in held) stream.Dispose(); }
    }
}
