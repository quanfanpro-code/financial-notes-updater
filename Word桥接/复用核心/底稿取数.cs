#nullable enable
using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Xml;
using AnnotationModal;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace 附注工具.核心;

public sealed record MergeRegion(int FirstRow, int LastRow, int FirstColumn, int LastColumn);
public enum ExcelEditKind { Text, Number, Formula, DateTime }
public sealed record ExcelEdit(string Sheet, string Address, ExcelEditKind Kind, string Value);
public sealed record ExcelWorksheetInfo(string Name, string UsedAddress);
public sealed record ExcelNamedRegionInfo(string Name, string DisplayName, string? ScopeSheet, string Reference, bool IsVisible = true);
public sealed record ExcelWorkbookInfo(string FilePath, string SourceHash,
    IReadOnlyList<ExcelWorksheetInfo> Worksheets, IReadOnlyList<ExcelNamedRegionInfo> NamedRegions);

public sealed record ExcelCellInfo(string Address, int Row, int Column, string Text,
    bool IsFormula, bool UsesSavedResult, string? Error, string? Formula = null,
    ExcelEditKind? EditableKind = null, string? EditableValue = null)
{
    public bool CanUpdateReport => Error is null;
}

public sealed record ExcelRegion(string FilePath, string SourceHash, string Sheet, string Address,
    int FirstRow, int LastRow, int FirstColumn, int LastColumn,
    IReadOnlyList<ExcelCellInfo> Cells, IReadOnlyList<MergeRegion> Merges)
{
    public int RowCount => LastRow - FirstRow + 1;
    public int ColumnCount => LastColumn - FirstColumn + 1;
}

public static class ExcelReader
{
    private const string EditableDateFormat = "yyyy-MM-ddTHH:mm:ss.fffffff";
    public static SavedFile SaveEdits(string file, IReadOnlyList<ExcelEdit> changes, string output, string? expectedHash = null)
    {
        var path = Path.GetFullPath(file);
        var bytes = ModernFile.ReadSnapshot(path, ".xlsx", ".xlsm");
        var sourceHash = Convert.ToHexString(SHA256.HashData(bytes));
        if (expectedHash is not null && !string.Equals(expectedHash, sourceHash, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("底稿在预览后发生变化，请重新读取后再编辑。");
        var macroIdentity = MacroIdentity(bytes);
        var calculated = new List<(string Sheet, string Address, XLCellValue Value)>();
        return SafeFile.SaveExcelCopy(path, sourceHash, output, memory =>
        {
            using var workbook = new XLWorkbook(memory);
            var edits = new List<(IXLCell Cell, ExcelEdit Edit, XLCellValue Value)>();
            var changed = new HashSet<(string Sheet, string Address)>();
            foreach (var edit in changes)
            {
                if (!IsAddress(edit.Address, false)) throw new InvalidDataException($"单元格地址无效：{edit.Address}。");
                var sheet = workbook.Worksheets.FirstOrDefault(s => string.Equals(s.Name, edit.Sheet, StringComparison.OrdinalIgnoreCase))
                    ?? throw new InvalidDataException($"底稿中没有工作表：{edit.Sheet}。");
                IXLCell cell;
                try { cell = sheet.Cell(edit.Address); }
                catch (ArgumentException e) { throw new InvalidDataException($"单元格地址无效：{edit.Address}。", e); }
                var address = cell.Address.ToStringRelative();
                if (!changed.Add((sheet.Name, address))) throw new InvalidDataException($"同一单元格存在重复编辑：{sheet.Name}!{address}。");
                if (cell.IsMerged() && cell.MergedRange().RangeAddress.FirstAddress.ToStringRelative() != address)
                    throw new InvalidOperationException($"{sheet.Name}!{address} 是合并格的一部分，请编辑左上角 {cell.MergedRange().RangeAddress.FirstAddress.ToStringRelative()}。");
                XLCellValue value = edit.Value;
                if (edit.Kind == ExcelEditKind.Number)
                {
                    if (!double.TryParse(edit.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || !double.IsFinite(number))
                        throw new InvalidDataException($"{sheet.Name}!{address} 不是有效数值。");
                    value = number;
                }
                else if (edit.Kind == ExcelEditKind.Formula)
                {
                    if (string.IsNullOrWhiteSpace(edit.Value.TrimStart('='))) throw new InvalidDataException("公式不能为空，请用文字类型清空单元格。");
                }
                else if (edit.Kind == ExcelEditKind.DateTime)
                {
                    if (!DateTime.TryParseExact(edit.Value, EditableDateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                        throw new InvalidDataException($"{sheet.Name}!{address} 不是有效日期，请通过日期控件选择。");
                    value = date;
                }
                else if (edit.Kind != ExcelEditKind.Text) throw new InvalidDataException("不支持的单元格编辑类型。");
                edits.Add((cell, edit, value));
            }
            // 沿用任务2对原XML的核对：修复ISO日期读取偏移；未识别错误只有被明确编辑修正后才能继续。
            foreach (var sheet in workbook.Worksheets)
                foreach (var saved in ReadSavedCells(bytes, sheet.Name, workbook, false, out _))
                {
                    var cell = sheet.Cell(saved.Key);
                    if (cell.HasFormula || changed.Contains((sheet.Name, saved.Key))) continue;
                    if (saved.Value.DataType?.Value == CellValues.Error && !cell.CachedValue.IsError)
                        throw new InvalidDataException($"工作表“{sheet.Name}”的 {saved.Key} 包含未识别错误 {saved.Value.CellValue?.Text}，请先修正该格。");
                    if (saved.Value.DataType?.Value == CellValues.Date && saved.Value.CellValue is not null)
                    {
                        var style = cell.Style;
                        cell.Value = XmlConvert.ToDateTime(saved.Value.CellValue.Text, XmlDateTimeSerializationMode.RoundtripKind);
                        cell.Style = style;
                    }
                }
            foreach (var item in edits)
            {
                var style = item.Cell.Style;
                var hadDateFormat = item.Cell.CachedValue.IsDateTime || OriginalDateFormat(style.NumberFormat.NumberFormatId) is not null
                    || style.NumberFormat.Format == OriginalLongDateFormat;
                if (item.Edit.Kind == ExcelEditKind.Formula) item.Cell.FormulaA1 = item.Edit.Value;
                else
                {
                    // 库会吃掉文字首个单引号；显式文字类型保留字面，因此仅在此处补一个供库消费。
                    item.Cell.Value = item.Edit.Kind == ExcelEditKind.Text && item.Edit.Value.StartsWith('\'')
                        ? "'" + item.Edit.Value : item.Value;
                }
                item.Cell.Style = style;
                // 已有日期继续使用原格式；用户明确把普通格改为日期时给出日期显示，不能留下数字序号。
                if (item.Edit.Kind == ExcelEditKind.DateTime && !hadDateFormat)
                    item.Cell.Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss";
            }
            var formulas = workbook.Worksheets.SelectMany(s => s.CellsUsed(XLCellsUsedOptions.Contents)).Where(c => c.HasFormula).ToArray();
            if (formulas.Length > 0 && workbook.Use1904DateSystem)
                throw new InvalidDataException("1904日期系统的公式重新计算暂不支持，已停止保存。");
            try
            {
                workbook.RecalculateAllFormulas();
                foreach (var cell in formulas)
                {
                    var value = cell.CachedValue;
                    if (cell.NeedsRecalculation || value.IsError || value.IsBlank)
                        throw new InvalidDataException($"{cell.Worksheet.Name}!{cell.Address} 没有有效的新公式结果：{value}。");
                    calculated.Add((cell.Worksheet.Name, cell.Address.ToStringRelative(), value));
                }
            }
            catch (Exception e) { throw new InvalidDataException($"公式重新计算失败，未保存副本：{e.Message}", e); }
            // 先显式计算和核对，避免保存选项吞掉计算异常后留下缺缓存的文件。
            workbook.Save(new SaveOptions { EvaluateFormulasBeforeSaving = true });
        }, staged =>
        {
            var resultBytes = ModernFile.ReadSnapshot(staged, ".xlsx", ".xlsm");
            var actualMacros = MacroIdentity(resultBytes);
            if (macroIdentity.Count != actualMacros.Count || macroIdentity.Any(p => !actualMacros.TryGetValue(p.Key, out var hash) || hash != p.Value))
                throw new InvalidDataException("保存后的宏部件或宏对应关系发生变化，已停止输出。");
            using var stream = new MemoryStream(resultBytes, false);
            using var result = new XLWorkbook(stream);
            foreach (var group in calculated.GroupBy(c => c.Sheet))
            {
                var saved = ReadSavedCells(resultBytes, group.Key, result, false, out _);
                foreach (var expected in group)
                    if (!saved.TryGetValue(expected.Address, out var raw) || !SavedResultMatches(raw, expected.Value))
                        throw new InvalidDataException($"{expected.Sheet}!{expected.Address} 保存后的公式结果不一致，已停止输出。");
            }
        });
    }

    private static bool SavedResultMatches(Cell cell, XLCellValue expected)
    {
        if (cell.CellValue is null) return false;
        var text = cell.CellValue.Text;
        if (expected.IsText) return cell.DataType?.Value == CellValues.String && text == expected.GetText();
        if (expected.IsBoolean) return cell.DataType?.Value == CellValues.Boolean && text == (expected.GetBoolean() ? "1" : "0");
        if (cell.DataType is not null && cell.DataType.Value != CellValues.Number) return false;
        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)) return false;
        if (expected.IsNumber) return number == expected.GetNumber();
        if (expected.IsTimeSpan) return number == expected.GetTimeSpan().TotalDays;
        return expected.IsDateTime && ((XLCellValue)number).TryConvert(out DateTime date) && date == expected.GetDateTime();
    }

    private static Dictionary<string, string> MacroIdentity(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, false);
        using var document = SpreadsheetDocument.Open(stream, false);
        var part = document.WorkbookPart ?? throw new InvalidDataException("底稿中没有工作簿。");
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        if (part.VbaProjectPart is null) return result;
        void Add(OpenXmlPart item)
        {
            using var content = item.GetStream(FileMode.Open, FileAccess.Read);
            result.Add(item.Uri.ToString(), Convert.ToHexString(SHA256.HashData(content)));
            foreach (var child in item.Parts) Add(child.OpenXmlPart);
        }
        Add(part.VbaProjectPart);
        result.Add("工作簿代码名称", part.Workbook?.WorkbookProperties?.CodeName?.Value ?? "");
        foreach (var sheet in part.WorksheetParts)
            result.Add("工作表代码名称:" + sheet.Uri, sheet.Worksheet?.SheetProperties?.CodeName?.Value ?? "");
        return result;
    }
    public static ExcelWorkbookInfo GetWorkbookInfo(string file)
    {
        var path = Path.GetFullPath(file);
        var bytes = ModernFile.ReadSnapshot(path, ".xlsx", ".xlsm");
        using var stream = new MemoryStream(bytes, false);
        using var workbook = new XLWorkbook(stream);
        using var xmlStream = new MemoryStream(bytes, false);
        using var document = SpreadsheetDocument.Open(xmlStream, false);
        var book = document.WorkbookPart?.Workbook ?? throw new InvalidDataException("底稿中没有工作簿。");
        var sheets = book.Sheets!.Elements<Sheet>().ToArray();
        var names = (book.DefinedNames?.Elements<DefinedName>() ?? []).Select(n =>
        {
            var scope = n.LocalSheetId is not null && n.LocalSheetId.Value < sheets.Length
                ? sheets[(int)n.LocalSheetId.Value].Name!.Value : null;
            var name = n.Name!.Value!;
            return new ExcelNamedRegionInfo(name, scope is null ? name : "'" + scope.Replace("'", "''") + "'!" + name,
                scope, n.Text, !(n.Hidden?.Value ?? false));
        }).ToArray();
        return new ExcelWorkbookInfo(path, Convert.ToHexString(SHA256.HashData(bytes)), workbook.Worksheets.Select(s =>
        {
            var used = s.RangeUsed();
            return new ExcelWorksheetInfo(s.Name, used is null ? "A1" : RangeAddressText(used.RangeAddress));
        }).ToArray(), names);
    }

    public static ExcelRegion ReadRegion(string file, string sheet, string address, bool recalculate = false, DalSetting? setting = null)
    {
        var path = Path.GetFullPath(file);
        var bytes = ModernFile.ReadSnapshot(path, ".xlsx", ".xlsm");
        if (!IsAddress(address, true)) throw new InvalidDataException($"区域地址无效：{address}。");
        return ReadRegion(path, bytes, sheet, address, recalculate, setting ?? new DalSetting());
    }

    public static ExcelRegion ReadCellSnapshot(string file, string sheet, string address, bool recalculate = false, DalSetting? setting = null)
    {
        if (!IsAddress(address, false)) throw new InvalidDataException($"单元格地址无效：{address}。");
        return ReadRegion(file, sheet, address, recalculate, setting);
    }
    public static ExcelRegion ReadNamedRegion(string file, string name, bool recalculate = false, DalSetting? setting = null)
    {
        var path = Path.GetFullPath(file);
        var bytes = ModernFile.ReadSnapshot(path, ".xlsx", ".xlsm");
        using var stream = new MemoryStream(bytes, false);
        using var document = SpreadsheetDocument.Open(stream, false);
        var part = document.WorkbookPart ?? throw new InvalidDataException("底稿中没有工作簿。");
        var sourceBook = part.Workbook ?? throw new InvalidDataException("底稿中没有工作簿内容。");
        var sheets = sourceBook.Sheets?.Elements<Sheet>().ToArray() ?? [];
        var requested = SplitSheetName(name);
        var matches = (sourceBook.DefinedNames?.Elements<DefinedName>() ?? [])
            .Where(n => string.Equals(n.Name?.Value, requested.Name, StringComparison.OrdinalIgnoreCase)
                && (requested.Sheet is null || n.LocalSheetId is not null
                    && n.LocalSheetId.Value < sheets.Length
                    && string.Equals(sheets[(int)n.LocalSheetId.Value].Name?.Value, requested.Sheet, StringComparison.OrdinalIgnoreCase)))
            .ToArray();
        if (matches.Length == 0) throw new InvalidDataException($"没有找到命名区域：{name}。");
        // 修复原 XSSFWorkbook.GetName 只取第一个同名项的错配风险，不猜测要用哪张工作表。
        if (matches.Length != 1) throw new InvalidDataException($"命名区域“{name}”存在同名歧义，请明确工作表。");
        var reference = SplitSheetName(matches[0].Text.TrimStart('='));
        if (reference.Sheet is null || reference.Sheet.IndexOfAny(['[', ']']) >= 0 || !IsAddress(reference.Name, true))
            throw new InvalidDataException($"命名区域“{name}”必须指向本文件内一个矩形区域；目前不支持命名公式、多区域或外部引用。");
        return ReadRegion(path, bytes, reference.Sheet, reference.Name, recalculate, setting ?? new DalSetting());
    }

    public static ExcelCellInfo ReadCell(string file, string sheet, string address, bool recalculate = false, DalSetting? setting = null)
        => ReadCellSnapshot(file, sheet, address, recalculate, setting).Cells.Single();

    private static ExcelRegion ReadRegion(string path, byte[] bytes, string sheetName, string address, bool recalculate, DalSetting setting)
    {
        using var stream = new MemoryStream(bytes, false);
        using var workbook = new XLWorkbook(stream);
        var sheet = workbook.Worksheets.FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidDataException($"底稿中没有工作表：{sheetName}。");
        IXLRange range;
        try { range = sheet.Range(address); }
        catch (ArgumentException e) { throw new InvalidDataException($"区域地址无效：{address}。", e); }
        var first = range.RangeAddress.FirstAddress;
        var last = range.RangeAddress.LastAddress;
        var caches = ReadSavedCells(bytes, sheet.Name, workbook, recalculate, out var calculationError);
        // 显式重算先使所有公式缓存失效；只实际计算所取单元格及其依赖，避免沿用依赖格的旧结果。
        if (recalculate && calculationError is null)
            foreach (var formulaCell in workbook.Worksheets.SelectMany(s => s.CellsUsed(XLCellsUsedOptions.Contents)).Where(c => c.HasFormula))
                formulaCell.InvalidateFormula();
        using var formatBook = new XLWorkbook();
        var formatCell = formatBook.AddWorksheet("显示值").Cell(1, 1);
        var cells = new List<ExcelCellInfo>();
        for (var row = first.RowNumber; row <= last.RowNumber; row++)
            for (var column = first.ColumnNumber; column <= last.ColumnNumber; column++)
            {
                var cell = sheet.Cell(row, column);
                caches.TryGetValue(cell.Address.ToStringRelative(), out var saved);
                cells.Add(ReadValue(cell, saved, formatCell, recalculate, calculationError, setting));
            }
        var merges = sheet.MergedRanges.Select(m => new MergeRegion(m.RangeAddress.FirstAddress.RowNumber,
            m.RangeAddress.LastAddress.RowNumber, m.RangeAddress.FirstAddress.ColumnNumber, m.RangeAddress.LastAddress.ColumnNumber))
            .Where(m => m.FirstRow <= last.RowNumber && m.LastRow >= first.RowNumber
                && m.FirstColumn <= last.ColumnNumber && m.LastColumn >= first.ColumnNumber).ToArray();
        return new ExcelRegion(path, Convert.ToHexString(SHA256.HashData(bytes)), sheet.Name,
            RangeAddressText(range.RangeAddress), first.RowNumber, last.RowNumber, first.ColumnNumber, last.ColumnNumber, cells, merges);
    }

    private static ExcelCellInfo ReadValue(IXLCell cell, Cell? saved, IXLCell formatCell, bool recalculate,
        string? calculationError, DalSetting setting)
    {
        var address = cell.Address.ToStringRelative();
        var isFormula = cell.HasFormula;
        var savedResult = isFormula && !recalculate;
        ExcelEditKind? editableKind = isFormula ? ExcelEditKind.Formula : null;
        string? editableValue = isFormula ? cell.FormulaA1 : null;
        ExcelCellInfo Result(string text, string? error = null, bool? usedSaved = null) =>
            new(address, cell.Address.RowNumber, cell.Address.ColumnNumber, text, isFormula, usedSaved ?? savedResult, error,
                isFormula ? cell.FormulaA1 : null, error is null || isFormula ? editableKind : null,
                error is null || isFormula ? editableValue : null);
        if (isFormula && recalculate && calculationError is not null)
            return Result("", calculationError, false);
        // 库的日期函数返回1900序号，而1904底稿中的数字可能仍是1904序号；在共同入口拒绝混算。
        if (isFormula && recalculate && cell.Worksheet.Workbook.Use1904DateSystem)
            return Result("", "1904日期系统的公式重新计算暂不支持，请读取文件已保存的结果。", false);
        // 保留原 VHIvjrMje 的规则：默认读取 CachedFormulaResultType 对应的已保存值，不求值。
        // 单独检查本次文件快照中的 <v>；数值空缓存不能当 0，字符串空缓存却是有效结果。
        if (savedResult && (saved?.CellValue is null || saved.CellValue.Text.Length == 0 && saved.DataType?.Value != CellValues.String))
            return Result("", "公式没有有效的已保存结果，不能据此更新报告。", false);
        // 新版错误（如 #SPILL!）可能被读取库转为空白，必须先尊重原 XML 的错误类型。
        // 公式明确重算时不使用旧错误，常量错误则始终保留原文。
        if ((!isFormula || !recalculate) && saved?.DataType?.Value == CellValues.Error)
            return Result(saved.CellValue?.Text ?? "", "单元格结果是 Excel 错误，不能作为正常数据更新报告。");
        try
        {
            XLCellValue value;
            if (isFormula && recalculate)
            {
                value = cell.Value;
            }
            // t=d 保存的是实际ISO日期，不能像日期序号一样再加1904偏移。
            else if (saved?.DataType?.Value == CellValues.Date && saved.CellValue is not null)
                value = XmlConvert.ToDateTime(saved.CellValue.Text, XmlDateTimeSerializationMode.RoundtripKind);
            else value = cell.CachedValue;
            if (value.IsError) return Result(value.ToString(CultureInfo.CurrentCulture), "单元格结果是 Excel 错误，不能作为正常数据更新报告。");
            // 编辑使用真实类型及原值，不能从百分比、千分位、横线等显示文字猜测；公式始终提供原式。
            if (!isFormula)
            {
                if (value.IsBlank) { editableKind = ExcelEditKind.Text; editableValue = ""; }
                else if (value.IsText) { editableKind = ExcelEditKind.Text; editableValue = value.GetText(); }
                else if (value.IsNumber) { editableKind = ExcelEditKind.Number; editableValue = value.GetNumber().ToString("R", CultureInfo.InvariantCulture); }
                else if (value.IsDateTime) { editableKind = ExcelEditKind.DateTime; editableValue = value.GetDateTime().ToString(EditableDateFormat, CultureInfo.InvariantCulture); }
            }
            if (value.IsBlank) return Result("");
            if (value.IsText) return Result(value.GetText());
            if (value.IsBoolean) return Result(value.GetBoolean().ToString());
            var numberFormat = cell.Style.NumberFormat;
            // 原 uRv0H7KoE 只对精确会计格式中的数值0使用 DalSetting.ZeroFormat，不影响其他格式。
            if (numberFormat.Format == OriginalAccountingFormat && value.IsNumber && value.GetNumber() == 0)
                return Result(setting.ZeroFormat ?? "");
            // 修正原函数仅看编号、忽略文件显式formatCode的缺陷；其余已确认的原日期规则照用。
            var dateFormat = numberFormat.Format == OriginalLongDateFormat ? OriginalFullDateFormat
                : string.IsNullOrWhiteSpace(numberFormat.Format) ? OriginalDateFormat(numberFormat.NumberFormatId) : null;
            if (dateFormat is not null)
            {
                // DateTime已由库换好基准；仍是数字的1904缓存才需换基准，然后调用库的日期转换。
                var dateValue = value.IsNumber && cell.Worksheet.Workbook.Use1904DateSystem
                    ? (XLCellValue)(value.GetNumber() + 1462) : value;
                if (!dateValue.TryConvert(out DateTime date))
                    return Result("", "日期值超出有效范围，不能据此更新报告。");
                if (!isFormula)
                {
                    editableKind = ExcelEditKind.DateTime;
                    editableValue = date.ToString(EditableDateFormat, CultureInfo.InvariantCulture);
                }
                return Result(date.ToString(dateFormat, CultureInfo.CurrentCulture));
            }
            if (string.IsNullOrWhiteSpace(numberFormat.Format) && numberFormat.NumberFormatId is >= 50 and <= 54)
                return Result("", $"内置地区格式 {numberFormat.NumberFormatId} 尚未确认显示规则，请在底稿中指定明确格式后再更新。");
            // 仅把数字类值放进无公式的临时格；Value 后赋 Style，避免日期自动样式覆盖原格式。
            formatCell.Value = value;
            formatCell.Style = cell.Style;
            return Result(formatCell.GetFormattedString(CultureInfo.CurrentCulture));
        }
        catch (Exception e)
        {
            // 显式重算失败不能再退回旧缓存；源工作簿只存在于内存，从不保存。
            return Result("", $"{(recalculate && isFormula ? "公式重新计算" : "单元格读取")}失败：{e.Message}", false);
        }
    }

    private static Dictionary<string, Cell> ReadSavedCells(byte[] bytes, string sheetName, XLWorkbook workbook,
        bool prepareCalculation, out string? calculationError)
    {
        calculationError = null;
        using var stream = new MemoryStream(bytes, false);
        using var document = SpreadsheetDocument.Open(stream, false);
        var part = document.WorkbookPart ?? throw new InvalidDataException("底稿中没有工作簿。");
        var sourceBook = part.Workbook ?? throw new InvalidDataException("底稿中没有工作簿内容。");
        var result = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in sourceBook.Sheets!.Elements<Sheet>())
        {
            var isRequested = string.Equals(sheet.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase);
            if (!isRequested && !prepareCalculation) continue;
            var worksheet = (WorksheetPart)part.GetPartById(sheet.Id!.Value!);
            var content = worksheet.Worksheet ?? throw new InvalidDataException($"工作表 {sheet.Name} 没有内容。");
            var memorySheet = workbook.Worksheet(sheet.Name!.Value!);
            var rowNumber = 0;
            foreach (var row in content.GetFirstChild<SheetData>()?.Elements<Row>() ?? [])
            {
                rowNumber = row.RowIndex is null ? rowNumber + 1 : checked((int)row.RowIndex.Value);
                var columnNumber = 0;
                foreach (var cell in row.Elements<Cell>())
                {
                    // OOXML允许省略行和单元格地址；与读取库一样按当前位置还原，不能直接跳过原值检查。
                    var address = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(address))
                        address = XLHelper.GetColumnLetterFromNumber(++columnNumber) + rowNumber;
                    var memoryCell = memorySheet.Cell(address);
                    columnNumber = memoryCell.Address.ColumnNumber;
                    if (memoryCell.Address.RowNumber != rowNumber)
                        throw new InvalidDataException($"工作表 {sheet.Name} 的单元格地址与行位置不一致：{address}。");
                    if (isRequested) result.Add(memoryCell.Address.ToStringRelative(), cell);
                    // 未识别错误不能被当成空白，也不能换成其他错误：ERROR.TYPE等公式会据类别算出不同结果。
                    // 只限制显式重算；默认读取缓存照常。公式的旧错误仍允许由实际重新计算替代。
                    if (prepareCalculation && cell.DataType?.Value == CellValues.Error
                        && !memoryCell.HasFormula && !memoryCell.CachedValue.IsError)
                        calculationError ??= $"工作表“{sheet.Name}”的 {address} 包含暂不支持重新计算的错误常量 {cell.CellValue?.Text}，请读取文件已保存的结果。";
                }
            }
        }
        return result;
    }

    // 复用原 uRv0H7KoE 的已静态还原字面；索引来自静态还原字符串_20260908_131919.json，
    // 文件SHA256：835941DF5BF5017865AEB119F1D9713608A7DAB82F3770B8C70D7EC8A76489F8。
    private const string OriginalAccountingFormat = "_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * \"-\"??_ ;_ @_ "; // 1684，末尾空格也是精确匹配的一部分。
    private const string OriginalLongDateFormat = "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy"; // 1914
    private const string OriginalFullDateFormat = "yyyy年M月d日"; // 1858
    private static string? OriginalDateFormat(int formatId) => formatId switch
    {
        14 => "yyyy-M-d", // 1792
        27 or 36 or 57 => "yyyy年M月", // 1812
        28 or 29 or 58 => "M月d日", // 1830
        30 => "M/d/yy", // 1842
        31 => OriginalFullDateFormat,
        32 or 34 or 55 => "h时mm分", // 1880
        33 or 35 or 56 => "h时mm分ss秒", // 1894
        _ => null
    };

    private static string RangeAddressText(IXLRangeAddress address) => address.FirstAddress.ToStringRelative() == address.LastAddress.ToStringRelative()
        ? address.FirstAddress.ToStringRelative() : address.ToStringRelative();

    private static bool IsAddress(string address, bool allowRange) => Regex.IsMatch(address,
        allowRange ? @"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}(:\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})?$"
            : @"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}$");

    private static (string? Sheet, string Name) SplitSheetName(string text)
    {
        var match = Regex.Match(text, @"^(?:'(?<quoted>(?:[^']|'')+)'|(?<plain>[^!'\[\]:]+))!(?<name>[^!]+)$");
        return match.Success
            ? (match.Groups["quoted"].Success ? match.Groups["quoted"].Value.Replace("''", "'") : match.Groups["plain"].Value, match.Groups["name"].Value)
            : (null, text);
    }
}
