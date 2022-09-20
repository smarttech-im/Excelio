using Packaging = DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

namespace ExcelIO;

public class WorkBook : IDisposable
{
    public string FileName { get; init; }
    public SheetList Sheets { get; private set; } = new();
    internal Packaging.SpreadsheetDocument Doc { get; set; }
    internal Packaging.WorkbookPart? Part { get; set; }
    internal Spreadsheet.Workbook? Book { get; set; }
    internal Spreadsheet.SharedStringItem[]? SharedItems { get; set; }

    public WorkBook(string FileName, bool IsEditable = true)
    {
        this.FileName = 
            string.IsNullOrEmpty(FileName) ? 
            "" : 
            FileName;

        if (!File.Exists(FileName))
        {
            throw new FileNotFoundException(FileName);
        }

        Doc = Packaging.SpreadsheetDocument.Open(FileName, IsEditable, new Packaging.OpenSettings { AutoSave = false });
        Part = Doc.WorkbookPart;
        Book = Part?.Workbook;
        var sharedStringPart = Part?.SharedStringTablePart;
        SharedItems = sharedStringPart?.SharedStringTable.Elements<Spreadsheet.SharedStringItem>().ToArray();

        RefreshSheets();
    }

    public void Dispose()
    {
        Doc.Close();
        Doc?.Dispose();
        GC.SuppressFinalize(this);
    }

    private void RefreshSheets()
    {
        var sheets = Book?.Descendants<Spreadsheet.Sheet>();
        if (sheets != null && 
            SharedItems != null)
        {
            Sheets = new SheetList(
                Doc,
                Book,
                sheets
                .Where(i => 
                    i != null && 
                    i.Name != null)
                .Select(i => new Sheet
            {
                SheetID = i?.SheetId,
                FileName = FileName,
                Name = i!.Name,
                Part = Part,
                ConnectedSheet = i,
                SharedItems = SharedItems,
            }).ToList());
        }
    }

    public void Save()
    {
        if (Doc == null)
        {
            return;
        }
        Doc.Save();
    }

    public void SaveAs(string FileName)
    {
        if (Doc == null)
        {
            return;
        }
        Doc.SaveAs(FileName).Close();
    }

    public static string[,] ToArray(string FileName, string SheetName)
    {
        using var wb = new WorkBook(FileName);
        var sheet = wb.Sheets.Where(i => i.Name == SheetName).FirstOrDefault() ?? new Sheet();
        return 
            sheet == null ?
            new string[0, 0] :
            Converter.ToArray(sheet);
    }

    public static Cell[,] ToCellArray(string FileName, string SheetName)
    {
        using var wb = new WorkBook(FileName);
        var sheet = wb.Sheets.Where(i => i.Name == SheetName).FirstOrDefault() ?? new Sheet();
        return 
            sheet == null ?
            new Cell[0, 0] :
            Converter.ToCellArray(sheet);
    }

    public static string ToCSVString(string FileName, string SheetName)
    {
        using var wb = new WorkBook(FileName);
        var sheet = wb.Sheets.Where(i => i.Name == SheetName).FirstOrDefault() ?? new Sheet();
        return 
            sheet == null ?
            "" :
            Converter.ToCSVString(sheet);
    }

    public static List<List<Cell>> ToList(string FileName, string SheetName)
    {
        using var wb = new WorkBook(FileName);
        var sheet = wb.Sheets.Where(i => i.Name == SheetName).FirstOrDefault() ?? new Sheet();
        return 
            sheet == null ?
            new List<List<Cell>>() :
            Converter.ToList(sheet);
    }

    public static DataTable ToDataTable(string FileName, string SheetName, bool FieldNamesOnFirstRow = false, bool InferTypes = false)
    {
        using var wb = new WorkBook(FileName);
        var sheet = wb.Sheets.Where(i => i.Name == SheetName).FirstOrDefault() ?? new Sheet();
        return 
            sheet == null ?
            new DataTable() :
            Converter.ToDataTable(sheet, FieldNamesOnFirstRow, InferTypes);
    }

}