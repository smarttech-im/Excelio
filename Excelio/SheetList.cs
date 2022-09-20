using Packaging = DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;

namespace ExcelIO;

public class SheetList : IEnumerable<Sheet>
{
    private readonly Packaging.SpreadsheetDocument Doc;
    private readonly Spreadsheet.Workbook? Book;
    private readonly List<Sheet> InnerList = new();

    public Sheet this[int index]
    {
        get { return InnerList[index]; }
        set { InnerList.Insert(index, value); }
    }

    public SheetList()
    {
        Book = new Spreadsheet.Workbook();
        InnerList = new List<Sheet>();
    }

    public SheetList(Packaging.SpreadsheetDocument Document, Spreadsheet.Workbook? Parent, List<Sheet> NewList)
    {
        Doc = Document;
        Book = Parent;
        InnerList = NewList;
    }

    public IEnumerator<Sheet> GetEnumerator()
    {
        return InnerList.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    public void Add(string Name)
    {
        Add(new Sheet { Name = Name });
    }

    private static string TidyName(string Name, string SheetID)
    {
        string result =
            string.IsNullOrWhiteSpace(Name) ?
            "Sheet" + SheetID :
            Name;

        if (result.ToLower() == "history")
        {
            result = "Sheet" + SheetID;
        }

        if (result.StartsWith('\''))
        {
            result = result[1..];
        }

        if (result.EndsWith('\''))
        {
            result = result.Remove(result.Length - 1);
        }

        foreach (var bannedChar in @"/\?*:[]".ToCharArray())
        {
            result = result.Replace(bannedChar.ToString(), "");
        }
        
        if (result.Length > 30)
        {
            result = result[..30];
        }

        return result;
    }

    private void Add(Sheet item)
    {
        if (Doc.WorkbookPart == null ||
            Book == null ||
            Book.WorkbookPart == null)
        {
            return;
        }

        var sheets = Doc.WorkbookPart.Workbook.GetFirstChild<Spreadsheet.Sheets>();
        if (sheets == null)
        {
            return;
        }

        // Add a blank WorksheetPart.
        var newWorksheetPart = Doc.WorkbookPart.AddNewPart<Packaging.WorksheetPart>();
        newWorksheetPart.Worksheet = new Spreadsheet.Worksheet(new Spreadsheet.SheetData());
        string relationshipID = Book.WorkbookPart.GetIdOfPart(newWorksheetPart);


        // Get a unique ID for the new worksheet.
        uint sheetID =
            sheets.Elements<Spreadsheet.Sheet>().Any() ?
            sheets.Elements<Spreadsheet.Sheet>().Select(s => s.SheetId?.Value ?? 0).Max() + 1 :
            1;

        // Give the new worksheet a name.
        string sheetName = TidyName(item.Name, sheetID.ToString());

        // Append the new worksheet and associate it with the workbook.
        var sheet = new Spreadsheet.Sheet()
        {
            Id = relationshipID,
            SheetId = sheetID,
            Name = sheetName
        };

        sheets.Append(sheet);
        InnerList.Add(new Sheet
        {
            SheetID = sheet?.SheetId,
            FileName = item.FileName,
            Name = sheet!.Name,
            Part = item.Part,
            ConnectedSheet = sheet,
            SharedItems = item.SharedItems,
        });

    }

    public void Remove(string Name)
    {
        if (Doc.WorkbookPart == null ||
            Book == null ||
            Book.WorkbookPart == null)
        {
            return;
        }

        var sheets = Doc.WorkbookPart.Workbook.GetFirstChild<Spreadsheet.Sheets>();
        if (sheets == null)
        {
            return;
        }

        var sheet = sheets.Elements<Spreadsheet.Sheet>().Where(i => i.Name == Name).FirstOrDefault();
        if (sheet == null)
        {
            return;
        }

        sheets.RemoveChild(sheet);

        var innerSheet = InnerList.Where(i => i.Name == Name).FirstOrDefault();
        if (innerSheet == null)
        {
            return;
        }

        InnerList.Remove(innerSheet);
    }
}
