using Packaging = DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelIO;

public class Sheet
{
    public string FileName { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    internal DocumentFormat.OpenXml.UInt32Value? SheetID { get; set; }
    internal Spreadsheet.Sheet ConnectedSheet { get; init; } = new Spreadsheet.Sheet();
    internal Packaging.WorkbookPart? Part { get; set; }
    internal Spreadsheet.SharedStringItem[] SharedItems { get; set; } = Array.Empty<Spreadsheet.SharedStringItem>();

    // Get a single cell
    public Cell GetCell(int X, int Y) => GetCell(Cell.TranslateToReference(X, Y));

    public Cell GetCell(string CellReference)
    {
        CellReference = CellReference.Trim().ToUpper();
        var (cells, sharedStringItems) = GetCells();
        string cellValue = cells
            .Where(i => i?.CellReference?.ToString() == CellReference)
            .Select(i =>
                i?.DataType?.Value == Spreadsheet.CellValues.SharedString &&
                int.TryParse(i?.CellValue?.Text, out int index) ?
                    sharedStringItems[index]?.InnerText ?? "" :
                    i?.CellValue?.Text ?? ""
            ).FirstOrDefault("");

        return new Cell { Value = cellValue, CellReference = CellReference };

    }

    // Update a single cell
    public void SetCell(int X, int Y, string Value) => SetCell(Cell.TranslateToReference(X, Y), Value);

    public void SetCell(string CellReference, string Value)
    {
        if (Part == null ||
            ConnectedSheet.Id == null)
        {
            return;
        }

        CellReference = CellReference.Trim().ToUpper();
        var worksheetPart = (Packaging.WorksheetPart)Part.GetPartById(ConnectedSheet!.Id);
        var cells = worksheetPart.Worksheet.Descendants<Spreadsheet.Cell>();
        if (cells == null)
        {
            return;
        }

        var cell = cells.Where(i => i?.CellReference?.ToString() == CellReference).FirstOrDefault();
        if (cell != null && cell.CellValue != null)
        {
            cell.CellValue.Text = Value;
            cell.DataType = Spreadsheet.CellValues.String;
            return;
        }

        uint rowIndex = uint.Parse(CellReference.ToCharArray().Where(i => char.IsNumber(i)).ToArray());
        Spreadsheet.Row row;
        var sheetData = worksheetPart.Worksheet.GetFirstChild<Spreadsheet.SheetData>();
        if (sheetData == null)
        {
            return;
        }

        if (sheetData.Elements<Spreadsheet.Row>().Any(r => r.RowIndex?.Value == rowIndex))
        {
            row = sheetData.Elements<Spreadsheet.Row>().Where(r => r?.RowIndex?.Value == rowIndex).First();
        }
        else
        {
            row = new Spreadsheet.Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        // Get next cell after the insert location, if there is one
        var refCell = row.Elements<Spreadsheet.Cell>().Where(c =>
                c.CellReference?.Value?.Length == CellReference.Length &&
                string.Compare(c.CellReference.Value, CellReference, true) > 0
            ).FirstOrDefault();

        // Create new cell then insert before the reference cell
        var newCell = new Spreadsheet.Cell()
        {
            CellReference = CellReference,
            CellValue = new Spreadsheet.CellValue { Text = Value },
            DataType = Spreadsheet.CellValues.String,
        };
        row.InsertBefore(newCell, refCell);
    }

    private (List<Spreadsheet.Cell>, Spreadsheet.SharedStringItem[]) GetCells()
    {
        if (!File.Exists(FileName) ||
            Part == null ||
            ConnectedSheet.Id == null)
        {
            return (new List<Spreadsheet.Cell>(), Array.Empty<Spreadsheet.SharedStringItem>());
        }

        var worksheetPart = (Packaging.WorksheetPart)Part.GetPartById(ConnectedSheet.Id);
        var cells = worksheetPart.Worksheet.Descendants<Spreadsheet.Cell>().ToList();
        if (cells == null)
        {
            return (new List<Spreadsheet.Cell>(), SharedItems);
        }

        return (cells, SharedItems);
    }

    internal List<Cell> ToCellList()
    {
        var result = new List<Cell>();
        var (cells, sharedStringItems) = GetCells();

        foreach (var cell in cells)
        {
            var cellReference = cell?.CellReference?.ToString() ?? "";
            var value = cell?.CellValue?.Text ?? "";

            // The cells contains a string input that is not a formula
            if (cell?.DataType?.Value == Spreadsheet.CellValues.SharedString &&
                int.TryParse(value, out int index))
            {
                value = sharedStringItems[index]?.InnerText ?? "";
            }

            result.Add(new Cell { Value = value, CellReference = cellReference });
        }

        return result;
    }

    internal (List<Cell> Cells, int Width, int Height) ToCellListWithRange()
    {
        var cells = ToCellList();
        var width = cells.Select(i => i.X).Max() + 1;
        var height = cells.Select(i => i.Y).Max() + 1;
        return (cells, width, height);
    }

    public string[,] ToArray() => Converter.ToArray(this);

    public Cell[,] ToCellArray() => Converter.ToCellArray(this);

    public string ToCSVString() => Converter.ToCSVString(this);

    public List<List<Cell>> ToList() => Converter.ToList(this);

    public System.Data.DataTable ToDataTable(bool FieldNamesOnFirstRow = false, bool InferTypes = false) =>
        Converter.ToDataTable(this, FieldNamesOnFirstRow, InferTypes);

    public void SaveToCSV(string FileName) => Functions.SaveToCSV(this, FileName);


    public void Set(string[,] Cells)
    {
        if (Part == null ||
            ConnectedSheet.Id == null)
        {
            return;
        }

        int limit = 65534;
        if (Cells.GetLength(0) > limit)
        {
            throw new ArgumentOutOfRangeException(nameof(Cells), "Array width exceeds size limit");
        }
        if (Cells.GetLength(1) > limit)
        {
            throw new ArgumentOutOfRangeException(nameof(Cells), "Array height exceeds size limit");
        }


        var worksheetPart = (Packaging.WorksheetPart)Part.GetPartById(ConnectedSheet!.Id);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<Spreadsheet.SheetData>();
        if (sheetData == null)
        {
            return;
        }
        sheetData.RemoveAllChildren();

        for (int y = 0; y < Cells.GetLength(1); y++)
        {
            var newRow = new Spreadsheet.Row();
            for (int x = 0; x < Cells.GetLength(0); x++)
            {
                var cellReference = Cell.TranslateToReference(x, y);
                newRow.Append(new Spreadsheet.Cell
                {
                    CellReference = cellReference,
                    CellValue = new Spreadsheet.CellValue { Text = Cells[x, y] },
                    DataType = Spreadsheet.CellValues.String,
                });
            }
            sheetData.Append(newRow);
        }
    }

    public void Set(string CSVString, string Delimiter = ",", string NewLine = "") => Set(Converter.FromCSVString(CSVString, Delimiter, NewLine));

    public void Set(IEnumerable<object> Items, bool TitlesOnFirstRow = false) => Set(Converter.FromGeneric(Items, TitlesOnFirstRow));
}
