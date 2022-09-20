using System.Data;

namespace ExcelIO;

internal static partial class Converter
{
    public static string[,] ToArray(Sheet Source)
    {
        var (cells, width, height) = Source.ToCellListWithRange();
        var matrix = new string[width, height];

        foreach (var cell in cells)
        {
            matrix[cell.X, cell.Y] = cell.Value;
        }

        return matrix;
    }

    public static Cell[,] ToCellArray(Sheet Source)
    {
        var (cells, width, height) = Source.ToCellListWithRange();
        var matrix = new Cell[width, height];

        foreach (var cell in cells)
        {
            matrix[cell.X, cell.Y] = cell;
        }

        return matrix;
    }

    public static string ToCSVString(Sheet Source, string Delimiter = ",", string NewLine = "")
    {
        var result = "";
        var arr = Source.ToCellArray();
        if (string.IsNullOrWhiteSpace(NewLine))
        {
            NewLine = Environment.NewLine;
        }

        for (int y = 0; y < arr.GetLength(1); y++)
        {
            for (int x = 0; x < arr.GetLength(0); x++)
            {
                var del = x == 0 ? "" : Delimiter;
                result += $"{del}\"{arr[x, y].Value}\"";
            }
            result += NewLine;
        }

        return result;
    }

    public static List<List<Cell>> ToList(Sheet Source)
    {
        var rows = new List<List<Cell>>();
        var arr = Source.ToCellArray();

        for (int y = 0; y < arr.GetLength(1); y++)
        {
            var row = new List<Cell>();
            for (int x = 0; x < arr.GetLength(0); x++)
            {
                row.Add(arr[x, y]);
            }
            rows.Add(row);
        }

        return rows;
    }

    public static DataTable ToDataTable(Sheet Source, bool FieldNamesOnFirstRow = false, bool InferTypes = false)
    {
        var (cells, width, height) = Source.ToCellListWithRange();

        var result = new DataTable
        {
            TableName = Source.Name
        };

        if (cells.Count == 0 || height == 0 || width == 0)
        {
            return result;
        }

        var firstRow = cells.Where(i => i.Y == 0).ToList();
        var columns = firstRow.Select(x => new DataColumn(FieldNamesOnFirstRow ? x.Value : "")).ToArray();
        result.Columns.AddRange(columns);
        if (columns == null)
        {
            return result;
        }

        if (FieldNamesOnFirstRow)
        {
            cells.RemoveAll(x => x.Y == 0);

            if (cells.Count == 0)
            {
                return result;
            }
        }

        if (InferTypes)
        {
            for (int i = 0; i < columns.Length; i++)
            {
                columns[i].DataType = cells
                    .Where(c => c.X == i)
                    .Select(s => s.Value)
                    .ToArray()
                    .LowestCommonType();
            }
        }

        for (int y = 0; y < height; y++)
        {
            var dataRow = result.NewRow();
            var row = cells.Where(i => i.Y == y).OrderBy(i => i.X).ToArray();

            for (int x = 0; x < width; x++)
            {
                var cellValue = row.Where(i => i.X == x).FirstOrDefault()?.Value;
                dataRow[x] = cellValue == null ? DBNull.Value : cellValue;
            }
            result.Rows.Add(dataRow);
        }

        return result;
    }
}
