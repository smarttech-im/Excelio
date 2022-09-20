namespace ExcelIO;

internal static partial class Converter
{
    public static string[,] FromCSVString(string CSVString, string Delimiter = ",", string NewLine = "")
    {
        if(string.IsNullOrWhiteSpace(CSVString))
        {
            return new string[0, 0];
        }

        if (string.IsNullOrEmpty(NewLine))
        {
            NewLine = Environment.NewLine;
        }

        var lines = CSVString.Split(NewLine);
        var maxColCount = lines.Select(i => i.Split(Delimiter).Length).Max();
        var cells = new string[maxColCount, lines.Length];

        for (int y = 0; y < lines.Length; y++)
        {
            var line = lines[y].Split(Delimiter);
            for (int x = 0; x < line.Length; x++)
            {
                cells[x, y] = line[x];
            }
        }

        return cells;
    }

    public static string[,] FromGeneric(IEnumerable<object> Items, bool TitlesOnFirstRow = false)
    {
        if (Items == null ||
            !Items.Any())
        {
            return new string[0, 0];
        }

        var itemList = Items.ToArray();
        var colNames = Items.First().GetType().GetProperties().Select(i => i.Name).ToArray();
        var offset = TitlesOnFirstRow ? 1 : 0;
        var cells = new string[colNames.Length, itemList.Length + offset];

        if (TitlesOnFirstRow)
        {
            for (int x = 0; x < colNames.Length; x++)
            {
                cells[x, 0] = colNames[x];
            }
        }

        for (int y = 0; y < itemList.Length; y++)
        {
            var itemValues = itemList[y].GetType().GetProperties().Select(i => i.GetValue(itemList[y])).ToArray();
            for (int x = 0; x < itemValues.Length && x < colNames.Length; x++)
            {
                cells[x, y + offset] = itemValues[x]?.ToString() ?? "";
            }
        }

        return cells;
    }
}
