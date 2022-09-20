namespace ExcelIO;

internal class Functions
{
    public static void SaveToCSV(Sheet Source, string FileName) => File.WriteAllTextAsync(FileName, Source.ToCSVString());
}
