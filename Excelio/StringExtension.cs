namespace ExcelIO;

internal static class StringExtension
{
    public static bool IsBool(this string Source) => bool.TryParse(Source, out _);

    public static bool IsInt(this string Source) => int.TryParse(Source, out _);

    public static bool IsDecimal(this string Source) => decimal.TryParse(Source, out _);

    public static bool IsDateTime(this string Source) => DateTime.TryParse(Source, out _);

    public static Type LowestCommonType(this string[] ValueList)
    {
        if (ValueList == null) return typeof(string);
        if (ValueList.All(x => x.IsBool())) return typeof(bool);
        if (ValueList.All(x => x.IsInt())) return typeof(int);
        if (ValueList.All(x => x.IsDecimal())) return typeof(decimal);
        if (ValueList.All(x => x.IsDateTime())) return typeof(DateTime);

        return typeof(string);
    }

}
