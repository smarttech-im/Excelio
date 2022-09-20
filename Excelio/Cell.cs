namespace ExcelIO;

public class Cell
{
    public string Value { get; init; } = string.Empty;
    public int X { get; private set; }
    public int Y { get; private set; }
    private string _CellReference = string.Empty;

    public string CellReference
    {
        get
        {
            return _CellReference;
        }
        set
        {
            _CellReference = value;
            var (x, y) = TranslateToIndex(value);
            X = x;
            Y = y;
        }
    }

    public Type ValueType
    {
        get
        {
            return InferType(Value);
        }
    }

    public override string ToString()
    {
        return Value;
    }

    public static (int X, int Y) TranslateToIndex(string CellReference)
    {
        if (string.IsNullOrWhiteSpace(CellReference))
        {
            return (X: 0, Y: 0);
        }

        var result = (X: 1, Y: 1);
        var letters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
        var letterPart = new string(CellReference.Where(i => char.IsLetter(i)).ToArray());
        var numberPart = new string(CellReference.Where(i => char.IsDigit(i)).ToArray());

        foreach (var c in letterPart.ToCharArray().Reverse())
        {
            result.X *= Array.IndexOf(letters, c);
        }

        result.X--;
        if (int.TryParse(numberPart, out result.Y))
        {
            result.Y--;
        }

        return result;
    }

    public static string TranslateToReference(int X, int Y)
    {
        var letters = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        int baseLength = 26;
        string result = string.Empty;
        int index = (X % baseLength) + 1;

        result = letters[index] + result;
        X /= baseLength;

        while (X > 0)
        {
            index = (X % baseLength);
            result = letters[index] + result;
            X /= baseLength;
        }

        return result + (Y + 1).ToString();
    }

    private static Type InferType(string Value)
    {
        if (bool.TryParse(Value, out _)) return typeof(bool);
        if (int.TryParse(Value, out _)) return typeof(int);
        if (decimal.TryParse(Value, out _)) return typeof(decimal);
        if (DateTime.TryParse(Value, out _)) return typeof(DateTime);
        return typeof(string);
    }
}
