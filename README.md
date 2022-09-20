# Excelio
### A simple way to read and write data to Excel.

Excelio is built to make getting data out of an Excel document as easy as possible. To do this Excelio uses the standard OpenXml library to handle the hard stuff then hides the complexity and provides a simple easy to use interface. As bonus using OpenXml directly makes Excelio pretty fast!

First of let’s start with a simple example. What if you want to open an Excel doc and pull a sheet into an array.


```
string[,] arr = ExcelIO.WorkBook.ToArray("Example.xlsx", "Sheet1");
```
 
Keep the document open and pull from more than one sheet. 
```
using (var wb = new ExcelIO.WorkBook("Example.xlsx"))
{
    string[,] sheet1Arr = wb.Sheets[0].ToArray();
    string[,] sheet2Arr = wb.Sheets[1].ToArray();
}
```

Open a document and pull a specific cell.
```
using (var wb = new ExcelIO.WorkBook("Example.xlsx"))
{
    string C4Value = wb.Sheets.First().GetCell("C4").Value;
}
```

Set a single cell to a value and save.
```
using (var wb = new ExcelIO.WorkBook("Example.xlsx"))
{
    var C4Value = "New Value";
    wb.Sheets.First().SetCell("C4", C4Value);
    wb.Save();
}
```

Fill in a full sheet from an array.
```
var cells = new string[,]
    {
        { "one" },
        { "two" },
        { "three" }
    };
    
using (var wb = new ExcelIO.WorkBook("Example.xlsx"))
{
    wb.Sheets.First().Set(cells);
    wb.Save();
}
```
