using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace Tests;

[TestClass]
public class ReadTests
{

    [TestMethod]
    public void ArrayStatic()
    {
        var arr = ExcelIO.WorkBook.ToArray("lib/List_With_Headers.xlsx", "Sheet1");
        Assert.IsNotNull(arr);
        Assert.AreEqual(10, arr.GetLength(0));
        Assert.AreEqual(14, arr.GetLength(1));
        Assert.AreEqual("2022-08-19", arr[2,6]);
    }

    [TestMethod]
    public void Array()
    {
        using var wb = new ExcelIO.WorkBook("lib/List_With_Headers.xlsx");
        var arr = wb.Sheets.First().ToArray();
        Assert.IsNotNull(arr);
        Assert.AreEqual(10, arr.GetLength(0));
        Assert.AreEqual(14, arr.GetLength(1));
    }
    
    [TestMethod]
    public void PickSpecificCells()
    {
        using var wb = new ExcelIO.WorkBook("lib/List_With_Headers.xlsx");
        var sheet = wb.Sheets.First();
        Assert.AreEqual("H123456-007", sheet.GetCell("A7").Value);
        Assert.AreEqual("Weekday", sheet.GetCell("G11").Value);
        Assert.AreEqual("booking", sheet.GetCell(1, 3).Value);
    }
}