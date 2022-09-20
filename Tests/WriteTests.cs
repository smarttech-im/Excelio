using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System;
using System.IO;

namespace Tests;

[TestClass]
public class WriteTests
{
    private static void ClearTempFiles()
    {
        var tempDir = $"lib/temp";

        if (!Directory.Exists(tempDir))
        {
            Directory.CreateDirectory(tempDir);
        }
        foreach (var file in Directory.GetFiles(tempDir))
        {
            File.Delete(file);
        }
    }

    private static string NewTempFileName()
    {
        ClearTempFiles();
        return $"lib/temp/{Guid.NewGuid()}.xlsx";
    }

    [TestMethod]
    public void NewSheet()
    {
        var NewFileName = NewTempFileName();
        var longName = Guid.NewGuid().ToString();
        var withBannedChars = @"/\?*:[] After Banned Chars";
        var startsWithApostrophe = "'Starts With Apostrophe";
        var bannedWord = "History";
        File.Copy("lib/Multi_Sheet.xlsx", NewFileName);

        using (var wb = new ExcelIO.WorkBook(NewFileName))
        {
            Assert.IsNotNull(wb);
            wb.Sheets.Add(longName);
            wb.Sheets.Add(withBannedChars);
            wb.Sheets.Add(startsWithApostrophe);
            wb.Sheets.Add(bannedWord);
            wb.Save();
        }

        using (var wb = new ExcelIO.WorkBook(NewFileName))
        {
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb.Sheets.Any(i => i.Name == "Starts With Apostrophe"));
            Assert.IsTrue(wb.Sheets.Any(i => i.Name == " After Banned Chars"));
            Assert.AreEqual(7, wb.Sheets.Count());
        }
    }

    [TestMethod]
    public void RemoveSheet()
    {
        using var wb = new ExcelIO.WorkBook("lib/Multi_Sheet.xlsx");
        Assert.IsNotNull(wb);
        Assert.AreEqual(3, wb.Sheets.Count());
        wb.Sheets.Remove("Sheet2");
        Assert.AreEqual(2, wb.Sheets.Count());
    }

    [TestMethod]
    public void SetCellValues()
    {
        var NewFileName = NewTempFileName();
        string A10 = Guid.NewGuid().ToString();
        using (var wb = new ExcelIO.WorkBook("lib/List_With_Headers.xlsx"))
        {
            Assert.IsNotNull(wb);
            wb.Sheets.First().SetCell("A10", A10);
            wb.SaveAs(NewFileName);
        }

        using (var wb = new ExcelIO.WorkBook(NewFileName))
        {
            Assert.IsNotNull(wb);
            Assert.AreEqual(A10, wb.Sheets.First().GetCell("A10").Value);
        }
    }

    [TestMethod]
    public void SetSheetArray()
    {
        var NewFileName = NewTempFileName();
        var cells = new string[,]
        {
            { "one" },
            { "two" },
            { "three" }
        };
        
        using (var wb = new ExcelIO.WorkBook("lib/List_With_Headers.xlsx"))
        {
            Assert.IsNotNull(wb);
            wb.Sheets.First().Set(cells);
            wb.SaveAs(NewFileName);
        }

        using (var wb = new ExcelIO.WorkBook(NewFileName))
        {
            Assert.IsNotNull(wb);
            Assert.AreEqual("two", wb.Sheets.First().GetCell("B1").Value);
        }
    }

    [TestMethod]
    public void SetSheetCSV()
    {
        var NewFileName = NewTempFileName();
        var cells = "one,two,three\nfour,five";

        using (var wb = new ExcelIO.WorkBook("lib/List_With_Headers.xlsx"))
        {
            Assert.IsNotNull(wb);
            wb.Sheets.First().Set(cells, Delimiter: ",", NewLine: "\n");
            wb.SaveAs(NewFileName);
        }

        using (var wb = new ExcelIO.WorkBook(NewFileName))
        {
            Assert.IsNotNull(wb);
            Assert.AreEqual("five", wb.Sheets.First().GetCell("B2").Value);
        }
    }

    [TestMethod]
    public void SetObjectList()
    {
        var NewFileName = NewTempFileName();
        var items = new TestItem[] 
        { 
            new TestItem() { Make = "Hyundai", Model = "Creta", Price = 104400},
            new TestItem() { Make = "Toyota", Model = "Fortuner", Price = 324000},
            new TestItem() { Make = "Mahindra", Model = null, Price = 781000},
            new TestItem() { Make = "Tata", Model = "Punch", Price = 593000},
        };

        using (var wb = new ExcelIO.WorkBook("lib/List_With_Headers.xlsx"))
        {
            Assert.IsNotNull(wb);
            wb.Sheets.First().Set(items, TitlesOnFirstRow: true);
            wb.SaveAs(NewFileName);
        }

        using (var wb = new ExcelIO.WorkBook(NewFileName))
        {
            Assert.IsNotNull(wb);
            Assert.AreEqual("Fortuner", wb.Sheets.First().GetCell("B3").Value);
        }
    }

}