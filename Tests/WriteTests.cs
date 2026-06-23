using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;

namespace Tests;

[TestClass]
public class WriteTests
{
    private static void ClearTempFiles()
    {
        string tempDir = $"lib/temp";

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
        string NewFileName = NewTempFileName();
        string longName = Guid.NewGuid().ToString();
        string withBannedChars = @"/\?*:[] After Banned Chars";
        string startsWithApostrophe = "'Starts With Apostrophe";
        string bannedWord = "History";
        File.Copy("lib/Multi_Sheet.xlsx", NewFileName);

        using (ExcelIO.WorkBook wb = new(NewFileName))
        {
            Assert.IsNotNull(wb);
            wb.Sheets.Add(longName);
            wb.Sheets.Add(withBannedChars);
            wb.Sheets.Add(startsWithApostrophe);
            wb.Sheets.Add(bannedWord);
            wb.Save();
        }

        using (ExcelIO.WorkBook wb = new(NewFileName))
        {
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb.Sheets.Any(i => i.Name == "Starts With Apostrophe"));
            Assert.IsTrue(wb.Sheets.Any(i => i.Name == " After Banned Chars"));
            Assert.AreEqual(7, wb.Sheets.Count());
        }
    }

    [TestMethod]
    public void NewFile()
    {
        string NewFileName = NewTempFileName();
        string longName = Guid.NewGuid().ToString();
        string withBannedChars = @"/\?*:[] After Banned Chars";
        string startsWithApostrophe = "'Starts With Apostrophe";
        string bannedWord = "History";

        using (ExcelIO.WorkBook wb = new(NewFileName))
        {
            Assert.IsNotNull(wb);
            wb.Sheets.Add(longName);
            wb.Sheets.Add(withBannedChars);
            wb.Sheets.Add(startsWithApostrophe);
            wb.Sheets.Add(bannedWord);
            wb.Save();
        }

        using (ExcelIO.WorkBook wb = new(NewFileName))
        {
            Assert.IsNotNull(wb);
            Assert.IsTrue(wb.Sheets.Any(i => i.Name == "Starts With Apostrophe"));
            Assert.IsTrue(wb.Sheets.Any(i => i.Name == " After Banned Chars"));
            Assert.AreEqual(5, wb.Sheets.Count());
        }
    }

    [TestMethod]
    public void RemoveSheet()
    {
        using ExcelIO.WorkBook wb = new("lib/Multi_Sheet.xlsx");
        Assert.IsNotNull(wb);
        Assert.AreEqual(3, wb.Sheets.Count());
        wb.Sheets.Remove("Sheet2");
        Assert.AreEqual(2, wb.Sheets.Count());
    }

    [TestMethod]
    public void SetCellValues()
    {
        string NewFileName = NewTempFileName();
        string A10 = Guid.NewGuid().ToString();
        using (ExcelIO.WorkBook wb = new("lib/List_With_Headers.xlsx"))
        {
            Assert.IsNotNull(wb);
            wb.Sheets.First().SetCell("A10", A10);
            wb.SaveAs(NewFileName);
        }

        using (ExcelIO.WorkBook wb = new(NewFileName))
        {
            Assert.IsNotNull(wb);
            Assert.AreEqual(A10, wb.Sheets.First().GetCell("A10").Value);
        }
    }

    [TestMethod]
    public void SetSheetArray()
    {
        string NewFileName = NewTempFileName();
        string[,] cells = new string[,]
        {
            { "one" },
            { "two" },
            { "three" }
        };

        using (ExcelIO.WorkBook wb = new("lib/List_With_Headers.xlsx"))
        {
            Assert.IsNotNull(wb);
            wb.Sheets.First().Set(cells);
            wb.SaveAs(NewFileName);
        }

        using (ExcelIO.WorkBook wb = new(NewFileName))
        {
            Assert.IsNotNull(wb);
            Assert.AreEqual("two", wb.Sheets.First().GetCell("B1").Value);
        }
    }

    [TestMethod]
    public void SetSheetCSV()
    {
        string NewFileName = NewTempFileName();
        string cells = "one,two,three\nfour,five";

        using (ExcelIO.WorkBook wb = new("lib/List_With_Headers.xlsx"))
        {
            Assert.IsNotNull(wb);
            wb.Sheets.First().Set(cells, Delimiter: ",", NewLine: "\n");
            wb.SaveAs(NewFileName);
        }

        using (ExcelIO.WorkBook wb = new(NewFileName))
        {
            Assert.IsNotNull(wb);
            Assert.AreEqual("five", wb.Sheets.First().GetCell("B2").Value);
        }
    }

    [TestMethod]
    public void SetObjectList()
    {
        string NewFileName = NewTempFileName();
        TestItem[] items = new TestItem[]
        {
            new() { Make = "Hyundai", Model = "Creta", Price = 104400},
            new() { Make = "Toyota", Model = "Fortuner", Price = 324000},
            new() { Make = "Mahindra", Model = null, Price = 781000},
            new() { Make = "Tata", Model = "Punch", Price = 593000},
        };

        using (ExcelIO.WorkBook wb = new("lib/List_With_Headers.xlsx"))
        {
            Assert.IsNotNull(wb);
            wb.Sheets.First().Set(items, TitlesOnFirstRow: true);
            wb.SaveAs(NewFileName);
        }

        using (ExcelIO.WorkBook wb = new(NewFileName))
        {
            Assert.IsNotNull(wb);
            Assert.AreEqual("Fortuner", wb.Sheets.First().GetCell("B3").Value);
        }
    }
}