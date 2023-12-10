using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.Rows
{
    [TestFixture]
    public class RowTests
    {
        [Test]
        public void RowsUsedIsFast()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().SetValue("Hello world!");
            var rowsUsed = ws.Column(1).AsRange().RowsUsed();
            Assert.That(rowsUsed.Count(), Is.EqualTo(1));
        }

        [Test]
        public void CopyRow()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Test").Style.Font.SetBold();
            ws.FirstRow().CopyTo(ws.Row(2));

            Assert.That(ws.Cell("A2").Style.Font.Bold, Is.True);
        }

        [Test]
        public void InsertingRowsAbove1()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Rows("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Row(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            var row1 = ws.Row(1);
            var row2 = ws.Row(2);
            var row3 = ws.Row(3);

            var rowIns = ws.Row(1).InsertRowsAbove(1).First();

            Assert.That(ws.Row(1).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));
            Assert.That(ws.Row(1).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));
            Assert.That(ws.Row(1).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));

            Assert.That(ws.Row(2).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(ws.Row(2).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(ws.Row(2).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

            Assert.That(ws.Row(3).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
            Assert.That(ws.Row(3).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
            Assert.That(ws.Row(3).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

            Assert.That(ws.Row(4).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(ws.Row(4).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(ws.Row(4).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

            Assert.That(ws.Row(3).Cell(2).GetString(), Is.EqualTo("X"));

            Assert.That(rowIns.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));
            Assert.That(rowIns.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));
            Assert.That(rowIns.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));

            Assert.That(row1.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(row1.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(row1.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

            Assert.That(row2.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
            Assert.That(row2.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
            Assert.That(row2.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

            Assert.That(row3.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(row3.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(row3.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

            Assert.That(row2.Cell(2).GetString(), Is.EqualTo("X"));
        }

        [Test]
        public void InsertingRowsAbove2()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Rows("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Row(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            var row1 = ws.Row(1);
            var row2 = ws.Row(2);
            var row3 = ws.Row(3);

            var rowIns = ws.Row(2).InsertRowsAbove(1).First();

            Assert.That(ws.Row(1).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(ws.Row(1).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(ws.Row(1).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

            Assert.That(ws.Row(2).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(ws.Row(2).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(ws.Row(2).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

            Assert.That(ws.Row(3).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
            Assert.That(ws.Row(3).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
            Assert.That(ws.Row(3).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

            Assert.That(ws.Row(4).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(ws.Row(4).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(ws.Row(4).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

            Assert.That(ws.Row(3).Cell(2).GetString(), Is.EqualTo("X"));

            Assert.That(rowIns.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(rowIns.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(rowIns.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

            Assert.That(row1.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(row1.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(row1.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

            Assert.That(row2.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
            Assert.That(row2.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
            Assert.That(row2.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

            Assert.That(row3.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(row3.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(row3.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

            Assert.That(row2.Cell(2).GetString(), Is.EqualTo("X"));
        }

        [Test]
        public void InsertingRowsAbove3()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Rows("1,3").Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.Row(2).Style.Fill.SetBackgroundColor(XLColor.Yellow);
            ws.Cell(2, 2).SetValue("X").Style.Fill.SetBackgroundColor(XLColor.Green);

            var row1 = ws.Row(1);
            var row2 = ws.Row(2);
            var row3 = ws.Row(3);

            var rowIns = ws.Row(3).InsertRowsAbove(1).First();

            Assert.That(ws.Row(1).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(ws.Row(1).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(ws.Row(1).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

            Assert.That(ws.Row(2).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
            Assert.That(ws.Row(2).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
            Assert.That(ws.Row(2).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

            Assert.That(ws.Row(3).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
            Assert.That(ws.Row(3).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
            Assert.That(ws.Row(3).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

            Assert.That(ws.Row(4).Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(ws.Row(4).Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(ws.Row(4).Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

            Assert.That(ws.Row(2).Cell(2).GetString(), Is.EqualTo("X"));

            Assert.That(rowIns.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
            Assert.That(rowIns.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
            Assert.That(rowIns.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

            Assert.That(row1.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(row1.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(row1.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

            Assert.That(row2.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));
            Assert.That(row2.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Green));
            Assert.That(row2.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Yellow));

            Assert.That(row3.Cell(1).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(row3.Cell(2).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
            Assert.That(row3.Cell(3).Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

            Assert.That(row2.Cell(2).GetString(), Is.EqualTo("X"));
        }

        [Test]
        public void InsertingRowsAbove4()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Row(2).Height = 15;
            ws.Row(3).Height = 20;
            ws.Row(4).Height = 25;
            ws.Row(5).Height = 35;

            ws.Row(2).FirstCell().SetValue("Row height: 15");
            ws.Row(3).FirstCell().SetValue("Row height: 20");
            ws.Row(4).FirstCell().SetValue("Row height: 25");
            ws.Row(5).FirstCell().SetValue("Row height: 35");

            ws.Range("3:3").InsertRowsAbove(1);

            Assert.That(ws.Row(2).Height, Is.EqualTo(15));
            Assert.That(ws.Row(4).Height, Is.EqualTo(20));
            Assert.That(ws.Row(5).Height, Is.EqualTo(25));
            Assert.That(ws.Row(6).Height, Is.EqualTo(35));

            Assert.That(ws.Row(3).Height, Is.EqualTo(20));
            ws.Row(3).ClearHeight();
            Assert.That(ws.Row(3).Height, Is.EqualTo(ws.RowHeight));
        }

        [Test]
        public void NoRowsUsed()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            var count = 0;

            foreach (var row in ws.RowsUsed())
            {
                count++;
            }

            foreach (var row in ws.Range("A1:C3").RowsUsed())
            {
                count++;
            }

            Assert.That(count, Is.EqualTo(0));
        }

        [Test]
        public void RowUsed()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 2).SetValue("Test");
            ws.Cell(1, 3).SetValue("Test");

            var fromRow = ws.Row(1).RowUsed();
            Assert.That(fromRow.RangeAddress.ToStringRelative(), Is.EqualTo("B1:C1"));

            var fromRange = ws.Range("A1:E1").FirstRow().RowUsed();
            Assert.That(fromRange.RangeAddress.ToStringRelative(), Is.EqualTo("B1:C1"));
        }

        [Test]
        public void RowsUsedWithDataValidation()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().SetValue("Hello world!");
            ws.Range("A1:A100").CreateDataValidation().WholeNumber.EqualTo(1);

            var range = ws.Column(1).AsRange();

            Assert.That(range.RowsUsed(XLCellsUsedOptions.DataValidation).Count(), Is.EqualTo(100));
            Assert.That(range.RowsUsed(XLCellsUsedOptions.All).Count(), Is.EqualTo(100));
        }

        [Test]
        public void RowsUsedWithConditionalFormatting()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().SetValue("Hello world!");
            ws.Range("A1:A100").AddConditionalFormat().WhenStartsWith("Hell").Fill.SetBackgroundColor(XLColor.Red).Font.SetFontColor(XLColor.White);

            var range = ws.Column(1).AsRange();

            Assert.That(range.RowsUsed(XLCellsUsedOptions.ConditionalFormats).Count(), Is.EqualTo(100));
            Assert.That(range.RowsUsed(XLCellsUsedOptions.All).Count(), Is.EqualTo(100));
        }

        [Test]
        public void UngroupFromAll()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.AddWorksheet("Sheet1");
            ws.Rows(1, 2).Group();
            ws.Rows(1, 2).Ungroup(true);
        }

        [Test]
        public void NegativeRowNumberIsInvalid()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.AddWorksheet("Sheet1") as XLWorksheet;

            var row = new XLRow(ws, -1);

            Assert.That(row.RangeAddress.IsValid, Is.False);
        }

        [Test]
        public void DeleteRowOnWorksheetWithComment()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.AddWorksheet();
            ws.Cell(4, 1).GetComment().AddText("test");
            ws.Column(1).Width = 100;
            Assert.DoesNotThrow(() => ws.Row(1).Delete());
        }
    }
}