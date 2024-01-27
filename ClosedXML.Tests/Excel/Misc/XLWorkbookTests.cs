using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.Misc
{
    [TestFixture]
    public class XLWorkbookTests
    {
        [Test]
        public void Cell1()
        {
            using var wb = new XLWorkbook();
            var cell = wb.Cell("ABC");
            Assert.That(cell, Is.Null);
        }

        [Test]
        public void Cell2()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
            var cell = wb.Cell("Sheet1!Result");
            Assert.That(cell, Is.Not.Null);
            Assert.That(cell.GetValue<int>(), Is.EqualTo(1));
        }

        [Test]
        public void Cell3()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result");
            var cell = wb.Cell("Sheet1!Result");
            Assert.That(cell, Is.Not.Null);
            Assert.That(cell.GetValue<int>(), Is.EqualTo(1));
        }

        [Test]
        public void Cells1()
        {
            using var wb = new XLWorkbook();
            var cells = wb.Cells("ABC");
            Assert.That(cells, Is.Not.Null);
            Assert.That(cells.Count(), Is.EqualTo(0));
        }

        [Test]
        public void Cells2()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
            var cells = wb.Cells("Sheet1!Result, ABC");
            Assert.That(cells, Is.Not.Null);
            Assert.That(cells.Count(), Is.EqualTo(1));
            Assert.That(cells.First().GetValue<int>(), Is.EqualTo(1));
        }

        [Test]
        public void Cells3()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result");
            var cells = wb.Cells("Sheet1!Result, ABC");
            Assert.That(cells, Is.Not.Null);
            Assert.That(cells.Count(), Is.EqualTo(1));
            Assert.That(cells.First().GetValue<int>(), Is.EqualTo(1));
        }

        [Test]
        public void GetCellFromFullAddress()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var ws2 = wb.AddWorksheet("O'Sheet 2");
            var c1 = ws.Cell("C123");
            var c2 = ws2.Cell("B7");

            var c1_full = wb.Cell("Sheet1!C123");
            var c2_full = wb.Cell("'O'Sheet 2'!B7");

            Assert.That(c1_full, Is.SameAs(c1));
            Assert.That(c2_full, Is.SameAs(c2));
            Assert.That(c1_full, Is.Not.Null);
            Assert.That(c2_full, Is.Not.Null);
        }

        [TestCase("Sheet1")]
        [TestCase("Sheet1!")]
        [TestCase("Sheet2!")]
        [TestCase("Sheet2!C1")]
        [TestCase("Sheet1!ZZZ1")]
        [TestCase("Sheet1!A")]
        public void GetCellFromNonExistingFullAddress(string address)
        {
            using var wb = new XLWorkbook();
            wb.AddWorksheet("Sheet1");

            var c = wb.Cell(address);

            Assert.That(c, Is.Null);
        }

        [Test]
        public void GetRangeFromFullAddress()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var r1 = ws.Range("C123:D125");

            var r2 = wb.Range("Sheet1!C123:D125");

            Assert.That(r2, Is.SameAs(r1));
            Assert.That(r2, Is.Not.Null);
        }

        [TestCase("Sheet2!C1:D2")]
        [TestCase("Sheet1!A")]
        public void GetRangeFromNonExistingFullAddress(string rangeAddress)
        {
            using var wb = new XLWorkbook();
            wb.AddWorksheet("Sheet1");

            var r = wb.Range(rangeAddress);

            Assert.That(r, Is.Null);
        }

        [Test]
        public void GetRangesFromFullAddress()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var r1 = ws.Ranges("A1:B2,C1:E3");

            var r2 = wb.Ranges("Sheet1!A1:B2,Sheet1!C1:E3");

            Assert.That(r2.Count, Is.EqualTo(2));
            Assert.That(r2.First(), Is.SameAs(r1.First()));
            Assert.That(r2.Last(), Is.SameAs(r1.Last()));
        }

        [TestCase("Sheet2!C1:D2,Sheet2!F1:G4")]
        [TestCase("Sheet1!A,Sheet1!B")]
        public void GetRangesFromNonExistingFullAddress(string rangesAddress)
        {
            using var wb = new XLWorkbook();
            wb.AddWorksheet("Sheet1");

            var r = wb.Ranges(rangesAddress);

            Assert.That(r, Is.Not.Null);
            Assert.That(r.Any(), Is.False);
        }

        [Test]
        public void NamedRange1()
        {
            using var wb = new XLWorkbook();
            var range = wb.NamedRange("ABC");
            Assert.That(range, Is.Null);
        }

        [Test]
        public void NamedRange2()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
            var range = wb.NamedRange("Sheet1!Result");
            Assert.That(range, Is.Not.Null);
            Assert.That(range.Ranges.Count, Is.EqualTo(1));
            Assert.That(range.Ranges.Cells().Count(), Is.EqualTo(1));
            Assert.That(range.Ranges.First().FirstCell().GetValue<int>(), Is.EqualTo(1));
        }

        [Test]
        public void NamedRange3()
        {
            using var wb = new XLWorkbook();
            wb.AddWorksheet("Sheet1");
            var range = wb.NamedRange("Sheet1!Result");
            Assert.That(range, Is.Null);
        }

        [Test]
        public void NamedRange4()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result");
            var range = wb.NamedRange("Sheet1!Result");
            Assert.That(range, Is.Not.Null);
            Assert.That(range.Ranges.Count, Is.EqualTo(1));
            Assert.That(range.Ranges.Cells().Count(), Is.EqualTo(1));
            Assert.That(range.Ranges.First().FirstCell().GetValue<int>(), Is.EqualTo(1));
        }

        [Test]
        public void Range1()
        {
            using var wb = new XLWorkbook();
            var range = wb.Range("ABC");
            Assert.That(range, Is.Null);
        }

        [Test]
        public void Range2()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
            var range = wb.Range("Sheet1!Result");
            Assert.That(range, Is.Not.Null);
            Assert.That(range.Cells().Count(), Is.EqualTo(1));
            Assert.That(range.FirstCell().GetValue<int>(), Is.EqualTo(1));
        }

        [Test]
        public void Range3()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result");
            var range = wb.Range("Sheet1!Result");
            Assert.That(range, Is.Not.Null);
            Assert.That(range.Cells().Count(), Is.EqualTo(1));
            Assert.That(range.FirstCell().GetValue<int>(), Is.EqualTo(1));
        }

        [Test]
        public void Ranges1()
        {
            using var wb = new XLWorkbook();
            var ranges = wb.Ranges("ABC");
            Assert.That(ranges, Is.Not.Null);
            Assert.That(ranges.Count, Is.EqualTo(0));
        }

        [Test]
        public void Ranges2()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result", XLScope.Worksheet);
            var ranges = wb.Ranges("Sheet1!Result, ABC");
            Assert.That(ranges, Is.Not.Null);
            Assert.That(ranges.Cells().Count(), Is.EqualTo(1));
            Assert.That(ranges.First().FirstCell().GetValue<int>(), Is.EqualTo(1));
        }

        [Test]
        public void Ranges3()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue(1).AddToNamed("Result");
            var ranges = wb.Ranges("Sheet1!Result, ABC");
            Assert.That(ranges, Is.Not.Null);
            Assert.That(ranges.Cells().Count(), Is.EqualTo(1));
            Assert.That(ranges.First().FirstCell().GetValue<int>(), Is.EqualTo(1));
        }

        [Test]
        public void WbNamedCell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell");
            Assert.That(wb.Cell("TestCell").GetString(), Is.EqualTo("Test"));
            Assert.That(ws.Cell("TestCell").GetString(), Is.EqualTo("Test"));
        }

        [Test]
        public void WbNamedCells()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell");
            ws.Cell(2, 1).SetValue("B").AddToNamed("Test2");
            var wbCells = wb.Cells("TestCell, Test2");
            Assert.That(wbCells.First().GetString(), Is.EqualTo("Test"));
            Assert.That(wbCells.Last().GetString(), Is.EqualTo("B"));

            var wsCells = ws.Cells("TestCell, Test2");
            Assert.That(wsCells.First().GetString(), Is.EqualTo("Test"));
            Assert.That(wsCells.Last().GetString(), Is.EqualTo("B"));
        }

        [Test]
        public void WbNamedRange()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("A");
            ws.Cell(2, 1).SetValue("B");
            var original = ws.Range("A1:A2");
            original.AddToNamed("TestRange");
            Assert.That(wb.Range("TestRange").RangeAddress.ToString(), Is.EqualTo(original.RangeAddress.ToStringFixed()));
            Assert.That(ws.Range("TestRange").RangeAddress.ToString(), Is.EqualTo(original.RangeAddress.ToStringFixed()));
        }

        [Test]
        public void WbNamedRanges()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("A");
            ws.Cell(2, 1).SetValue("B");
            ws.Cell(3, 1).SetValue("C").AddToNamed("Test2");
            var original = ws.Range("A1:A2");
            original.AddToNamed("TestRange");
            var wbRanges = wb.Ranges("TestRange, Test2");
            Assert.That(wbRanges.First().RangeAddress.ToString(), Is.EqualTo(original.RangeAddress.ToStringFixed()));
            Assert.That(wbRanges.Last().RangeAddress.ToStringFixed(), Is.EqualTo("$A$3:$A$3"));

            var wsRanges = wb.Ranges("TestRange, Test2");
            Assert.That(wsRanges.First().RangeAddress.ToString(), Is.EqualTo(original.RangeAddress.ToStringFixed()));
            Assert.That(wsRanges.Last().RangeAddress.ToStringFixed(), Is.EqualTo("$A$3:$A$3"));
        }

        [Test]
        public void WbNamedRangesOneString()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            wb.NamedRanges.Add("TestRange", "Sheet1!$A$1,Sheet1!$A$3");

            var wbRanges = ws.Ranges("TestRange");
            Assert.That(wbRanges.First().RangeAddress.ToStringFixed(), Is.EqualTo("$A$1:$A$1"));
            Assert.That(wbRanges.Last().RangeAddress.ToStringFixed(), Is.EqualTo("$A$3:$A$3"));

            var wsRanges = ws.Ranges("TestRange");
            Assert.That(wsRanges.First().RangeAddress.ToStringFixed(), Is.EqualTo("$A$1:$A$1"));
            Assert.That(wsRanges.Last().RangeAddress.ToStringFixed(), Is.EqualTo("$A$3:$A$3"));
        }

        [Test]
        public void WbProtect1()
        {
            using var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet1");
            wb.Protect();
            Assert.That(wb.LockStructure, Is.True);
            Assert.That(wb.LockWindows, Is.False);
            Assert.That(wb.IsPasswordProtected, Is.False);
        }

        [Test]
        public void WbProtect2()
        {
            using var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet1");
#pragma warning disable CS0618 // Type or member is obsolete, but still should be tested
            wb.Protect(true, false);
#pragma warning restore CS0618 // Type or member is obsolete, but still should be tested
            Assert.That(wb.LockStructure, Is.True);
            Assert.That(wb.LockWindows, Is.False);
            Assert.That(wb.IsPasswordProtected, Is.False);
        }

        [Test]
        public void WbProtect3()
        {
            using var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet1");
            wb.Protect("Abc@123");
            Assert.That(wb.LockStructure, Is.True);
            Assert.That(wb.LockWindows, Is.False);
            Assert.That(wb.IsPasswordProtected, Is.True);
            Assert.Throws<InvalidOperationException>(() => wb.Protect());
            Assert.Throws<InvalidOperationException>(() => wb.Unprotect());
            Assert.Throws<ArgumentException>(() => wb.Unprotect("Cde@345"));
        }

        [Test]
        public void WbProtect4()
        {
            using var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet1");
            wb.Protect();
            Assert.That(wb.LockStructure, Is.True);
            Assert.That(wb.LockWindows, Is.False);
            Assert.That(wb.IsPasswordProtected, Is.False);
            wb.Unprotect();
            wb.Protect("Abc@123");
            Assert.That(wb.LockStructure, Is.True);
            Assert.That(wb.LockWindows, Is.False);
            Assert.That(wb.IsPasswordProtected, Is.True);
        }

        [Test]
        public void WbProtect5()
        {
            using var wb = new XLWorkbook();
            wb.Worksheets.Add("Sheet1");
#pragma warning disable CS0618 // Type or member is obsolete, but still should be tested
            wb.Protect(true, false, "Abc@123");
#pragma warning restore CS0618 // Type or member is obsolete, but still should be tested
            Assert.That(wb.LockStructure, Is.True);
            Assert.That(wb.LockWindows, Is.False);
            Assert.That(wb.IsPasswordProtected, Is.True);
            wb.Unprotect("Abc@123");
            Assert.That(wb.LockStructure, Is.False);
            Assert.That(wb.LockWindows, Is.False);
            Assert.That(wb.IsPasswordProtected, Is.False);
        }

        [Test]
        public void FileSharingProperties()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                wb.AddWorksheet("Sheet1").Cell("A1").Value = "Hello world!";
                wb.FileSharing.ReadOnlyRecommended = true;
                wb.FileSharing.UserName = Environment.UserName;
                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                Assert.That(wb.FileSharing.ReadOnlyRecommended, Is.True);
                Assert.That(wb.FileSharing.UserName, Is.EqualTo(Environment.UserName));
            }
        }

        [Test]
        public void AccessDisposedWorkbookThrowsException()
        {
            IXLWorkbook wb;
            using (wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                ws.FirstCell().SetValue("Hello world");
            }

            Assert.Throws<ObjectDisposedException>(() => Console.WriteLine(wb.Worksheets.First().FirstCell().Value));
        }
    }
}