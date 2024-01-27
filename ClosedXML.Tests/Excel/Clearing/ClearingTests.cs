using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.Clearing
{
    [TestFixture]
    public class ClearingTests
    {
        private static readonly XLColor backgroundColor = XLColor.LightBlue;
        private static readonly XLColor foregroundColor = XLColor.DarkBrown;

        private IXLWorkbook SetupWorkbook(XLWorkbook wb)
        {
            var ws = wb.Worksheets.Add("Sheet1");

            var c = ws.FirstCell()
                .SetValue("Hello world!");

            c.GetComment().AddText("Some comment");

            c.Style.Fill.BackgroundColor = backgroundColor;
            c.Style.Font.FontColor = foregroundColor;
            c.CreateDataValidation().Custom("B1");

            ////

            c = ws.FirstCell()
                .CellBelow()
                .SetFormulaA1("=LEFT(A1,5)");

            c.GetComment().AddText("Another comment");

            c.Style.Fill.BackgroundColor = backgroundColor;
            c.Style.Font.FontColor = foregroundColor;

            ////

            c = ws.FirstCell()
                .CellBelow(2)
                .SetValue(new DateTime(2018, 1, 15));

            c.GetComment().AddText("A date");

            c.Style.Fill.BackgroundColor = backgroundColor;
            c.Style.Font.FontColor = foregroundColor;

            ws.Column(1)
                .AddConditionalFormat().WhenStartsWith("Hell")
                .Fill.SetBackgroundColor(XLColor.Red)
                .Border.SetOutsideBorder(XLBorderStyleValues.Thick)
                .Border.SetOutsideBorderColor(XLColor.Blue)
                .Font.SetBold();

            Assert.That(ws.Cell("A1").DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.Cell("A2").DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.Cell("A3").DataType, Is.EqualTo(XLDataType.DateTime));

            Assert.That(ws.Cell("A1").HasFormula, Is.EqualTo(false));
            Assert.That(ws.Cell("A2").HasFormula, Is.EqualTo(true));
            Assert.That(ws.Cell("A1").HasFormula, Is.EqualTo(false));

            foreach (var cell in ws.Range("A1:A3").Cells())
            {
                Assert.That(cell.Style.Fill.BackgroundColor, Is.EqualTo(backgroundColor));
                Assert.That(cell.Style.Font.FontColor, Is.EqualTo(foregroundColor));
                Assert.That(ws.ConditionalFormats.Any(), Is.True);
                Assert.That(cell.HasComment, Is.True);
            }

            Assert.That(ws.Cell("A1").GetDataValidation().Value, Is.EqualTo("B1"));

            return wb;
        }

        [Test]
        public void WorksheetClearAll()
        {
            using var emptyWb = new XLWorkbook();

            using var wb = SetupWorkbook(emptyWb);
            var ws = wb.Worksheets.First();

            ws.Clear(XLClearOptions.All);

            foreach (var c in ws.Range("A1:A10").Cells())
            {
                Assert.That(c.IsEmpty(), Is.True);
                Assert.That(c.DataType, Is.EqualTo(XLDataType.Text));
                Assert.That(c.Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));
                Assert.That(c.Style.Font.FontColor, Is.EqualTo(ws.Style.Font.FontColor));
                Assert.That(ws.ConditionalFormats.Any(), Is.False);
                Assert.That(c.HasComment, Is.False);
                Assert.That(c.GetDataValidation().Value, Is.EqualTo(string.Empty));
            }
        }

        [Test]
        public void WorksheetClearContents()
        {
            using var emptyWb = new XLWorkbook();

            using var wb = SetupWorkbook(emptyWb);
            var ws = wb.Worksheets.First();

            ws.Clear(XLClearOptions.Contents);

            foreach (var c in ws.Range("A1:A3").Cells())
            {
                Assert.That(c.IsEmpty(XLCellsUsedOptions.Contents), Is.True);
                Assert.That(c.Style.Fill.BackgroundColor, Is.EqualTo(backgroundColor));
                Assert.That(c.Style.Font.FontColor, Is.EqualTo(foregroundColor));
                Assert.That(ws.ConditionalFormats.Any(), Is.True);
                Assert.That(c.HasComment, Is.True);
            }

            Assert.That(ws.Cell("A1").GetDataValidation().Value, Is.EqualTo("B1"));

            Assert.That(ws.Cell("A1").DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.Cell("A2").DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.Cell("A3").DataType, Is.EqualTo(XLDataType.DateTime));
        }

        [Test]
        public void WorksheetClearDataType()
        {
            using var emptyWb = new XLWorkbook();

            using var wb = SetupWorkbook(emptyWb);
            var ws = wb.Worksheets.First();

            ws.Clear(XLClearOptions.DataType);

            foreach (var c in ws.Range("A1:A3").Cells())
            {
                Assert.That(c.IsEmpty(), Is.False);
                Assert.That(c.DataType, Is.EqualTo(XLDataType.Text));
                Assert.That(c.Style.Fill.BackgroundColor, Is.EqualTo(backgroundColor));
                Assert.That(c.Style.Font.FontColor, Is.EqualTo(foregroundColor));
                Assert.That(ws.ConditionalFormats.Any(), Is.True);
                Assert.That(c.HasComment, Is.True);
            }

            Assert.That(ws.Cell("A1").GetDataValidation().Value, Is.EqualTo("B1"));
        }

        [Test]
        public void WorksheetClearNormalFormats()
        {
            using var emptyWb = new XLWorkbook();

            using var wb = SetupWorkbook(emptyWb);
            var ws = wb.Worksheets.First();

            ws.Clear(XLClearOptions.NormalFormats);

            foreach (var c in ws.Range("A1:A3").Cells())
            {
                Assert.That(c.IsEmpty(), Is.False);
                Assert.That(c.Style.Fill.BackgroundColor, Is.EqualTo(ws.Style.Fill.BackgroundColor));
                Assert.That(c.Style.Font.FontColor, Is.EqualTo(ws.Style.Font.FontColor));
                Assert.That(ws.ConditionalFormats.Any(), Is.True);
                Assert.That(c.HasComment, Is.True);
            }

            Assert.That(ws.Cell("A1").DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.Cell("A2").DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.Cell("A3").DataType, Is.EqualTo(XLDataType.DateTime));

            Assert.That(ws.Cell("A1").GetDataValidation().Value, Is.EqualTo("B1"));
        }

        [Test]
        public void WorksheetClearConditionalFormats()
        {
            using var emptyWb = new XLWorkbook();

            using var wb = SetupWorkbook(emptyWb);
            var ws = wb.Worksheets.First();

            ws.Clear(XLClearOptions.ConditionalFormats);

            foreach (var c in ws.Range("A1:A3").Cells())
            {
                Assert.That(c.IsEmpty(), Is.False);
                Assert.That(c.Style.Fill.BackgroundColor, Is.EqualTo(backgroundColor));
                Assert.That(c.Style.Font.FontColor, Is.EqualTo(foregroundColor));
                Assert.That(ws.ConditionalFormats.Any(), Is.False);
                Assert.That(c.HasComment, Is.True);
            }

            Assert.That(ws.Cell("A1").DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.Cell("A2").DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.Cell("A3").DataType, Is.EqualTo(XLDataType.DateTime));

            Assert.That(ws.Cell("A1").GetDataValidation().Value, Is.EqualTo("B1"));
        }

        [Test]
        public void WorksheetClearComments()
        {
            using var emptyWb = new XLWorkbook();

            using var wb = SetupWorkbook(emptyWb);
            var ws = wb.Worksheets.First();

            ws.Clear(XLClearOptions.Comments);

            foreach (var c in ws.Range("A1:A3").Cells())
            {
                Assert.That(c.IsEmpty(), Is.False);
                Assert.That(c.Style.Fill.BackgroundColor, Is.EqualTo(backgroundColor));
                Assert.That(c.Style.Font.FontColor, Is.EqualTo(foregroundColor));
                Assert.That(ws.ConditionalFormats.Any(), Is.True);
                Assert.That(c.HasComment, Is.False);
            }

            Assert.That(ws.Cell("A1").DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.Cell("A2").DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.Cell("A3").DataType, Is.EqualTo(XLDataType.DateTime));

            Assert.That(ws.Cell("A1").GetDataValidation().Value, Is.EqualTo("B1"));
        }

        [Test]
        public void WorksheetClearDataValidation()
        {
            using var emptyWb = new XLWorkbook();
            using var wb = SetupWorkbook(emptyWb);
            var ws = wb.Worksheets.First();

            ws.Clear(XLClearOptions.DataValidation);

            foreach (var c in ws.Range("A1:A3").Cells())
            {
                Assert.That(c.IsEmpty(), Is.False);
                Assert.That(c.Style.Fill.BackgroundColor, Is.EqualTo(backgroundColor));
                Assert.That(c.Style.Font.FontColor, Is.EqualTo(foregroundColor));
                Assert.That(ws.ConditionalFormats.Any(), Is.True);
                Assert.That(c.HasComment, Is.True);
            }

            Assert.That(ws.Cell("A1").DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.Cell("A2").DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.Cell("A3").DataType, Is.EqualTo(XLDataType.DateTime));

            Assert.That(ws.Cell("A1").GetDataValidation().Value, Is.EqualTo(string.Empty));
        }

        [Test]
        public void DeleteClearedCellValue()
        {
            using var emptyWb = new XLWorkbook();
            using var ms = new MemoryStream();
            using (var wb = SetupWorkbook(emptyWb))
            {
                var ws = wb.Worksheets.First();
                Assert.That(ws.Cell("A1").GetString(), Is.EqualTo("Hello world!"));
                Assert.That(ws.Cell("A3").GetDateTime(), Is.EqualTo(new DateTime(2018, 1, 15)));

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                ws.Clear(XLClearOptions.Contents);
                Assert.That(ws.Cell("A1").GetString(), Is.EqualTo(""));
                Assert.Throws<FormatException>(() => ws.Cell("A3").GetDateTime());

                wb.Save();
            }

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                Assert.That(ws.Cell("A1").GetString(), Is.EqualTo(""));
                Assert.Throws<FormatException>(() => ws.Cell("A3").GetDateTime());
            }
        }

        [TestCase(XLClearOptions.All, 2)]
        [TestCase(XLClearOptions.AllContents, 4)]
        [TestCase(XLClearOptions.AllFormats, 4)]
        [TestCase(XLClearOptions.Contents, 4)]
        [TestCase(XLClearOptions.MergedRanges, 2)]
        public void CanClearMergedRanges(XLClearOptions options, int expectedCount)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Test");

            ws.Range("A1:C3").Merge();
            ws.Range("A4:B6").Merge();
            ws.Range("D1:F3").Merge();
            ws.Range("E4:F6").Merge();

            ws.Range("C1:D6").Clear(options);

            Assert.That(ws.MergedRanges.Count, Is.EqualTo(expectedCount));
        }
    }
}