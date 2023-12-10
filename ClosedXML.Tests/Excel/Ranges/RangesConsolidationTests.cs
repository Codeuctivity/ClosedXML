using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.Ranges
{
    [TestFixture]
    public class RangesConsolidationTests
    {
        [Test]
        public void ConsolidateRangesSameWorksheet()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            var ranges = new XLRanges
            {
                ws.Range("A1:E3"),
                ws.Range("A4:B10"),
                ws.Range("E2:F12"),
                ws.Range("C6:I8"),
                ws.Range("G9:G9"),
                ws.Range("C9:D9"),
                ws.Range("H9:H9"),
                ws.Range("I9:I13"),
                ws.Range("C4:D5")
            };

            var consRanges = ranges.Consolidate().ToList();

            Assert.That(consRanges.Count, Is.EqualTo(6));
            Assert.That(consRanges[0].RangeAddress.ToString(), Is.EqualTo("A1:E9"));
            Assert.That(consRanges[1].RangeAddress.ToString(), Is.EqualTo("F2:F12"));
            Assert.That(consRanges[2].RangeAddress.ToString(), Is.EqualTo("G6:I9"));
            Assert.That(consRanges[3].RangeAddress.ToString(), Is.EqualTo("A10:B10"));
            Assert.That(consRanges[4].RangeAddress.ToString(), Is.EqualTo("E10:E12"));
            Assert.That(consRanges[5].RangeAddress.ToString(), Is.EqualTo("I10:I13"));
        }

        [Test]
        public void ConsolidateWideRangesSameWorksheet()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            var ranges = new XLRanges
            {
                ws.Row(5),
                ws.Row(7),
                ws.Row(6),
                ws.Column("D"),
                ws.Column("F"),
                ws.Column("E")
            };

            var consRanges = ranges.Consolidate()
                .OrderBy(r => r.Worksheet.Name)
                .ThenBy(r => r.RangeAddress.FirstAddress.RowNumber)
                .ThenBy(r => r.RangeAddress.FirstAddress.ColumnNumber)
                .ToList();

            Assert.That(consRanges.Count, Is.EqualTo(3));
            Assert.That(consRanges[0].RangeAddress.ToString(), Is.EqualTo("D:F"));
            Assert.That(consRanges[1].RangeAddress.ToString(), Is.EqualTo("A5:C7"));
            Assert.That(consRanges[2].RangeAddress.ToString(), Is.EqualTo("G5:XFD7"));
        }

        [Test]
        public void ConsolidateRangesDifferentWorksheets()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            var ws2 = wb.Worksheets.Add("Sheet2");
            var ranges = new XLRanges
            {
                ws1.Range("A1:E3"),
                ws1.Range("A4:B10"),
                ws1.Range("E2:F12"),
                ws1.Range("C6:I8"),
                ws1.Range("G9:G9"),

                ws2.Row(5),
                ws2.Row(7),
                ws2.Row(6),
                ws2.Column("D"),
                ws2.Column("F"),
                ws2.Column("E"),

                ws1.Range("C9:D9"),
                ws1.Range("H9:H9"),
                ws1.Range("I9:I13"),
                ws1.Range("C4:D5")
            };

            var consRanges = ranges.Consolidate()
                .OrderBy(r => r.Worksheet.Name)
                .ThenBy(r => r.RangeAddress.FirstAddress.RowNumber)
                .ThenBy(r => r.RangeAddress.FirstAddress.ColumnNumber)
                .ToList();

            Assert.That(consRanges.Count, Is.EqualTo(9));
            Assert.That(consRanges[0].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("Sheet1!$A$1:$E$9"));
            Assert.That(consRanges[1].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("Sheet1!$F$2:$F$12"));
            Assert.That(consRanges[2].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("Sheet1!$G$6:$I$9"));
            Assert.That(consRanges[3].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("Sheet1!$A$10:$B$10"));
            Assert.That(consRanges[4].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("Sheet1!$E$10:$E$12"));
            Assert.That(consRanges[5].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("Sheet1!$I$10:$I$13"));

            Assert.That(consRanges[6].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("Sheet2!$D:$F"));
            Assert.That(consRanges[7].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("Sheet2!$A$5:$C$7"));
            Assert.That(consRanges[8].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true), Is.EqualTo("Sheet2!$G$5:$XFD$7"));
        }

        [Test]
        public void ConsolidateSparsedRanges()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            var ranges = new XLRanges
            {
                ws.Range("A1:C1"),
                ws.Range("E1:G1"),
                ws.Range("A3:C3"),
                ws.Range("E3:G3")
            };

            var consRanges = ranges.Consolidate().ToList();

            Assert.That(consRanges.Count, Is.EqualTo(4));
            Assert.That(consRanges[0].RangeAddress.ToString(), Is.EqualTo("A1:C1"));
            Assert.That(consRanges[1].RangeAddress.ToString(), Is.EqualTo("E1:G1"));
            Assert.That(consRanges[2].RangeAddress.ToString(), Is.EqualTo("A3:C3"));
            Assert.That(consRanges[3].RangeAddress.ToString(), Is.EqualTo("E3:G3"));
        }
    }
}