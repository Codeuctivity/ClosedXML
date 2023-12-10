using ClosedXML.Excel;
using ClosedXML.Excel.Ranges;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Tests.Excel.Ranges
{
    [TestFixture]
    public class XLRangeBaseTests
    {
        [Test]
        public void IsEmpty1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            _ = ws.Cell(1, 1);
            var range = ws.Range("A1:B2");
            var actual = range.IsEmpty();
            var expected = true;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            _ = ws.Cell(1, 1);
            var range = ws.Range("A1:B2");
            var actual = range.IsEmpty(XLCellsUsedOptions.All);
            var expected = true;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty3()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            var range = ws.Range("A1:B2");
            var actual = range.IsEmpty();
            var expected = true;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty4()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            var range = ws.Range("A1:B2");
            var actual = range.IsEmpty(XLCellsUsedOptions.AllContents);
            var expected = true;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty5()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            var range = ws.Range("A1:B2");
            var actual = range.IsEmpty(XLCellsUsedOptions.All);
            var expected = false;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty6()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Value = "X";
            var range = ws.Range("A1:B2");
            var actual = range.IsEmpty();
            var expected = false;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void SingleCell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).Value = "Hello World!";
            wb.NamedRanges.Add("SingleCell", "Sheet1!$A$1");
            var range = wb.Range("SingleCell");
            Assert.That(range.CellsUsed().Count(), Is.EqualTo(1));
            Assert.That(range.CellsUsed().Single().GetString(), Is.EqualTo("Hello World!"));
        }

        [Test]
        public void TableRange()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            var rangeColumn = ws.Column(1).Column(1, 4);
            rangeColumn.Cell(1).Value = "FName";
            rangeColumn.Cell(2).Value = "John";
            rangeColumn.Cell(3).Value = "Hank";
            rangeColumn.Cell(4).Value = "Dagny";
            var table = rangeColumn.CreateTable();
            wb.NamedRanges.Add("FNameColumn", string.Format("{0}[{1}]", table.Name, "FName"));

            var namedRange = wb.Range("FNameColumn");
            Assert.That(namedRange.Cells().Count(), Is.EqualTo(3));
            Assert.That(
                namedRange.CellsUsed().Select(cell => cell.GetString()).SequenceEqual(new[] { "John", "Hank", "Dagny" }), Is.True);
        }

        [Test]
        public void WsNamedCell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell", XLScope.Worksheet);
            Assert.That(ws.Cell("TestCell").GetString(), Is.EqualTo("Test"));
        }

        [Test]
        public void WsNamedCells()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("Test").AddToNamed("TestCell", XLScope.Worksheet);
            ws.Cell(2, 1).SetValue("B");
            var cells = ws.Cells("TestCell, A2");
            Assert.That(cells.First().GetString(), Is.EqualTo("Test"));
            Assert.That(cells.Last().GetString(), Is.EqualTo("B"));
        }

        [Test]
        public void WsNamedRange()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("A");
            ws.Cell(2, 1).SetValue("B");
            var original = ws.Range("A1:A2");
            original.AddToNamed("TestRange", XLScope.Worksheet);
            var named = ws.Range("TestRange");
            Assert.That(named.RangeAddress.ToString(), Is.EqualTo(original.RangeAddress.ToStringFixed()));
        }

        [Test]
        public void WsNamedRanges()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell(1, 1).SetValue("A");
            ws.Cell(2, 1).SetValue("B");
            ws.Cell(3, 1).SetValue("C");
            var original = ws.Range("A1:A2");
            original.AddToNamed("TestRange", XLScope.Worksheet);
            var namedRanges = ws.Ranges("TestRange, A3");
            Assert.That(namedRanges.First().RangeAddress.ToString(), Is.EqualTo(original.RangeAddress.ToStringFixed()));
            Assert.That(namedRanges.Last().RangeAddress.ToStringFixed(), Is.EqualTo("$A$3:$A$3"));
        }

        [Test]
        public void WsNamedRangesOneString()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");
            ws.NamedRanges.Add("TestRange", "Sheet1!$A$1,Sheet1!$A$3");
            var namedRanges = ws.Ranges("TestRange");

            Assert.That(namedRanges.First().RangeAddress.ToStringFixed(), Is.EqualTo("$A$1:$A$1"));
            Assert.That(namedRanges.Last().RangeAddress.ToStringFixed(), Is.EqualTo("$A$3:$A$3"));
        }

        //[Test]
        //public void WsNamedRangeLiteral()
        //{
        //using var wb = new XLWorkbook();
        //    var ws = wb.Worksheets.Add("Sheet1");
        //    ws.NamedRanges.Add("TestRange", "\"Hello\"");
        //    using (MemoryStream memoryStream = new MemoryStream())
        //    {
        //        wb.SaveAs(memoryStream, true);
        //        var wb2 = new XLWorkbook(memoryStream);
        //        var text = wb2.Worksheet("Sheet1").NamedRanges.First()
        //        memoryStream.Close();
        //    }

        //}

        [Test]
        public void GrowRange()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            Assert.That(ws.Cell("A1").AsRange().Grow().RangeAddress.ToString(), Is.EqualTo("A1:B2"));
            Assert.That(ws.Cell("A2").AsRange().Grow().RangeAddress.ToString(), Is.EqualTo("A1:B3"));
            Assert.That(ws.Cell("B1").AsRange().Grow().RangeAddress.ToString(), Is.EqualTo("A1:C2"));

            Assert.That(ws.Cell("F5").AsRange().Grow().RangeAddress.ToString(), Is.EqualTo("E4:G6"));
            Assert.That(ws.Cell("F5").AsRange().Grow(2).RangeAddress.ToString(), Is.EqualTo("D3:H7"));
            Assert.That(ws.Cell("F5").AsRange().Grow(100).RangeAddress.ToString(), Is.EqualTo("A1:DB105"));
        }

        [Test]
        public void ShrinkRange()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            Assert.That(ws.Cell("A1").AsRange().Shrink(), Is.Null);
            Assert.That(ws.Range("B2:C3").Shrink(), Is.Null);
            Assert.That(ws.Range("B2:D4").Shrink().RangeAddress.ToString(), Is.EqualTo("C3:C3"));
            Assert.That(ws.Range("A1:Z26").Shrink(10).RangeAddress.ToString(), Is.EqualTo("K11:P16"));

            // Grow and shrink back
            Assert.That(ws.Cell("Z26").AsRange().Grow(10).Shrink(10).RangeAddress.ToString(), Is.EqualTo("Z26:Z26"));
        }

        [Test]
        public void Intersection()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            Assert.That(ws.Range("B9:I11").Intersection(ws.Range("D4:G16")).ToString(), Is.EqualTo("D9:G11"));
            Assert.That(ws.Range("E9:I11").Intersection(ws.Range("D4:G16")).ToString(), Is.EqualTo("E9:G11"));
            Assert.That(ws.Cell("E9").AsRange().Intersection(ws.Range("D4:G16")).ToString(), Is.EqualTo("E9:E9"));
            Assert.That(ws.Range("D4:G16").Intersection(ws.Cell("E9").AsRange()).ToString(), Is.EqualTo("E9:E9"));

            XLRangeAddress rangeAddress;

            rangeAddress = (XLRangeAddress)ws.Cell("C3").AsRange().Intersection(ws.Cell("A1").AsRange());
            Assert.That(rangeAddress.IsValid, Is.False);

            rangeAddress = (XLRangeAddress)ws.Cell("A1").AsRange().Intersection(ws.Cell("C3").AsRange());
            Assert.That(rangeAddress.IsValid, Is.False);

            Assert.That(ws.Range("A1:C3").Intersection(null), Is.Null);

            var otherWs = wb.AddWorksheet("Sheet2");
            Assert.That(ws.Intersection(otherWs), Is.Null);
            Assert.That(ws.Cell("A1").AsRange().Intersection(otherWs.Cell("A2").AsRange()), Is.Null);
        }

        [Test]
        public void Union()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            Assert.That(ws.Range("B9:I11").Union(ws.Range("D4:G16")).Count(), Is.EqualTo(64));
            Assert.That(ws.Range("E9:I11").Union(ws.Range("D4:G16")).Count(), Is.EqualTo(58));
            Assert.That(ws.Cell("E9").AsRange().Union(ws.Range("D4:G16")).Count(), Is.EqualTo(52));
            Assert.That(ws.Range("D4:G16").Union(ws.Cell("E9").AsRange()).Count(), Is.EqualTo(52));

            Assert.That(ws.Cell("A1").AsRange().Union(ws.Cell("C3").AsRange()).Count(), Is.EqualTo(2));

            Assert.That(ws.Range("A1:C3").Union(null).Count(), Is.EqualTo(9));

            var otherWs = wb.AddWorksheet("Sheet2");
            Assert.That(ws.Union(otherWs).Any(), Is.False);
            Assert.That(ws.Cell("A1").AsRange().Union(otherWs.Cell("A2").AsRange()).Any(), Is.False);
        }

        [Test]
        public void Difference()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            Assert.That(ws.Range("B9:I11").Difference(ws.Range("D4:G16")).Count(), Is.EqualTo(12));
            Assert.That(ws.Range("E9:I11").Difference(ws.Range("D4:G16")).Count(), Is.EqualTo(6));
            Assert.That(ws.Cell("E9").AsRange().Difference(ws.Range("D4:G16")).Count(), Is.EqualTo(0));
            Assert.That(ws.Range("D4:G16").Difference(ws.Cell("E9").AsRange()).Count(), Is.EqualTo(51));

            Assert.That(ws.Cell("A1").AsRange().Difference(ws.Cell("C3").AsRange()).Count(), Is.EqualTo(1));

            Assert.That(ws.Range("A1:C3").Difference(null).Count(), Is.EqualTo(9));

            var otherWs = wb.AddWorksheet("Sheet2");
            Assert.That(ws.Difference(otherWs).Any(), Is.False);
            Assert.That(ws.Cell("A1").AsRange().Difference(otherWs.Cell("A2").AsRange()).Any(), Is.False);
        }

        [Test]
        public void SurroundingCells()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            Assert.That(ws.FirstCell().AsRange().SurroundingCells().Count(), Is.EqualTo(3));
            Assert.That(ws.Cell("C3").AsRange().SurroundingCells().Count(), Is.EqualTo(8));
            Assert.That(ws.Range("C3:D6").AsRange().SurroundingCells().Count(), Is.EqualTo(16));

            Assert.That(ws.Range("C3:D6").AsRange().SurroundingCells(c => !c.IsEmpty()).Count(), Is.EqualTo(0));
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeAbove1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("B2:E3").Clear(XLClearOptions.ConditionalFormats);

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
            Assert.That(ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative(), Is.EqualTo("C4:D7"));
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeAbove2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("C3:D3").Clear(XLClearOptions.ConditionalFormats);

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
            Assert.That(ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative(), Is.EqualTo("C4:D7"));
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeBelow1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("B7:E8").Clear(XLClearOptions.ConditionalFormats);

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
            Assert.That(ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative(), Is.EqualTo("C3:D6"));
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeBelow2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("C7:D7").Clear(XLClearOptions.ConditionalFormats);

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
            Assert.That(ws.ConditionalFormats.Single().Range.RangeAddress.ToStringRelative(), Is.EqualTo("C3:D6"));
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeRowInMiddle()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            ws.Range("C3:D7").AddConditionalFormat();
            ws.Range("C5:E5").Clear(XLClearOptions.ConditionalFormats);

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
            Assert.That(ws.ConditionalFormats.First().Ranges.First().RangeAddress.ToStringRelative(), Is.EqualTo("C3:D4"));
            Assert.That(ws.ConditionalFormats.First().Ranges.Last().RangeAddress.ToStringRelative(), Is.EqualTo("C6:D7"));
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeColumnInMiddle()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            ws.Range("C3:G4").AddConditionalFormat();
            ws.Range("E2:E4").Clear(XLClearOptions.ConditionalFormats);

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
            Assert.That(ws.ConditionalFormats.First().Ranges.First().RangeAddress.ToStringRelative(), Is.EqualTo("C3:D4"));
            Assert.That(ws.ConditionalFormats.First().Ranges.Last().RangeAddress.ToStringRelative(), Is.EqualTo("F3:G4"));
        }

        [Test]
        public void ClearConditionalFormattingsWhenRangeContainsFormatWhole()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            ws.Range("C3:G4").AddConditionalFormat();
            ws.Range("B2:G4").Clear(XLClearOptions.ConditionalFormats);

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(0));
        }

        [Test]
        public void NoClearConditionalFormattingsWhenRangePartiallySuperimposed()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            ws.Range("C3:G4").AddConditionalFormat();
            ws.Range("C2:D3").Clear(XLClearOptions.ConditionalFormats);

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
            Assert.That(ws.ConditionalFormats.Single().Ranges.Count, Is.EqualTo(1));
            Assert.That(ws.ConditionalFormats.Single().Ranges.Single().RangeAddress.ToStringRelative(), Is.EqualTo("C3:G4"));
        }

        [Test]
        public void RangesRemoveAllWithoutDispose()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var ranges = new XLRanges
            {
                ws.Range("A1:A2"),
                ws.Range("B1:B2")
            };
            var rangesCopy = ranges.ToList();

            ranges.RemoveAll(null, false);
            ws.FirstColumn().InsertColumnsBefore(1);

            Assert.That(ranges.Count, Is.EqualTo(0));
            // if ranges were not disposed they addresses should change
            Assert.That(rangesCopy.First().RangeAddress.ToString(), Is.EqualTo("B1:B2"));
            Assert.That(rangesCopy.Last().RangeAddress.ToString(), Is.EqualTo("C1:C2"));
        }

        [Test]
        public void RangesRemoveAllByCriteria()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var ranges = new XLRanges
            {
                ws.Range("A1:A2"),
                ws.Range("B1:B3"),
                ws.Range("C1:C4")
            };
            var otherRange = ws.Range("A3:D3");

            ranges.RemoveAll(r => r.Intersects(otherRange));

            Assert.That(ranges.Count, Is.EqualTo(1));
            Assert.That(ranges.Single().RangeAddress.ToString(), Is.EqualTo("A1:A2"));
        }

        [Test]
        public void XLRangesReturnsRangesInDeterministicOrder()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            var ws2 = wb.Worksheets.Add("Another sheet");

            var ranges = new XLRanges
            {
                ws2.Range("F1:F12"),
                ws1.Range("F12:F16"),
                ws1.Range("B1:F2"),
                ws2.Range("A13:B14"),
                ws2.Range("E1:E2"),
                ws1.Range("E1:H2"),
                ws1.Range("G2:G13"),
                ws1.Range("G20:G20")
            };

            var expectedRanges = new List<IXLRange>
            {
                ws1.Range("B1:F2"),
                ws1.Range("E1:H2"),
                ws1.Range("G2:G13"),
                ws1.Range("F12:F16"),
                ws1.Range("G20:G20"),

                ws2.Range("E1:E2"),
                ws2.Range("F1:F12"),
                ws2.Range("A13:B14"),
            };

            var actualRanges = ranges.ToList();

            Assert.That(actualRanges.Count, Is.EqualTo(expectedRanges.Count));
            for (var i = 0; i < actualRanges.Count; i++)
            {
                Assert.That(actualRanges[i], Is.EqualTo(expectedRanges[i]));
            }
        }

        [Test]
        public void ClearRangeRemovesSparklines()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            ws.SparklineGroups.Add("B1:B3", "C1:E3");

            ws.Range("B1:C1").Clear(XLClearOptions.All);
            ws.Range("B2:C2").Clear(XLClearOptions.Sparklines);

            Assert.That(ws.SparklineGroups.Single().Count(), Is.EqualTo(1));
            Assert.That(ws.Cell("B1").HasSparkline, Is.False);
            Assert.That(ws.Cell("B2").HasSparkline, Is.False);
            Assert.That(ws.Cell("B3").HasSparkline, Is.True);
        }

        [TestCase("B2:G7", "D4:E5", true, "B2:G3,B4:C5,D4:E5,F4:G5,B6:G7")]
        [TestCase("B2:G7", "D4:E5", false, "B2:G3,B4:C5,F4:G5,B6:G7")]
        [TestCase("B2:G7", "B2:G7", true, "B2:G7")]
        [TestCase("B2:G7", "B2:G7", false, "")]
        [TestCase("B2:G7", "A1:H8", true, "B2:G7")]
        [TestCase("B2:G7", "A1:H8", false, "")]
        [TestCase("B2:G7", "A1:B2", true, "B2:B2,C2:G2,B3:G7")]
        [TestCase("B2:G7", "A1:B2", false, "C2:G2,B3:G7")]
        [TestCase("B2:G7", "E4:J5", true, "B2:G3,B4:D5,E4:G5,B6:G7")]
        [TestCase("B2:G7", "E4:J5", false, "B2:G3,B4:D5,B6:G7")]
        [TestCase("B2:G7", "A11:H18", true, "B2:G7")]
        [TestCase("B2:G7", "A11:H18", false, "B2:G7")]
        [TestCase("B2:G7", "A1:H1", true, "B2:G7")]
        [TestCase("B2:G7", "A1:A12", true, "B2:G7")]
        [TestCase("B2:G7", "A8:H8", true, "B2:G7")]
        [TestCase("B2:G7", "H1:H8", true, "B2:G7")]
        public void CanSplitRange(string rangeAddress, string splitBy, bool includeIntersection, string expectedResult)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var range = ws.Range(rangeAddress) as XLRange;
            var splitter = ws.Range(splitBy);

            var result = range.Split(splitter.RangeAddress, includeIntersection);

            var actualAddresses = string.Join(",", result.Select(r => r.RangeAddress.ToString()));

            Assert.That(actualAddresses, Is.EqualTo(expectedResult));
        }
    }
}