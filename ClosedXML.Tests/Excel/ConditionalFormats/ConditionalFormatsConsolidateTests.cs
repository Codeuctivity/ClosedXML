﻿using ClosedXML.Excel;
using ClosedXML.Excel.Ranges;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.ConditionalFormats
{
    [TestFixture]
    public class ConditionalFormatsConsolidateTests
    {
        [Test]
        public void ConsecutivelyRowsConsolidateTest()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("B2:C2").AddConditionalFormat());
            SetFormat1(ws.Range("B4:C4").AddConditionalFormat());
            SetFormat1(ws.Range("B3:C3").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
            var format = ws.ConditionalFormats.First();
            Assert.That(format.Range.RangeAddress.ToStringRelative(), Is.EqualTo("B2:C4"));
            Assert.That(format.Values.Values.First().Value, Is.EqualTo("F2"));
        }

        [Test]
        public void ConsecutivelyColumnsConsolidateTest()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("D2:D3").AddConditionalFormat());
            SetFormat1(ws.Range("B2:B3").AddConditionalFormat());
            SetFormat1(ws.Range("C2:C3").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
            var format = ws.ConditionalFormats.First();
            Assert.That(format.Ranges.First().RangeAddress.ToStringRelative(), Is.EqualTo("B2:D3"));
            Assert.That(format.Values.Values.First().Value, Is.EqualTo("F2"));
        }

        [Test]
        public void Contains1ConsolidateTest()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("B11:D12").AddConditionalFormat());
            SetFormat1(ws.Range("C12:D12").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
            var format = ws.ConditionalFormats.First();
            Assert.That(format.Range.RangeAddress.ToStringRelative(), Is.EqualTo("B11:D12"));
            Assert.That(format.Values.Values.First().Value, Is.EqualTo("F11"));
        }

        [Test]
        public void Contains2ConsolidateTest()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("B14:C14").AddConditionalFormat());
            SetFormat1(ws.Range("B14:B14").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
            var format = ws.ConditionalFormats.First();
            Assert.That(format.Range.RangeAddress.ToStringRelative(), Is.EqualTo("B14:C14"));
            Assert.That(format.Values.Values.First().Value, Is.EqualTo("F14"));
        }

        [Test]
        public void SuperimposedConsolidateTest()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("B16:D18").AddConditionalFormat());
            SetFormat1(ws.Range("B18:D19").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
            var format = ws.ConditionalFormats.First();
            Assert.That(format.Range.RangeAddress.ToStringRelative(), Is.EqualTo("B16:D19"));
            Assert.That(format.Values.Values.First().Value, Is.EqualTo("F16"));
        }

        [Test]
        public void DifferentFormatNoConsolidateTest()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("B11:D12").AddConditionalFormat());
            SetFormat2(ws.Range("C12:D12").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(2));
        }

        [Test]
        public void ConsolidatePreservesPriorities()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("A1:A5").AddConditionalFormat());
            SetFormat2(ws.Range("A1:A5").AddConditionalFormat());
            SetFormat2(ws.Range("A6:A10").AddConditionalFormat());
            SetFormat1(ws.Range("A6:A10").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(3));
            Assert.That((ws.ConditionalFormats.Last().Style as XLStyle).Value, Is.EqualTo((ws.ConditionalFormats.First().Style as XLStyle).Value));
            Assert.That((ws.ConditionalFormats.ElementAt(1).Style as XLStyle).Value, Is.Not.EqualTo((ws.ConditionalFormats.First().Style as XLStyle).Value));
        }

        [Test]
        public void ConsolidatePreservesPriorities2()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");

            SetFormat1(ws.Range("A1:A1").AddConditionalFormat());
            SetFormat2(ws.Range("A2:A3").AddConditionalFormat());
            SetFormat1(ws.Range("A2:A6").AddConditionalFormat());
            SetFormat1(ws.Range("A7:A8").AddConditionalFormat());

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(3));
            Assert.That((ws.ConditionalFormats.Last().Style as XLStyle).Value, Is.EqualTo((ws.ConditionalFormats.First().Style as XLStyle).Value));
            Assert.That((ws.ConditionalFormats.ElementAt(1).Style as XLStyle).Value, Is.Not.EqualTo((ws.ConditionalFormats.First().Style as XLStyle).Value));
            Assert.That(ws.ConditionalFormats.All(cf => cf.Ranges.Count == 1), Is.True, "Number of ranges in consolidated conditional formats is expected to be 1");
            Assert.That(ws.ConditionalFormats.ElementAt(0).Ranges.Single().RangeAddress.ToString(), Is.EqualTo("A1:A1"));
            Assert.That(ws.ConditionalFormats.ElementAt(1).Ranges.Single().RangeAddress.ToString(), Is.EqualTo("A2:A3"));
            Assert.That(ws.ConditionalFormats.ElementAt(2).Ranges.Single().RangeAddress.ToString(), Is.EqualTo("A2:A8"));
        }

        [Test]
        public void ConsolidateShiftsFormulaRelativelyToTopMostCell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");

            var ranges = ws.Ranges("B3:B8,C3:C4,A3:A4,C5:C8,A5:A8").Cast<XLRange>();
            var cf = new XLConditionalFormat(ranges);
            cf.Values.Add(new XLFormula("=A3=$D3"));
            cf.Style.Fill.SetBackgroundColor(XLColor.Red);
            ws.ConditionalFormats.Add(cf);

            ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

            Assert.That(ws.ConditionalFormats.Count(), Is.EqualTo(1));
            Assert.That((cf.Style as XLStyle).Value, Is.EqualTo((ws.ConditionalFormats.Single().Style as XLStyle).Value));
            Assert.That(ws.ConditionalFormats.Single().Ranges.Single().RangeAddress.ToString(), Is.EqualTo("A3:C8"));
            Assert.That(ws.ConditionalFormats.Single().Values.Single().Value.IsFormula, Is.True);
            Assert.That(ws.ConditionalFormats.Single().Values.Single().Value.Value, Is.EqualTo("A3=$D3"));
        }

        [Test]
        public void ColorScaleComparing()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet");

            var ranges = ws.Ranges("B3:B8,C3:C4,A3:A4,C5:C8,A5:A8").Cast<XLRange>();
            var cf1 = new XLConditionalFormat(ranges);
            cf1.ColorScale()
                .LowestValue(XLColor.Red)
                .HighestValue(XLColor.Green);

            var cf2 = new XLConditionalFormat(ranges);
            cf2.ColorScale()
                .LowestValue(XLColor.Red)
                .HighestValue(XLColor.Green);
            Assert.That(XLConditionalFormat.NoRangeComparer.Equals(cf1, cf2), Is.True);
        }

        private static void SetFormat1(IXLConditionalFormat format)
        {
            format.WhenEquals("=" + format.Range.FirstCell().CellRight(4).Address.ToStringRelative()).Fill.SetBackgroundColor(XLColor.Blue);
        }

        private static void SetFormat2(IXLConditionalFormat format)
        {
            format.WhenEquals(5).Fill.SetBackgroundColor(XLColor.AliceBlue);
        }
    }
}