using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Ranges
{
    public class RangeShiftingTests
    {
        [Test]
        public void CellReferenceRemainAfterColumnDeleted()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var d4 = ws.Cell("D4");

            ws.Column("C").Delete();

            Assert.That(ws.Cell("C4"), Is.SameAs(d4));
        }

        [Test]
        public void CellReferenceRemainAfterRowDeleted()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var d4 = ws.Cell("D4");

            ws.Row(3).Delete();

            Assert.That(ws.Cell("D3"), Is.SameAs(d4));
        }

        [Test]
        public void CellReferenceRemainAfterColumnInserted()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var d4 = ws.Cell("D4");

            ws.Column("C").InsertColumnsBefore(1);

            Assert.That(ws.Cell("E4"), Is.SameAs(d4));
        }

        [Test]
        public void CellReferenceRemainAfterRowInserted()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var d4 = ws.Cell("D4");

            ws.Row(3).InsertRowsAbove(1);

            Assert.That(ws.Cell("D5"), Is.SameAs(d4));
        }

        [Test]
        public void CellReferenceRemainAfterRangeDeleted()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var d4 = ws.Cell("D4");
            var f8 = ws.Cell("F8");

            ws.Range("B2:C5").Delete(XLShiftDeletedCells.ShiftCellsLeft);
            ws.Range("E5:F7").Delete(XLShiftDeletedCells.ShiftCellsUp);

            Assert.That(ws.Cell("B4"), Is.SameAs(d4));
            Assert.That(ws.Cell("F5"), Is.SameAs(f8));
        }
    }
}
