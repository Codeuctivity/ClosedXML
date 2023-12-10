using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Linq;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class FormulaCachingTests
    {
        [Test]
        public void NewWorkbookDoesNotNeedRecalculation()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var cell = sheet.Cell(1, 1);

            Assert.That(wb.RecalculationCounter, Is.EqualTo(0));
            Assert.That(cell.NeedsRecalculation, Is.False);
        }

        [Test]
        public void EditCellCausesCounterIncreasing()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var cell = sheet.Cell(1, 1);
            cell.Value = "1234567";

            Assert.That(wb.RecalculationCounter, Is.GreaterThan(0));
        }

        [Test]
        public void StaticCellDoesNotNeedRecalculation()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var cell = sheet.Cell(1, 1);
            cell.Value = "1234567";

            Assert.That(cell.NeedsRecalculation, Is.False);
        }

        [Test]
        public void EditCellInvalidatesDependentCells()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var cell = sheet.Cell(1, 1);
            var dependentCell = sheet.Cell(2, 1);
            dependentCell.FormulaA1 = "=A1";
            _ = dependentCell.Value;

            cell.Value = "1234567";

            Assert.That(dependentCell.NeedsRecalculation, Is.True);
        }

        [Test]
        public void EditFormulaA1InvalidatesDependentCells()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var a1 = sheet.Cell("A1");
            var a2 = sheet.Cell("A2");
            var a3 = sheet.Cell("A3");
            var a4 = sheet.Cell("A4");
            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=SUM(A1:A3)";
            a1.Value = 15;

            var res1 = a4.Value;
            a2.FormulaA1 = "=A1*20";
            var res2 = a4.Value;

            Assert.That(res1, Is.EqualTo(15 + 150 + 1500));
            Assert.That(res2, Is.EqualTo(15 + 300 + 3000));
        }

        [Test]
        public void EditFormulaR1C1InvalidatesDependentCells()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var a1 = sheet.Cell("A1");
            var a2 = sheet.Cell("A2");
            var a3 = sheet.Cell("A3");
            var a4 = sheet.Cell("A4");
            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=SUM(A1:A3)";
            a1.Value = 15;

            var res1 = a4.Value;
            a2.FormulaR1C1 = "=R[-1]C*2";
            var res2 = a4.Value;

            Assert.That(res1, Is.EqualTo(15 + 150 + 1500));
            Assert.That(res2, Is.EqualTo(15 + 30 + 300));
        }

        [Test]
        public void InsertRowInvalidatesValues()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var a4 = sheet.Cell("A4");
            a4.FormulaA1 = "=COUNTBLANK(A1:A3)";

            var res1 = a4.Value;
            sheet.Row(2).InsertRowsAbove(2);
            var res2 = a4.Value;

            Assert.That(res1, Is.EqualTo(3));
            Assert.That(res2, Is.EqualTo(5));
        }

        [Test]
        public void DeleteRowInvalidatesValues()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var a4 = sheet.Cell("A4");
            a4.FormulaA1 = "=COUNTBLANK(A1:A3)";

            var res1 = a4.Value;
            sheet.Row(2).Delete();
            var res2 = a4.Value;

            Assert.That(res1, Is.EqualTo(3));
            Assert.That(res2, Is.EqualTo(2));
        }

        [Test]
        public void ChainedCalculationPreservesIntermediateValues()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var a1 = sheet.Cell("A1");
            var a2 = sheet.Cell("A2");
            var a3 = sheet.Cell("A3");
            var a4 = sheet.Cell("A4");
            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=SUM(A1:A3)";

            a1.Value = 15;
            var res = a4.Value;

            Assert.That(res, Is.EqualTo(15 + 150 + 1500));
            Assert.That(a4.NeedsRecalculation, Is.False);
            Assert.That(a3.NeedsRecalculation, Is.False);
            Assert.That(a2.NeedsRecalculation, Is.False);
            Assert.That(a2.CachedValue, Is.EqualTo(150));
            Assert.That(a3.CachedValue, Is.EqualTo(1500));
            Assert.That(a4.CachedValue, Is.EqualTo(15 + 150 + 1500));
        }

        [Test]
        public void EditingAffectsDependentCells()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var a1 = sheet.Cell("A1");
            var a2 = sheet.Cell("A2");
            var a3 = sheet.Cell("A3");
            var a4 = sheet.Cell("A4");
            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=SUM(A1:A3)";
            a1.Value = 15;

            var res1 = a4.Value;
            a1.Value = 20;
            var res2 = a4.Value;

            Assert.That(res1, Is.EqualTo(15 + 150 + 1500));
            Assert.That(res2, Is.EqualTo(20 + 200 + 2000));
        }

        [Test]
        [TestCase("C4", new string[] { "C5" })]
        [TestCase("D4", new string[] { })]
        [TestCase("A1", new string[] { "A2", "A3", "A4", "C1", "C2", "C3", "C5" })]
        [TestCase("B2", new string[] { "B3", "B4", "C2", "C3", "C5" })]
        [TestCase("C2", new string[] { "C5" })]
        public void EditingDoesNotAffectNonDependingCells(string changedCell, string[] affectedCells)
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            sheet.Cell("A2").FormulaA1 = "A1+1";
            sheet.Cell("A3").FormulaA1 = "SUM(A1:A2)";
            sheet.Cell("A4").FormulaA1 = "SUM(A1:A3)";
            sheet.Cell("B2").FormulaA1 = "B1+1";
            sheet.Cell("B3").FormulaA1 = "SUM(B1:B2)";
            sheet.Cell("B4").FormulaA1 = "SUM(B1:B3)";
            sheet.Cell("C1").FormulaA1 = "SUM(A1:B1)";
            sheet.Cell("C2").FormulaA1 = "SUM(A2:B2)";
            sheet.Cell("C3").FormulaA1 = "SUM(A3:B3)";
            sheet.Cell("C5").FormulaA1 = "SUM($A$1:$C$4)";
            sheet.RecalculateAllFormulas();
            var allCells = sheet.CellsUsed();

            sheet.Cell(changedCell).Value = 100;
            var modifiedCells = allCells.Where(cell => cell.NeedsRecalculation);

            Assert.That(modifiedCells.Count(), Is.EqualTo(affectedCells?.Length));
            foreach (var cellAddress in affectedCells)
            {
                Assert.That(modifiedCells.Any(cell => cell.Address.ToString() == cellAddress), Is.True,
                    string.Format("Cell {0} is expected to need recalculation, but it does not", cellAddress));
            }
        }

        [Test]
        public void CircularReferenceFailsCalculating()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var a1 = sheet.Cell("A1");
            var a2 = sheet.Cell("A2");
            var a3 = sheet.Cell("A3");
            var a4 = sheet.Cell("A4");

            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=A3*10";
            a1.FormulaA1 = "A2+A3+A4";

            var getValueA1 = new TestDelegate(() => { _ = a1.Value; });
            var getValueA2 = new TestDelegate(() => { _ = a2.Value; });
            var getValueA3 = new TestDelegate(() => { _ = a3.Value; });
            var getValueA4 = new TestDelegate(() => { _ = a4.Value; });

            Assert.Throws(typeof(InvalidOperationException), getValueA1);
            Assert.Throws(typeof(InvalidOperationException), getValueA2);
            Assert.Throws(typeof(InvalidOperationException), getValueA3);
            Assert.Throws(typeof(InvalidOperationException), getValueA4);
        }

        [Test]
        public void CircularReferenceRecalculationNeededDoesNotFail()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var a1 = sheet.Cell("A1");
            var a2 = sheet.Cell("A2");
            var a3 = sheet.Cell("A3");
            var a4 = sheet.Cell("A4");

            a2.FormulaA1 = "=A1*10";
            a3.FormulaA1 = "=A2*10";
            a4.FormulaA1 = "=A3*10";
            _ = a4.Value;
            a1.FormulaA1 = "=SUM(A2:A4)";

            var recalcNeededA1 = a1.NeedsRecalculation;
            var recalcNeededA2 = a2.NeedsRecalculation;
            var recalcNeededA3 = a3.NeedsRecalculation;
            var recalcNeededA4 = a4.NeedsRecalculation;

            Assert.That(recalcNeededA1, Is.True);
            Assert.That(recalcNeededA2, Is.True);
            Assert.That(recalcNeededA3, Is.True);
            Assert.That(recalcNeededA4, Is.True);
        }

        [Test]
        public void DeleteWorksheetInvalidatesValues()
        {
            using var wb = new XLWorkbook();
            var sheet1 = wb.Worksheets.Add("Sheet1");
            var sheet2 = wb.Worksheets.Add("Sheet2");
            var sheet1_a1 = sheet1.Cell("A1");
            var sheet2_a1 = sheet2.Cell("A1");
            sheet1_a1.FormulaA1 = "Sheet2!A1";
            sheet2_a1.Value = "TestValue";

            var val1 = sheet1_a1.Value;
            sheet2.Delete();
            var getValue = new TestDelegate(() => { _ = sheet1_a1.Value; });

            Assert.That(val1.ToString(), Is.EqualTo("TestValue"));
            Assert.Throws(typeof(ArgumentOutOfRangeException), getValue);
        }

        [Test]
        public void TestValueCellsCachedValue()
        {
            using var wb = new XLWorkbook();
            var sheet = wb.Worksheets.Add("TestSheet");
            var cell = sheet.Cell(1, 1);

            var date = new DateTime(2018, 4, 19);
            cell.Value = date;

            Assert.That(cell.DataType, Is.EqualTo(XLDataType.DateTime));
            Assert.That(cell.CachedValue, Is.EqualTo(date));

            cell.DataType = XLDataType.Number;

            Assert.That(cell.DataType, Is.EqualTo(XLDataType.Number));
            Assert.That(cell.CachedValue, Is.EqualTo(date.ToOADate()));
        }

        [Test]
        public void CachedValueToExternalWorkbook()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\ExternalLinks\WorkbookWithExternalLink.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheets.First();
            var cell = ws.Cell("B2");
            Assert.That(cell.NeedsRecalculation, Is.False);
            Assert.That(cell.HasFormula, Is.True);

            // This will fail when we start supporting external links
            Assert.That(cell.FormulaA1.StartsWith("[1]"), Is.True);

            Assert.That(cell.CachedValue, Is.EqualTo("hello world"));
            Assert.That(cell.Value, Is.EqualTo("hello world"));

            Assert.That(ws.Evaluate("LEN(B2)"), Is.EqualTo(11));

            Assert.Throws<ArgumentOutOfRangeException>(() => wb.RecalculateAllFormulas());
        }

        [Test]
        public void ChangingDataTypeChangesCachedValue()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Test");
            ws.Cell(1, 1).Value = new DateTime(2019, 1, 1, 14, 0, 0);
            ws.Cell(1, 2).Value = new DateTime(2019, 1, 1, 17, 45, 0);
            var cell = ws.Cell(1, 3);
            cell.FormulaA1 = "=B1-A1";
            cell.Style.DateFormat.Format = "hh:mm";

            Assert.That(cell.CachedValue, Is.Null);

            var value = (double)cell.Value;
            Assert.That(cell.CachedValue, Is.EqualTo(value));

            cell.DataType = XLDataType.DateTime;
            Assert.That(cell.CachedValue, Is.EqualTo(DateTime.FromOADate(value)));
            Assert.That(cell.GetFormattedString(), Is.EqualTo("03:45"));

            cell.DataType = XLDataType.Number;
            Assert.That((double)cell.CachedValue, Is.EqualTo(value).Within(1e-10));
            Assert.That(cell.GetFormattedString(), Is.EqualTo("03:45"));

            cell.DataType = XLDataType.TimeSpan;
            Assert.That((TimeSpan)cell.CachedValue, Is.EqualTo(TimeSpan.FromDays(value)));
            Assert.That(cell.GetFormattedString(), Is.EqualTo("03:45:00")); // I think the seconds in this string is due to a shortcoming in the ExcelNumberFormat library
        }
    }
}