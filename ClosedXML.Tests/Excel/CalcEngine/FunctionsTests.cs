using ClosedXML.Excel;
using NUnit.Framework;
using System;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class FunctionsTests
    {
        [SetUp]
        public void Init()
        {
            // Make sure tests run on a deterministic culture
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        }

        [Test]
        public void Asc()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr(@"Asc(""Text"")");
            Assert.That(actual, Is.EqualTo("Text"));
        }

        [Test]
        public void Clean()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr(string.Format(@"Clean(""A{0}B"")", XLConstants.NewLine));
            Assert.That(actual, Is.EqualTo("AB"));
        }

        [Test]
        public void Combin()
        {
            var actual1 = XLWorkbook.EvaluateExpr("Combin(200, 2)");
            Assert.That(actual1, Is.EqualTo(19900.0));

            var actual2 = XLWorkbook.EvaluateExpr("Combin(20.1, 2.9)");
            Assert.That(actual2, Is.EqualTo(190.0));
        }

        [Test]
        public void Degrees()
        {
            var actual1 = XLWorkbook.EvaluateExpr("Degrees(180)");
            Assert.That(Math.PI - (double)actual1 < XLHelper.Epsilon, Is.True);
        }

        [Test]
        public void Dollar()
        {
            var actual = XLWorkbook.EvaluateExpr("Dollar(12345.123)");
            Assert.That(actual, Is.EqualTo(TestHelper.CurrencySymbol + "12,345.12"));

            actual = XLWorkbook.EvaluateExpr("Dollar(12345.123, 1)");
            Assert.That(actual, Is.EqualTo(TestHelper.CurrencySymbol + "12,345.1"));
        }

        [Test]
        public void Even()
        {
            var actual = XLWorkbook.EvaluateExpr("Even(3)");
            Assert.That(actual, Is.EqualTo(4));

            actual = XLWorkbook.EvaluateExpr("Even(2)");
            Assert.That(actual, Is.EqualTo(2));

            actual = XLWorkbook.EvaluateExpr("Even(-1)");
            Assert.That(actual, Is.EqualTo(-2));

            actual = XLWorkbook.EvaluateExpr("Even(-2)");
            Assert.That(actual, Is.EqualTo(-2));

            actual = XLWorkbook.EvaluateExpr("Even(0)");
            Assert.That(actual, Is.EqualTo(0));

            actual = XLWorkbook.EvaluateExpr("Even(1.5)");
            Assert.That(actual, Is.EqualTo(2));

            actual = XLWorkbook.EvaluateExpr("Even(2.01)");
            Assert.That(actual, Is.EqualTo(4));
        }

        [Test]
        public void Exact()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr("Exact(\"A\", \"A\")");
            Assert.That(actual, Is.EqualTo(true));

            actual = XLWorkbook.EvaluateExpr("Exact(\"A\", \"a\")");
            Assert.That(actual, Is.EqualTo(false));
        }

        [Test]
        public void Fact()
        {
            var actual = XLWorkbook.EvaluateExpr("Fact(5.9)");
            Assert.That(actual, Is.EqualTo(120.0));
        }

        [Test]
        public void FactDouble()
        {
            var actual1 = XLWorkbook.EvaluateExpr("FactDouble(6)");
            Assert.That(actual1, Is.EqualTo(48.0));
            var actual2 = XLWorkbook.EvaluateExpr("FactDouble(7)");
            Assert.That(actual2, Is.EqualTo(105.0));
        }

        [Test]
        public void Fixed()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr("Fixed(12345.123)");
            Assert.That(actual, Is.EqualTo("12,345.12"));

            actual = XLWorkbook.EvaluateExpr("Fixed(12345.123, 1)");
            Assert.That(actual, Is.EqualTo("12,345.1"));

            actual = XLWorkbook.EvaluateExpr("Fixed(12345.123, 1, TRUE)");
            Assert.That(actual, Is.EqualTo("12345.1"));
        }

        [Test]
        public void Formula_from_another_sheet()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("ws1");
            ws1.FirstCell().SetValue(1).CellRight().SetFormulaA1("A1 + 1");
            var ws2 = wb.AddWorksheet("ws2");
            ws2.FirstCell().SetFormulaA1("ws1!B1 + 1");
            var v = ws2.FirstCell().Value;
            Assert.That(v, Is.EqualTo(3.0));
        }

        [Test]
        public void Gcd()
        {
            var actual = XLWorkbook.EvaluateExpr("Gcd(24, 36)");
            Assert.That(actual, Is.EqualTo(12));

            var actual1 = XLWorkbook.EvaluateExpr("Gcd(5, 0)");
            Assert.That(actual1, Is.EqualTo(5));

            var actual2 = XLWorkbook.EvaluateExpr("Gcd(0, 5)");
            Assert.That(actual2, Is.EqualTo(5));

            var actual3 = XLWorkbook.EvaluateExpr("Gcd(240, 360, 30)");
            Assert.That(actual3, Is.EqualTo(30));
        }

        [Test]
        public void Lcm()
        {
            var actual = XLWorkbook.EvaluateExpr("Lcm(24, 36)");
            Assert.That(actual, Is.EqualTo(72));

            var actual1 = XLWorkbook.EvaluateExpr("Lcm(5, 0)");
            Assert.That(actual1, Is.EqualTo(0));

            var actual2 = XLWorkbook.EvaluateExpr("Lcm(0, 5)");
            Assert.That(actual2, Is.EqualTo(0));

            var actual3 = XLWorkbook.EvaluateExpr("Lcm(240, 360, 30)");
            Assert.That(actual3, Is.EqualTo(720));
        }

        [Test]
        public void MDetem()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.AddWorksheet("Sheet1");
            ws.Cell("A1").SetValue(2).CellRight().SetValue(4);
            ws.Cell("A2").SetValue(3).CellRight().SetValue(5);

            object actual;

            ws.Cell("A5").FormulaA1 = "MDeterm(A1:B2)";
            actual = ws.Cell("A5").Value;

            Assert.That(XLHelper.AreEqual(-2.0, (double)actual), Is.True);

            ws.Cell("A6").FormulaA1 = "Sum(A5)";
            actual = ws.Cell("A6").Value;

            Assert.That(XLHelper.AreEqual(-2.0, (double)actual), Is.True);

            ws.Cell("A7").FormulaA1 = "Sum(MDeterm(A1:B2))";
            actual = ws.Cell("A7").Value;

            Assert.That(XLHelper.AreEqual(-2.0, (double)actual), Is.True);
        }

        [Test]
        public void MInverse()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.AddWorksheet("Sheet1");
            ws.Cell("A1").SetValue(1).CellRight().SetValue(2).CellRight().SetValue(1);
            ws.Cell("A2").SetValue(3).CellRight().SetValue(4).CellRight().SetValue(-1);
            ws.Cell("A3").SetValue(0).CellRight().SetValue(2).CellRight().SetValue(0);

            object actual;

            ws.Cell("A5").FormulaA1 = "MInverse(A1:C3)";
            actual = ws.Cell("A5").Value;

            Assert.That(XLHelper.AreEqual(0.25, (double)actual), Is.True);

            ws.Cell("A6").FormulaA1 = "Sum(A5)";
            actual = ws.Cell("A6").Value;

            Assert.That(XLHelper.AreEqual(0.25, (double)actual), Is.True);

            ws.Cell("A7").FormulaA1 = "Sum(MInverse(A1:C3))";
            actual = ws.Cell("A7").Value;

            Assert.That(XLHelper.AreEqual(0.5, (double)actual), Is.True);
        }

        [Test]
        public void MMult()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.AddWorksheet("Sheet1");
            ws.Cell("A1").SetValue(2).CellRight().SetValue(4);
            ws.Cell("A2").SetValue(3).CellRight().SetValue(5);
            ws.Cell("A3").SetValue(2).CellRight().SetValue(4);
            ws.Cell("A4").SetValue(3).CellRight().SetValue(5);

            object actual;

            ws.Cell("A5").FormulaA1 = "MMult(A1:B2, A3:B4)";
            actual = ws.Cell("A5").Value;

            Assert.That(actual, Is.EqualTo(16.0));

            ws.Cell("A6").FormulaA1 = "Sum(A5)";
            actual = ws.Cell("A6").Value;

            Assert.That(actual, Is.EqualTo(16.0));

            ws.Cell("A7").FormulaA1 = "Sum(MMult(A1:B2, A3:B4))";
            actual = ws.Cell("A7").Value;

            Assert.That(actual, Is.EqualTo(102.0));
        }

        [Test]
        public void Mod()
        {
            var actual = XLWorkbook.EvaluateExpr("Mod(3, 2)");
            Assert.That(actual, Is.EqualTo(1));

            var actual1 = XLWorkbook.EvaluateExpr("Mod(-3, 2)");
            Assert.That(actual1, Is.EqualTo(1));

            var actual2 = XLWorkbook.EvaluateExpr("Mod(3, -2)");
            Assert.That(actual2, Is.EqualTo(-1));

            var actual3 = XLWorkbook.EvaluateExpr("Mod(-3, -2)");
            Assert.That(actual3, Is.EqualTo(-1));
        }

        [Test]
        public void Multinomial()
        {
            var actual = XLWorkbook.EvaluateExpr("Multinomial(2,3,4)");
            Assert.That(actual, Is.EqualTo(1260.0));
        }

        [Test]
        public void Odd()
        {
            var actual = XLWorkbook.EvaluateExpr("Odd(1.5)");
            Assert.That(actual, Is.EqualTo(3));

            var actual1 = XLWorkbook.EvaluateExpr("Odd(3)");
            Assert.That(actual1, Is.EqualTo(3));

            var actual2 = XLWorkbook.EvaluateExpr("Odd(2)");
            Assert.That(actual2, Is.EqualTo(3));

            var actual3 = XLWorkbook.EvaluateExpr("Odd(-1)");
            Assert.That(actual3, Is.EqualTo(-1));

            var actual4 = XLWorkbook.EvaluateExpr("Odd(-2)");
            Assert.That(actual4, Is.EqualTo(-3));

            actual = XLWorkbook.EvaluateExpr("Odd(0)");
            Assert.That(actual, Is.EqualTo(1));
        }

        [Test]
        public void Product()
        {
            var actual = XLWorkbook.EvaluateExpr("Product(2,3,4)");
            Assert.That(actual, Is.EqualTo(24.0));
        }

        [Test]
        public void Quotient()
        {
            var actual = XLWorkbook.EvaluateExpr("Quotient(5,2)");
            Assert.That(actual, Is.EqualTo(2));

            actual = XLWorkbook.EvaluateExpr("Quotient(4.5,3.1)");
            Assert.That(actual, Is.EqualTo(1));

            actual = XLWorkbook.EvaluateExpr("Quotient(-10,3)");
            Assert.That(actual, Is.EqualTo(-3));
        }

        [Test]
        public void Radians()
        {
            var actual = XLWorkbook.EvaluateExpr("Radians(270)");
            Assert.That(Math.Abs(4.71238898038469 - (double)actual) < XLHelper.Epsilon, Is.True);
        }

        [Test]
        public void Roman()
        {
            var actual = XLWorkbook.EvaluateExpr("Roman(3046, 1)");
            Assert.That(actual, Is.EqualTo("MMMXLVI"));

            actual = XLWorkbook.EvaluateExpr("Roman(270)");
            Assert.That(actual, Is.EqualTo("CCLXX"));

            actual = XLWorkbook.EvaluateExpr("Roman(3999, true)");
            Assert.That(actual, Is.EqualTo("MMMCMXCIX"));
        }

        [Test]
        public void Round()
        {
            var actual = XLWorkbook.EvaluateExpr("Round(2.15, 1)");
            Assert.That(actual, Is.EqualTo(2.2));

            actual = XLWorkbook.EvaluateExpr("Round(2.149, 1)");
            Assert.That(actual, Is.EqualTo(2.1));

            actual = XLWorkbook.EvaluateExpr("Round(-1.475, 2)");
            Assert.That(actual, Is.EqualTo(-1.48));

            actual = XLWorkbook.EvaluateExpr("Round(21.5, -1)");
            Assert.That(actual, Is.EqualTo(20.0));

            actual = XLWorkbook.EvaluateExpr("Round(626.3, -3)");
            Assert.That(actual, Is.EqualTo(1000.0));

            actual = XLWorkbook.EvaluateExpr("Round(1.98, -1)");
            Assert.That(actual, Is.EqualTo(0.0));

            actual = XLWorkbook.EvaluateExpr("Round(-50.55, -2)");
            Assert.That(actual, Is.EqualTo(-100.0));

            actual = XLWorkbook.EvaluateExpr("ROUND(59 * 0.535, 2)"); // (59 * 0.535) = 31.565
            Assert.That(actual, Is.EqualTo(31.57));

            actual = XLWorkbook.EvaluateExpr("ROUND(59 * -0.535, 2)"); // (59 * -0.535) = -31.565
            Assert.That(actual, Is.EqualTo(-31.57));
        }

        [Test]
        public void RoundDown()
        {
            var actual = XLWorkbook.EvaluateExpr("RoundDown(3.2, 0)");
            Assert.That(actual, Is.EqualTo(3.0));

            actual = XLWorkbook.EvaluateExpr("RoundDown(76.9, 0)");
            Assert.That(actual, Is.EqualTo(76.0));

            actual = XLWorkbook.EvaluateExpr("RoundDown(3.14159, 3)");
            Assert.That(actual, Is.EqualTo(3.141));

            actual = XLWorkbook.EvaluateExpr("RoundDown(-3.14159, 1)");
            Assert.That(actual, Is.EqualTo(-3.1));

            actual = XLWorkbook.EvaluateExpr("RoundDown(31415.92654, -2)");
            Assert.That(actual, Is.EqualTo(31400.0));

            actual = XLWorkbook.EvaluateExpr("RoundDown(0, 3)");
            Assert.That(actual, Is.EqualTo(0.0));
        }

        [Test]
        public void RoundUp()
        {
            var actual = XLWorkbook.EvaluateExpr("RoundUp(3.2, 0)");
            Assert.That(actual, Is.EqualTo(4.0));

            actual = XLWorkbook.EvaluateExpr("RoundUp(76.9, 0)");
            Assert.That(actual, Is.EqualTo(77.0));

            actual = XLWorkbook.EvaluateExpr("RoundUp(3.14159, 3)");
            Assert.That(actual, Is.EqualTo(3.142));

            actual = XLWorkbook.EvaluateExpr("RoundUp(-3.14159, 1)");
            Assert.That(actual, Is.EqualTo(-3.2));

            actual = XLWorkbook.EvaluateExpr("RoundUp(31415.92654, -2)");
            Assert.That(actual, Is.EqualTo(31500.0));

            actual = XLWorkbook.EvaluateExpr("RoundUp(0, 3)");
            Assert.That(actual, Is.EqualTo(0.0));
        }

        [Test]
        public void SeriesSum()
        {
            var actual = XLWorkbook.EvaluateExpr("SERIESSUM(2,3,4,5)");
            Assert.That(actual, Is.EqualTo(40.0));

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A2").FormulaA1 = "PI()/4";
            ws.Cell("A3").Value = 1;
            ws.Cell("A4").FormulaA1 = "-1/FACT(2)";
            ws.Cell("A5").FormulaA1 = "1/FACT(4)";
            ws.Cell("A6").FormulaA1 = "-1/FACT(6)";

            actual = ws.Evaluate("SERIESSUM(A2,0,2,A3:A6)");
            Assert.That(Math.Abs(0.70710321482284566 - (double)actual) < XLHelper.Epsilon, Is.True);
        }

        [Test]
        public void SqrtPi()
        {
            var actual = XLWorkbook.EvaluateExpr("SqrtPi(1)");
            Assert.That(Math.Abs(1.7724538509055159 - (double)actual) < XLHelper.Epsilon, Is.True);

            actual = XLWorkbook.EvaluateExpr("SqrtPi(2)");
            Assert.That(Math.Abs(2.5066282746310002 - (double)actual) < XLHelper.Epsilon, Is.True);
        }

        [Test]
        public void SubtotalAverage()
        {
            var actual = XLWorkbook.EvaluateExpr("Subtotal(1,2,3)");
            Assert.That(actual, Is.EqualTo(2.5));

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(1,""A"",3, 2)");
            Assert.That(actual, Is.EqualTo(2.5));
        }

        [Test]
        public void SubtotalCount()
        {
            var actual = XLWorkbook.EvaluateExpr("Subtotal(2,2,3)");
            Assert.That(actual, Is.EqualTo(2));

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(2,""A"",3)");
            Assert.That(actual, Is.EqualTo(1));
        }

        [Test]
        public void SubtotalCountA()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr("Subtotal(3,2,3)");
            Assert.That(actual, Is.EqualTo(2.0));

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(3,"""",3)");
            Assert.That(actual, Is.EqualTo(1.0));
        }

        [Test]
        public void SubtotalMax()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(4,2,3,""A"")");
            Assert.That(actual, Is.EqualTo(3.0));
        }

        [Test]
        public void SubtotalMin()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(5,2,3,""A"")");
            Assert.That(actual, Is.EqualTo(2.0));
        }

        [Test]
        public void SubtotalProduct()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(6,2,3,""A"")");
            Assert.That(actual, Is.EqualTo(6.0));
        }

        [Test]
        public void SubtotalStDev()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(7,2,3,""A"")");
            Assert.That(Math.Abs(0.70710678118654757 - (double)actual) < XLHelper.Epsilon, Is.True);
        }

        [Test]
        public void SubtotalStDevP()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(8,2,3,""A"")");
            Assert.That(actual, Is.EqualTo(0.5));
        }

        [Test]
        public void SubtotalSum()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(9,2,3,""A"")");
            Assert.That(actual, Is.EqualTo(5.0));
        }

        [Test]
        public void SubtotalVar()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(10,2,3,""A"")");
            Assert.That(Math.Abs(0.5 - (double)actual) < XLHelper.Epsilon, Is.True);
        }

        [Test]
        public void SubtotalVarP()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr(@"Subtotal(11,2,3,""A"")");
            Assert.That(actual, Is.EqualTo(0.25));
        }

        [Test]
        public void SubtotalCalc()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.NamedRanges.Add("subtotalrange", "A37:A38");

            ws.Cell("A1").Value = 2;
            ws.Cell("A2").Value = 4;
            ws.Cell("A3").FormulaA1 = "SUBTOTAL(9, A1:A2)"; // simple add subtotal
            ws.Cell("A4").Value = 8;
            ws.Cell("A5").Value = 16;
            ws.Cell("A6").FormulaA1 = "SUBTOTAL(9, A4:A5)"; // simple add subtotal
            ws.Cell("A7").Value = 32;
            ws.Cell("A8").Value = 64;
            ws.Cell("A9").FormulaA1 = "SUM(A7:A8)"; // func but not subtotal
            ws.Cell("A10").Value = 128;
            ws.Cell("A11").Value = 256;
            ws.Cell("A12").FormulaA1 = "SUBTOTAL(1, A10:A11)"; // simple avg subtotal
            ws.Cell("A13").Value = 512;
            ws.Cell("A14").FormulaA1 = "SUBTOTAL(9, A1:A13)"; // subtotals in range
            ws.Cell("A15").Value = 1024;
            ws.Cell("A16").Value = 2048;
            ws.Cell("A17").FormulaA1 = "42 + SUBTOTAL(9, A15:A16)"; // simple add subtotal in formula
            ws.Cell("A18").Value = 4096;
            ws.Cell("A19").FormulaA1 = "SUBTOTAL(9, A15:A18)"; // subtotals in range
            ws.Cell("A20").Value = 8192;
            ws.Cell("A21").Value = 16384;
            ws.Cell("A22").FormulaA1 = @"32768 * SEARCH(""SUBTOTAL(9, A1:A2)"", A28)"; // subtotal literal in formula
            ws.Cell("A23").FormulaA1 = "SUBTOTAL(9, A20:A22)"; // subtotal literal in formula in range
            ws.Cell("A24").Value = 65536;
            ws.Cell("A25").FormulaA1 = "A23"; // link to subtotal
            ws.Cell("A26").FormulaA1 = "PRODUCT(SUBTOTAL(9, A24:A25), 2)"; // subtotal as parameter in func
            ws.Cell("A27").Value = 131072;
            ws.Cell("A28").Value = "SUBTOTAL(9, A1:A2)"; // subtotal literal
            ws.Cell("A29").FormulaA1 = "SUBTOTAL(9, A27:A28)"; // subtotal literal in range
            ws.Cell("A30").FormulaA1 = "SUBTOTAL(9, A31:A32)"; // simple add subtotal backward
            ws.Cell("A31").Value = 262144;
            ws.Cell("A32").Value = 524288;
            ws.Cell("A33").FormulaA1 = "SUBTOTAL(9, A20:A32)"; // subtotals in range
            ws.Cell("A34").FormulaA1 = @"SUBTOTAL(VALUE(""9""), A1:A33, A35:A41)"; // func as parameter in subtotal and many ranges
            ws.Cell("A35").Value = 1048576;
            ws.Cell("A36").FormulaA1 = "SUBTOTAL(9, A31:A32, A35)"; // many ranges
            ws.Cell("A37").Value = 2097152;
            ws.Cell("A38").Value = 4194304;
            ws.Cell("A39").FormulaA1 = "SUBTOTAL(3*3, subtotalrange)"; // formula as parameter in subtotal and named range
            ws.Cell("A40").Value = 8388608;
            ws.Cell("A41").FormulaA1 = "PRODUCT(SUBTOTAL(A4+1, A35:A40), 2)"; // formula with link as parameter in subtotal
            ws.Cell("A42").FormulaA1 = "PRODUCT(SUBTOTAL(A4+1, A35:A40), 2) + SUBTOTAL(A4+1, A35:A40)"; // two subtotals in one formula

            Assert.That(ws.Cell("A3").Value, Is.EqualTo(6));
            Assert.That(ws.Cell("A6").Value, Is.EqualTo(24));
            Assert.That(ws.Cell("A12").Value, Is.EqualTo(192));
            Assert.That(ws.Cell("A14").Value, Is.EqualTo(1118));
            Assert.That(ws.Cell("A17").Value, Is.EqualTo(3114));
            Assert.That(ws.Cell("A19").Value, Is.EqualTo(7168));
            Assert.That(ws.Cell("A23").Value, Is.EqualTo(57344));
            Assert.That(ws.Cell("A26").Value, Is.EqualTo(245760));
            Assert.That(ws.Cell("A29").Value, Is.EqualTo(131072));
            Assert.That(ws.Cell("A30").Value, Is.EqualTo(786432));
            Assert.That(ws.Cell("A33").Value, Is.EqualTo(1097728));
            Assert.That(ws.Cell("A34").Value, Is.EqualTo(16834654));
            Assert.That(ws.Cell("A36").Value, Is.EqualTo(1835008));
            Assert.That(ws.Cell("A39").Value, Is.EqualTo(6291456));
            Assert.That(ws.Cell("A41").Value, Is.EqualTo(31457280));
            Assert.That(ws.Cell("A42").Value, Is.EqualTo(47185920));
        }

        [Test]
        public void Sum()
        {
            using var xLWorkbook = new XLWorkbook();
            var cell = xLWorkbook.AddWorksheet("Sheet1").FirstCell();
            var fCell = cell.SetValue(1).CellBelow().SetValue(2).CellBelow();
            fCell.FormulaA1 = "sum(A1:A2)";

            Assert.That(fCell.Value, Is.EqualTo(3.0));
        }

        [Test]
        public void SumDateTimeAndNumber()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = new DateTime(2018, 1, 1);
            Assert.That(ws.Evaluate("SUM(A1:A2)"), Is.EqualTo(43102));

            ws.Cell("A1").Value = 2;
            ws.Cell("A2").FormulaA1 = "DATE(2018,1,1)";
            Assert.That(ws.Evaluate("SUM(A1:A2)"), Is.EqualTo(43103));
        }

        [Test]
        public void SumSq()
        {
            object actual;

            actual = XLWorkbook.EvaluateExpr(@"SumSq(3,4)");
            Assert.That(actual, Is.EqualTo(25.0));
        }

        [Test]
        public void TextConcat()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 1;
            ws.Cell("B1").Value = 1;
            ws.Cell("B2").Value = 1;

            ws.Cell("C1").FormulaA1 = "\"The total value is: \" & SUM(A1:B2)";

            var r = ws.Cell("C1").Value;
            Assert.That(r.ToString(), Is.EqualTo("The total value is: 4"));
        }

        [Test]
        public void Trim()
        {
            Assert.That(XLWorkbook.EvaluateExpr("Trim(\"Test    \")"), Is.EqualTo("Test"));

            //Should not trim non breaking space
            //See http://office.microsoft.com/en-us/excel-help/trim-function-HP010062581.aspx
            Assert.That(XLWorkbook.EvaluateExpr("Trim(\"Test\u00A0 \")"), Is.EqualTo("Test\u00A0"));
        }

        [Test]
        public void TestEmptyTallyOperations()
        {
            //In these test no values have been set
            using var wb = new XLWorkbook();
            wb.Worksheets.Add("TallyTests");
            var cell = wb.Worksheet(1).Cell(1, 1).SetFormulaA1("=MAX(D1,D2)");
            Assert.That(cell.Value, Is.EqualTo(0));
            cell = wb.Worksheet(1).Cell(2, 1).SetFormulaA1("=MIN(D1,D2)");
            Assert.That(cell.Value, Is.EqualTo(0));
            cell = wb.Worksheet(1).Cell(3, 1).SetFormulaA1("=SUM(D1,D2)");
            Assert.That(cell.Value, Is.EqualTo(0));
            Assert.That(() => wb.Worksheet(1).Cell(3, 1).SetFormulaA1("=AVERAGE(D1,D2)").Value, Throws.TypeOf<ApplicationException>());
        }

        [Test]
        public void TestOmittedParameters()
        {
            using var wb = new XLWorkbook();
            object value;
            value = wb.Evaluate("=IF(TRUE,1)");
            Assert.That(value, Is.EqualTo(1));

            value = wb.Evaluate("=IF(TRUE,1,)");
            Assert.That(value, Is.EqualTo(1));

            value = wb.Evaluate("=IF(FALSE,1,)");
            Assert.That(value, Is.EqualTo(false));

            value = wb.Evaluate("=IF(FALSE,,2)");
            Assert.That(value, Is.EqualTo(2));
        }

        [Test]
        public void TestDefaultExcelFunctionNamespace()
        {
            Assert.DoesNotThrow(() => XLWorkbook.EvaluateExpr("TODAY()"));
            Assert.DoesNotThrow(() => XLWorkbook.EvaluateExpr("_xlfn.TODAY()"));
            Assert.That((bool)XLWorkbook.EvaluateExpr("_xlfn.TODAY() = TODAY()"), Is.True);
        }

        [TestCase("=1234%", 12.34)]
        [TestCase("=1234%%", 0.1234)]
        [TestCase("=100+200%", 102.0)]
        [TestCase("=100%+200", 201.0)]
        [TestCase("=(100+200)%", 3.0)]
        [TestCase("=200%^5", 32.0)]
        [TestCase("=200%^400%", 16.0)]
        [TestCase("=SUM(100,200,300)%", 6.0)]
        public void PercentOperator(string formula, double expectedResult)
        {
            var res = (double)XLWorkbook.EvaluateExpr(formula);

            Assert.That(res, Is.EqualTo(expectedResult).Within(XLHelper.Epsilon));
        }

        [TestCase("=--1", 1)]
        [TestCase("=++1", 1)]
        [TestCase("=-+-+-1", -1)]
        [TestCase("=2^---2", 0.25)]
        public void MultipleUnaryOperators(string formula, double expectedResult)
        {
            var res = (double)XLWorkbook.EvaluateExpr(formula);

            Assert.That(res, Is.EqualTo(expectedResult).Within(XLHelper.Epsilon));
        }

        [TestCase("RIGHT(\"2020\", 2) + 1", 21)]
        [TestCase("LEFT(\"20.2020\", 6) + 1", 21.202)]
        [TestCase("2 + (\"3\" & \"4\")", 36)]
        [TestCase("2 + \"3\" & \"4\"", "54")]
        [TestCase("\"7\" & \"4\"", "74")]
        public void TestStringSubExpression(string formula, object expectedResult)
        {
            var actual = XLWorkbook.EvaluateExpr(formula);

            Assert.That(actual, Is.EqualTo(expectedResult));
        }
    }
}