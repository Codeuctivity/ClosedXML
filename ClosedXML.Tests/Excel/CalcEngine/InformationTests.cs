using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Globalization;
using System.Threading;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class InformationTests
    {
        [SetUp]
        public void SetCultureInfo()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
        }

        #region IsBlank Tests

        [Test]
        public void IsBlank_MultipleAllEmpty_true()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            var actual = ws.Evaluate("=IsBlank(A1:A3)");
            Assert.That(actual, Is.EqualTo(true));
        }

        [Test]
        public void IsBlank_MultipleAllFill_false()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = "1";
            ws.Cell("A2").Value = "1";
            ws.Cell("A3").Value = "1";
            var actual = ws.Evaluate("=IsBlank(A1:A3)");
            Assert.That(actual, Is.EqualTo(false));
        }

        [Test]
        public void IsBlank_MultipleMixedFill_false()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = "1";
            ws.Cell("A3").Value = "1";
            var actual = ws.Evaluate("=IsBlank(A1:A3)");
            Assert.That(actual, Is.EqualTo(false));
        }

        [Test]
        public void IsBlank_Single_false()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = " ";
            var actual = ws.Evaluate("=IsBlank(A1)");
            Assert.That(actual, Is.EqualTo(false));
        }

        [Test]
        public void IsBlank_Single_true()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            var actual = ws.Evaluate("=IsBlank(A1)");
            Assert.That(actual, Is.EqualTo(true));
        }

        #endregion IsBlank Tests

        #region IsEven Tests

        [Test]
        public void IsEven_Single_False()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");

            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 1.2;
            ws.Cell("A3").Value = 3;

            var actual = ws.Evaluate("=IsEven(A1)");
            Assert.That(actual, Is.EqualTo(false));

            actual = ws.Evaluate("=IsEven(A2)");
            Assert.That(actual, Is.EqualTo(false));

            actual = ws.Evaluate("=IsEven(A3)");
            Assert.That(actual, Is.EqualTo(false));
        }

        [Test]
        public void IsEven_Single_True()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");

            ws.Cell("A1").Value = 4;
            ws.Cell("A2").Value = 0.2;
            ws.Cell("A3").Value = 12.2;

            var actual = ws.Evaluate("=IsEven(A1)");
            Assert.That(actual, Is.EqualTo(true));

            actual = ws.Evaluate("=IsEven(A2)");
            Assert.That(actual, Is.EqualTo(true));

            actual = ws.Evaluate("=IsEven(A3)");
            Assert.That(actual, Is.EqualTo(true));
        }

        #endregion IsEven Tests

        #region IsLogical Tests

        [Test]
        public void IsLogical_Simpe_False()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");

            ws.Cell("A1").Value = 123;

            var actual = ws.Evaluate("=IsLogical(A1)");
            Assert.That(actual, Is.EqualTo(false));
        }

        [Test]
        public void IsLogical_Simple_True()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");

            ws.Cell("A1").Value = true;

            var actual = ws.Evaluate("=IsLogical(A1)");
            Assert.That(actual, Is.EqualTo(true));
        }

        #endregion IsLogical Tests

        [Test]
        public void IsNA()
        {
            object actual;
            actual = XLWorkbook.EvaluateExpr("ISNA(#N/A)");
            Assert.That(actual, Is.EqualTo(true));

            actual = XLWorkbook.EvaluateExpr("ISNA(#REF!)");
            Assert.That(actual, Is.EqualTo(false));
        }

        #region IsNotText Tests

        [Test]
        public void IsNotText_Simple_false()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = "asd";
            var actual = ws.Evaluate("=IsNonText(A1)");
            Assert.That(actual, Is.EqualTo(false));
        }

        [Test]
        public void IsNotText_Simple_true()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = "123"; //Double Value
            ws.Cell("A2").Value = DateTime.Now; //Date Value
            ws.Cell("A3").Value = "12,235.5"; //Comma Formatting
            ws.Cell("A4").Value = "$12,235.5"; //Currency Value
            ws.Cell("A5").Value = true; //Bool Value
            ws.Cell("A6").Value = "12%"; //Percentage Value

            var actual = ws.Evaluate("=IsNonText(A1)");
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("=IsNonText(A2)");
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("=IsNonText(A3)");
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("=IsNonText(A4)");
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("=IsNonText(A5)");
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("=IsNonText(A6)");
            Assert.That(actual, Is.EqualTo(true));
        }

        #endregion IsNotText Tests

        #region IsNumber Tests

        [Test]
        public void IsNumber_Simple_false()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = "asd"; //String Value
            ws.Cell("A2").Value = true; //Bool Value

            var actual = ws.Evaluate("=IsNumber(A1)");
            Assert.That(actual, Is.EqualTo(false));
            actual = ws.Evaluate("=IsNumber(A2)");
            Assert.That(actual, Is.EqualTo(false));
        }

        [Test]
        public void IsNumber_Simple_true()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = "123"; //Double Value
            ws.Cell("A2").Value = DateTime.Now; //Date Value
            ws.Cell("A3").Value = "12,235.5"; //Coma Formatting
            ws.Cell("A4").Value = "$12,235.5"; //Currency Value
            ws.Cell("A5").Value = "12%"; //Percentage Value

            var actual = ws.Evaluate("=IsNumber(A1)");
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("=IsNumber(A2)");
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("=IsNumber(A3)");
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("=IsNumber(A4)");
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("=IsNumber(A5)");
            Assert.That(actual, Is.EqualTo(true));
        }

        #endregion IsNumber Tests

        #region IsOdd Test

        [Test]
        public void IsOdd_Simple_false()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");

            ws.Cell("A1").Value = 4;
            ws.Cell("A2").Value = 0.2;
            ws.Cell("A3").Value = 12.2;

            var actual = ws.Evaluate("=IsOdd(A1)");
            Assert.That(actual, Is.EqualTo(false));
            actual = ws.Evaluate("=IsOdd(A2)");
            Assert.That(actual, Is.EqualTo(false));
            actual = ws.Evaluate("=IsOdd(A3)");
            Assert.That(actual, Is.EqualTo(false));
        }

        [Test]
        public void IsOdd_Simple_true()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");

            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 1.2;
            ws.Cell("A3").Value = 3;

            var actual = ws.Evaluate("=IsOdd(A1)");
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("=IsOdd(A2)");
            Assert.That(actual, Is.EqualTo(true));
            actual = ws.Evaluate("=IsOdd(A3)");
            Assert.That(actual, Is.EqualTo(true));
        }

        #endregion IsOdd Test

        [Test]
        public void IsRef()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = "123";

            ws.Cell("B1").FormulaA1 = "ISREF(A1)";
            ws.Cell("B2").FormulaA1 = "ISREF(5)";
            ws.Cell("B3").FormulaA1 = "ISREF(YEAR(TODAY()))";

            bool actual;
            actual = ws.Cell("B1").GetValue<bool>();
            Assert.That(actual, Is.EqualTo(true));

            actual = ws.Cell("B2").GetValue<bool>();
            Assert.That(actual, Is.EqualTo(false));

            actual = ws.Cell("B3").GetValue<bool>();
            Assert.That(actual, Is.EqualTo(false));
        }

        #region IsText Tests

        [Test]
        public void IsText_Simple_false()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = "123"; //Double Value
            ws.Cell("A2").Value = DateTime.Now; //Date Value
            ws.Cell("A3").Value = "12,235.5"; //Comma Formatting
            ws.Cell("A4").Value = "$12,235.5"; //Currency Value
            ws.Cell("A5").Value = true; //Bool Value
            ws.Cell("A6").Value = "12%"; //Percentage Value

            var actual = ws.Evaluate("=IsText(A1)");
            Assert.That(actual, Is.EqualTo(false));
            actual = ws.Evaluate("=IsText(A2)");
            Assert.That(actual, Is.EqualTo(false));
            actual = ws.Evaluate("=IsText(A3)");
            Assert.That(actual, Is.EqualTo(false));
            actual = ws.Evaluate("=IsText(A4)");
            Assert.That(actual, Is.EqualTo(false));
            actual = ws.Evaluate("=IsText(A5)");
            Assert.That(actual, Is.EqualTo(false));
            actual = ws.Evaluate("=IsText(A6)");
            Assert.That(actual, Is.EqualTo(false));
        }

        [Test]
        public void IsText_Simple_true()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");

            ws.Cell("A1").Value = "asd";

            var actual = ws.Evaluate("=IsText(A1)");
            Assert.That(actual, Is.EqualTo(true));
        }

        #endregion IsText Tests

        #region N Tests

        [Test]
        public void N_Date_SerialNumber()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            var testedDate = DateTime.Now;
            ws.Cell("A1").Value = testedDate;
            var actual = ws.Evaluate("=N(A1)");
            Assert.That(actual, Is.EqualTo(testedDate.ToOADate()));
        }

        [Test]
        public void N_False_Zero()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = false;
            var actual = ws.Evaluate("=N(A1)");
            Assert.That(actual, Is.EqualTo(0));
        }

        [Test]
        public void N_Number_Number()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            var testedValue = 123;
            ws.Cell("A1").Value = testedValue;
            var actual = ws.Evaluate("=N(A1)");
            Assert.That(actual, Is.EqualTo(testedValue));
        }

        [Test]
        public void N_String_Zero()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = "asd";
            var actual = ws.Evaluate("=N(A1)");
            Assert.That(actual, Is.EqualTo(0));
        }

        [Test]
        public void N_True_One()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = true;
            var actual = ws.Evaluate("=N(A1)");
            Assert.That(actual, Is.EqualTo(1));
        }

        #endregion N Tests
    }
}
