using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;

namespace ClosedXML.Tests.Excel.Cells
{
    [TestFixture]
    public class XLCellTests
    {
        [Test]
        public void CellsUsed()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            ws.Cell(1, 1);
            ws.Cell(2, 2);
            var count = ws.Range("A1:B2").CellsUsed().Count();
            Assert.That(count, Is.EqualTo(0));
        }

        [Test]
        public void CellsUsedIncludeStyles1()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            ws.Row(3).Style.Fill.BackgroundColor = XLColor.Red;
            ws.Column(3).Style.Fill.BackgroundColor = XLColor.Red;
            ws.Cell(2, 2).Value = "ASDF";
            var range = ws.RangeUsed(XLCellsUsedOptions.All).RangeAddress.ToString();
            Assert.That(range, Is.EqualTo("B2:C3"));
        }

        [Test]
        public void CellsUsedIncludeStyles2()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            ws.Row(2).Style.Fill.BackgroundColor = XLColor.Red;
            ws.Column(2).Style.Fill.BackgroundColor = XLColor.Red;
            ws.Cell(3, 3).Value = "ASDF";
            var range = ws.RangeUsed(XLCellsUsedOptions.All).RangeAddress.ToString();
            Assert.That(range, Is.EqualTo("B2:C3"));
        }

        [Test]
        public void CellsUsedIncludeStyles3()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var range = ws.RangeUsed(XLCellsUsedOptions.All);
            Assert.That(range, Is.EqualTo(null));
        }

        [Test]
        public void CellUsedIncludesSparklines()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            ws.Range("C3:E4").Value = 1;
            ws.SparklineGroups.Add("B2", "C3:E3");
            ws.SparklineGroups.Add("F5", "C4:E4");

#pragma warning disable CS0618 // Type or member is obsolete, but still should be tested
            var range = ws.RangeUsed(true).RangeAddress.ToString();
#pragma warning restore CS0618 // Type or member is obsolete, but still should be tested
            Assert.That(range, Is.EqualTo("B2:F5"));
        }

        [Test]
        public void Double_Infinity_is_a_string()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1");
            var doubleList = new List<double> { 1.0 / 0.0 };

            cell.Value = 5;
            cell.Value = doubleList;
            Assert.That(cell.DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(cell.Value, Is.EqualTo(CultureInfo.CurrentCulture.NumberFormat.PositiveInfinitySymbol));

            cell.Value = 5;
            cell.SetValue(doubleList);
            Assert.That(cell.DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(cell.Value, Is.EqualTo(CultureInfo.CurrentCulture.NumberFormat.PositiveInfinitySymbol));
        }

        [Test]
        public void Double_NaN_is_a_string()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1");
            var doubleList = new List<double> { 0.0 / 0.0 };

            cell.Value = 5;
            cell.Value = doubleList;
            Assert.That(cell.DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(cell.Value, Is.EqualTo(CultureInfo.CurrentCulture.NumberFormat.NaNSymbol));

            cell.Value = 5;
            cell.SetValue(doubleList);
            Assert.That(cell.DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(cell.Value, Is.EqualTo(CultureInfo.CurrentCulture.NumberFormat.NaNSymbol));
        }

        [Test]
        public void GetValue_Nullable()
        {
            var backupCulture = Thread.CurrentThread.CurrentCulture;

            // Set thread culture to French, which should format numbers using a space as thousands separator
            try
            {
                var culture = CultureInfo.CreateSpecificCulture("fr-FR");
                // but use a period instead of a comma as for decimal separator
                culture.NumberFormat.CurrencyDecimalSeparator = ".";
                Thread.CurrentThread.CurrentCulture = culture;

                using var xLWorkbook = new XLWorkbook();
                var cell = xLWorkbook.AddWorksheet().FirstCell();

                Assert.That(cell.Clear().GetValue<double?>(), Is.Null);
                Assert.That(cell.SetValue(1.5).GetValue<double?>(), Is.EqualTo(1.5));
                Assert.That(cell.SetValue(1.5).GetValue<int?>(), Is.EqualTo(2));
                Assert.That(cell.SetValue("2.5").GetValue<double?>(), Is.EqualTo(2.5));
                Assert.Throws<FormatException>(() => cell.SetValue("text").GetValue<double?>());
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = backupCulture;
            }
        }

        [Test]
        public void InsertData1()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var range = ws.Cell(2, 2).InsertData(new[] { "a", "b", "c" });
            Assert.That(range.ToString(), Is.EqualTo("Sheet1!B2:B4"));
        }

        [Test]
        public void InsertData2()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var range = ws.Cell(2, 2).InsertData(new[] { "a", "b", "c" }, false);
            Assert.That(range.ToString(), Is.EqualTo("Sheet1!B2:B4"));
        }

        [Test]
        public void InsertData3()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var range = ws.Cell(2, 2).InsertData(new[] { "a", "b", "c" }, true);
            Assert.That(range.ToString(), Is.EqualTo("Sheet1!B2:D2"));
        }

        [Test]
        public void InsertData_with_Guids()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            ws.FirstCell().InsertData(Enumerable.Range(1, 20).Select(i => new { Guid = Guid.NewGuid() }));

            Assert.That(ws.FirstCell().DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.FirstCell().GetString().Length, Is.EqualTo(Guid.NewGuid().ToString().Length));
        }

        [Test]
        public void InsertData_with_Nulls()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");

            using var table = new DataTable
            {
                TableName = "Patients"
            };
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            table.Rows.Add(25, "Indocin", "David", new DateTime(2000, 1, 1));
            table.Rows.Add(50, "Enebrel", "Sam", new DateTime(2000, 1, 2));
            table.Rows.Add(10, "Hydralazine", "Christoff", new DateTime(2000, 1, 3));
            table.Rows.Add(21, "Combivent", DBNull.Value, new DateTime(2000, 1, 4));
            table.Rows.Add(100, "Dilantin", "Melanie", DBNull.Value);

            ws.FirstCell().InsertData(table);

            Assert.That(ws.Cell("A1").Value, Is.EqualTo(25));
            Assert.That(ws.Cell("C4").Value, Is.EqualTo(""));
            Assert.That(ws.Cell("D5").Value, Is.EqualTo(""));
        }

        [Test]
        public void InsertData_with_Nulls_IEnumerable()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");

            var dateTimeList = new List<DateTime?>()
            {
                new DateTime(2000, 1, 1),
                new DateTime(2000, 1, 2),
                new DateTime(2000, 1, 3),
                new DateTime(2000, 1, 4),
                null
            };

            ws.FirstCell().InsertData(dateTimeList);

            Assert.That(ws.Cell("A1").GetDateTime(), Is.EqualTo(new DateTime(2000, 1, 1)));
            Assert.That(ws.Cell("A5").Value, Is.EqualTo(string.Empty));
        }

        [Test]
        public void IsEmpty1()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            var actual = cell.IsEmpty();
            var expected = true;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty2()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            var actual = cell.IsEmpty(XLCellsUsedOptions.All);
            var expected = true;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty3()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            var actual = cell.IsEmpty();
            var expected = true;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty4()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            var actual = cell.IsEmpty(XLCellsUsedOptions.AllContents);
            var expected = true;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty5()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            var actual = cell.IsEmpty(XLCellsUsedOptions.All);
            var expected = false;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty6()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Value = "X";
            var actual = cell.IsEmpty();
            var expected = false;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty_Comment()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.GetComment().AddText("comment");
            var actual = cell.IsEmpty();
            var expected = false;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty_Comment_Value()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.GetComment().AddText("comment");
            cell.SetValue("value");

            var actual = cell.IsEmpty();
            var expected = false;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        [TestCase(XLCellsUsedOptions.Contents, true)]
        [TestCase(XLCellsUsedOptions.DataType, true)]
        [TestCase(XLCellsUsedOptions.NormalFormats, true)]
        [TestCase(XLCellsUsedOptions.ConditionalFormats, true)]
        [TestCase(XLCellsUsedOptions.Comments, false)]
        [TestCase(XLCellsUsedOptions.DataValidation, true)]
        [TestCase(XLCellsUsedOptions.MergedRanges, true)]
        [TestCase(XLCellsUsedOptions.Sparklines, true)]
        [TestCase(XLCellsUsedOptions.AllFormats, true)]
        [TestCase(XLCellsUsedOptions.AllContents, false)]
        [TestCase(XLCellsUsedOptions.All, false)]
        public void IsEmpty_Comment_Options(XLCellsUsedOptions options, bool expected)
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.GetComment().AddText("comment");

            var actual = cell.IsEmpty(options);

            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        [TestCase(XLCellsUsedOptions.Contents, false)]
        [TestCase(XLCellsUsedOptions.DataType, true)]
        [TestCase(XLCellsUsedOptions.NormalFormats, true)]
        [TestCase(XLCellsUsedOptions.ConditionalFormats, true)]
        [TestCase(XLCellsUsedOptions.Comments, false)]
        [TestCase(XLCellsUsedOptions.DataValidation, true)]
        [TestCase(XLCellsUsedOptions.MergedRanges, true)]
        [TestCase(XLCellsUsedOptions.Sparklines, true)]
        [TestCase(XLCellsUsedOptions.AllFormats, true)]
        [TestCase(XLCellsUsedOptions.AllContents, false)]
        [TestCase(XLCellsUsedOptions.All, false)]
        public void IsEmpty_Comment_Options_Value(XLCellsUsedOptions options, bool expected) // see #1575
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.GetComment().AddText("comment");
            cell.SetValue("value");

            var actual = cell.IsEmpty(options);

            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty_DataType()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.DataType = XLDataType.Boolean;
            var actual = cell.IsEmpty();
            var expected = false;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty_DataType_Text()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.DataType = XLDataType.Text;
            var actual = cell.IsEmpty();
            var expected = true;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty_DataType_Value()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.DataType = XLDataType.Number;
            cell.SetValue(42);

            var actual = cell.IsEmpty();
            var expected = false;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty_DataType_Text_Value()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.DataType = XLDataType.Text;
            cell.SetValue("value");

            var actual = cell.IsEmpty();
            var expected = false;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        [TestCase(XLCellsUsedOptions.Contents, true)]
        [TestCase(XLCellsUsedOptions.DataType, false)]
        [TestCase(XLCellsUsedOptions.NormalFormats, true)]
        [TestCase(XLCellsUsedOptions.ConditionalFormats, true)]
        [TestCase(XLCellsUsedOptions.Comments, true)]
        [TestCase(XLCellsUsedOptions.DataValidation, true)]
        [TestCase(XLCellsUsedOptions.MergedRanges, true)]
        [TestCase(XLCellsUsedOptions.Sparklines, true)]
        [TestCase(XLCellsUsedOptions.AllFormats, true)]
        [TestCase(XLCellsUsedOptions.AllContents, false)]
        [TestCase(XLCellsUsedOptions.All, false)]
        public void IsEmpty_DataType_Options(XLCellsUsedOptions options, bool expected)
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.DataType = XLDataType.Number;

            var actual = cell.IsEmpty(options);

            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        [TestCase(XLCellsUsedOptions.Contents, false)]
        [TestCase(XLCellsUsedOptions.DataType, false)]
        [TestCase(XLCellsUsedOptions.NormalFormats, true)]
        [TestCase(XLCellsUsedOptions.ConditionalFormats, true)]
        [TestCase(XLCellsUsedOptions.Comments, true)]
        [TestCase(XLCellsUsedOptions.DataValidation, true)]
        [TestCase(XLCellsUsedOptions.MergedRanges, true)]
        [TestCase(XLCellsUsedOptions.Sparklines, true)]
        [TestCase(XLCellsUsedOptions.AllFormats, true)]
        [TestCase(XLCellsUsedOptions.AllContents, false)]
        [TestCase(XLCellsUsedOptions.All, false)]
        public void IsEmpty_DataType_Options_Value(XLCellsUsedOptions options, bool expected) // see #1575
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.DataType = XLDataType.Number;
            cell.SetValue(42);

            var actual = cell.IsEmpty(options);

            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void NaN_is_not_a_number()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1");
            cell.Value = "NaN";

            Assert.That(cell.DataType, Is.Not.EqualTo(XLDataType.Number));
        }

        [Test]
        public void Nan_is_not_a_number()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1");
            cell.Value = "Nan";

            Assert.That(cell.DataType, Is.Not.EqualTo(XLDataType.Number));
        }

        [Test]
        public void TryGetValue_Boolean_Bad()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1").SetValue("ABC");
            var success = cell.TryGetValue(out bool outValue);
            Assert.That(success, Is.False);
        }

        [Test]
        public void TryGetValue_Boolean_False()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1").SetValue(false);
            var success = cell.TryGetValue(out bool outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.False);
        }

        [Test]
        public void TryGetValue_Boolean_Good()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1").SetValue("true");
            var success = cell.TryGetValue(out bool outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.True);
        }

        [Test]
        public void TryGetValue_Boolean_True()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1").SetValue(true);
            var success = cell.TryGetValue(out bool outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.True);
        }

        [Test]
        public void TryGetValue_DateTime_Good()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var date = "2018-01-01";
            var success = ws.Cell("A1").SetValue(date).TryGetValue(out DateTime outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.EqualTo(new DateTime(2018, 1, 1)));
        }

        [Test]
        public void TryGetValue_DateTime_Good2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var success = ws.Cell("A1").SetFormulaA1("=TODAY() + 10").TryGetValue(out DateTime outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.EqualTo(DateTime.Today.AddDays(10)));
        }

        [Test]
        public void TryGetValue_DateTime_BadButFormulaGood()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var success = ws.Cell("A1").SetFormulaA1("=\"44\"&\"020\"").TryGetValue(out DateTime outValue);
            Assert.That(success, Is.False);

            ws.Cell("B1").SetFormulaA1("=A1+1");

            success = ws.Cell("B1").TryGetValue(out outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.EqualTo(new DateTime(2020, 07, 09)));
        }

        [Test]
        public void TryGetValue_DateTime_BadString()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var date = "ABC";
            var success = ws.Cell("A1").SetValue(date).TryGetValue(out DateTime outValue);
            Assert.That(success, Is.False);
        }

        [Test]
        public void TryGetValue_DateTime_BadString2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var date = 5545454;
            ws.FirstCell().SetValue(date).DataType = XLDataType.DateTime;
            var success = ws.FirstCell().TryGetValue(out DateTime outValue);
            Assert.That(success, Is.False);
        }

        [Test]
        public void TryGetValue_Enum_Good()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.AddWorksheet();
            Assert.That(ws.FirstCell().SetValue(NumberStyles.AllowCurrencySymbol).TryGetValue(out NumberStyles value), Is.True);
            Assert.That(value, Is.EqualTo(NumberStyles.AllowCurrencySymbol));

            // Nullable alternative
            Assert.That(ws.FirstCell().SetValue(NumberStyles.AllowCurrencySymbol).TryGetValue(out NumberStyles? value2), Is.True);
            Assert.That(value2, Is.EqualTo(NumberStyles.AllowCurrencySymbol));
        }

        [Test]
        public void TryGetValue_Enum_BadString()
        {
            using var xLWorkbook = new XLWorkbook();
            var ws = xLWorkbook.AddWorksheet();
            Assert.That(ws.FirstCell().SetValue("ABC").TryGetValue(out NumberStyles value), Is.False);
            Assert.That(ws.FirstCell().SetValue("ABC").TryGetValue(out NumberStyles? value2), Is.False);
        }

        [Test]
        public void TryGetValue_RichText_Bad()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1").SetValue("Anything");
            var success = cell.TryGetValue(out IXLRichText outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.EqualTo(cell.GetRichText()));
            Assert.That(outValue.ToString(), Is.EqualTo("Anything"));
        }

        [Test]
        public void TryGetValue_RichText_Good()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1");
            cell.GetRichText().AddText("Anything");
            var success = cell.TryGetValue(out IXLRichText outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.EqualTo(cell.GetRichText()));
        }

        [Test]
        public void TryGetValue_TimeSpan_BadString()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var timeSpan = "ABC";
            var success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out TimeSpan outValue);
            Assert.That(success, Is.False);
        }

        [Test]
        public void TryGetValue_TimeSpan_Good()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var timeSpan = new TimeSpan(1, 1, 1);
            var success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out TimeSpan outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.EqualTo(timeSpan));
        }

        [Test]
        public void TryGetValue_TimeSpan_Good_Large()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var timeSpan = TimeSpan.FromMilliseconds((double)int.MaxValue + 1);
            var success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out TimeSpan outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.EqualTo(timeSpan));
        }

        [Test]
        public void TryGetValue_TimeSpan_GoodString()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var timeSpan = new TimeSpan(1, 1, 1);
            var success = ws.Cell("A1").SetValue(timeSpan.ToString()).TryGetValue(out TimeSpan outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.EqualTo(timeSpan));
        }

        [Test]
        public void TryGetValue_sbyte_Bad()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1").SetValue(255);
            var success = cell.TryGetValue(out sbyte outValue);
            Assert.That(success, Is.False);
        }

        [Test]
        public void TryGetValue_sbyte_Bad2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1").SetValue("255");
            var success = cell.TryGetValue(out sbyte outValue);
            Assert.That(success, Is.False);
        }

        [Test]
        public void TryGetValue_sbyte_Good()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1").SetValue(5);
            var success = cell.TryGetValue(out sbyte outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.EqualTo(5));
        }

        [Test]
        public void TryGetValue_sbyte_Good2()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1").SetValue("5");
            var success = cell.TryGetValue(out sbyte outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.EqualTo(5));
        }

        [Test]
        public void TryGetValue_decimal_Good()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1").SetValue("5");
            var success = cell.TryGetValue(out decimal outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.EqualTo(5));
        }

        [Test]
        public void TryGetValue_decimal_Good2()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");

            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell("A1").SetValue("1.60000001869776E-06");
            var success = cell.TryGetValue(out decimal outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.EqualTo(1.60000001869776E-06));
        }

        [Test]
        public void TryGetValue_Hyperlink()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            var ws2 = wb.Worksheets.Add("Sheet2");

            var targetCell = ws2.Cell("A1");

            var linkCell1 = ws1.Cell("A1");
            linkCell1.Value = "Link to IXLCell";
            linkCell1.SetHyperlink(new XLHyperlink(targetCell));

            var success = linkCell1.TryGetValue(out XLHyperlink hyperlink);
            Assert.That(success, Is.True);
            Assert.That(hyperlink.InternalAddress, Is.EqualTo("Sheet2!A1"));
        }

        [Test]
        public void TryGetValue_Unicode_String()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");

            var success = ws.Cell("A1")
                  .SetValue("Site_x0020_Column_x0020_Test")
                  .TryGetValue(out string outValue);
            Assert.That(success, Is.True);
            Assert.That(outValue, Is.EqualTo("Site Column Test"));

            success = ws.Cell("A1")
                .SetValue("Site_x005F_x0020_Column_x005F_x0020_Test")
                .TryGetValue(out outValue);

            Assert.That(success, Is.True);
            Assert.That(outValue, Is.EqualTo("Site_x005F_x0020_Column_x005F_x0020_Test"));
        }

        [Test]
        public void TryGetValue_Nullable()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cell("A1").Clear();
            ws.Cell("A2").SetValue(1.5);
            ws.Cell("A3").SetValue("2.5");
            ws.Cell("A4").SetValue("text");

            foreach (var cell in ws.Range("A1:A3").Cells())
            {
                Assert.That(cell.TryGetValue(out double? value), Is.True);
            }

            Assert.That(ws.Cell("A4").TryGetValue(out double? _), Is.False);
        }

        [Test]
        public void SetCellValueToGuid()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.AddWorksheet("Sheet1");
            var guid = Guid.NewGuid();
            ws.FirstCell().Value = guid;
            Assert.That(ws.FirstCell().DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.FirstCell().Value, Is.EqualTo(guid.ToString()));
            Assert.That(ws.FirstCell().GetString(), Is.EqualTo(guid.ToString()));

            guid = Guid.NewGuid();
            ws.FirstCell().SetValue(guid);
            Assert.That(ws.FirstCell().DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.FirstCell().Value, Is.EqualTo(guid.ToString()));
            Assert.That(ws.FirstCell().GetString(), Is.EqualTo(guid.ToString()));
        }

        [Test]
        public void SetCellValueToEnum()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.AddWorksheet("Sheet1");
            var dataType = XLDataType.Number;
            ws.FirstCell().Value = dataType;
            Assert.That(ws.FirstCell().DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.FirstCell().Value, Is.EqualTo(dataType.ToString()));
            Assert.That(ws.FirstCell().GetString(), Is.EqualTo(dataType.ToString()));

            dataType = XLDataType.TimeSpan;
            ws.FirstCell().SetValue(dataType);
            Assert.That(ws.FirstCell().DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.FirstCell().Value, Is.EqualTo(dataType.ToString()));
            Assert.That(ws.FirstCell().GetString(), Is.EqualTo(dataType.ToString()));
        }

        [Test]
        public void SetCellValueToRange()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.AddWorksheet("Sheet1");

            ws.Cell("A1").SetValue(2)
                .CellRight().SetValue(3)
                .CellRight().SetValue(5)
                .CellRight().SetValue(7);

            var range = ws.Range("1:1");

            ws.Cell("B2").Value = range;

            Assert.That(ws.Cell("B2").Value, Is.EqualTo(2));
            Assert.That(ws.Cell("C2").Value, Is.EqualTo(3));
            Assert.That(ws.Cell("D2").Value, Is.EqualTo(5));
            Assert.That(ws.Cell("E2").Value, Is.EqualTo(7));
        }

        [Test]
        public void ValueSetToEmptyString()
        {
            var expected = string.Empty;

            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Value = new DateTime(2000, 1, 2);
            cell.Value = string.Empty;
            Assert.That(cell.GetString(), Is.EqualTo(expected));
            Assert.That(cell.Value, Is.EqualTo(expected));

            cell.Value = new DateTime(2000, 1, 2);
            cell.SetValue(string.Empty);
            Assert.That(cell.GetString(), Is.EqualTo(expected));
            Assert.That(cell.Value, Is.EqualTo(expected));
        }

        [Test]
        public void ValueSetToNull()
        {
            var expected = string.Empty;

            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.Value = new DateTime(2000, 1, 2);
            cell.Value = null;
            Assert.That(cell.GetString(), Is.EqualTo(expected));
            Assert.That(cell.Value, Is.EqualTo(expected));

            cell.Value = new DateTime(2000, 1, 2);
            cell.SetValue(null as string);
            Assert.That(cell.GetString(), Is.EqualTo(expected));
            Assert.That(cell.Value, Is.EqualTo(expected));
        }

        [Test]
        public void ValueSetDateWithShortUserDateFormat()
        {
            // For this test to make sense, user's local date format should be dd/MM/yy (note without the 2 century digits)
            // What happened previously was that the century digits got lost in .ToString() conversion and wrong century was sometimes returned.
            var ci = new CultureInfo(CultureInfo.InvariantCulture.LCID);
            ci.DateTimeFormat.ShortDatePattern = "dd/MM/yy";
            Thread.CurrentThread.CurrentCulture = ci;
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            var cell = ws.Cell(1, 1);
            var expected = DateTime.Today.AddYears(20);
            cell.Value = expected;
            var actual = (DateTime)cell.Value;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void SetStringCellValues()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var cell = ws.FirstCell();

            object expected;

            var date = new DateTime(2018, 4, 18);
            expected = date.ToString(CultureInfo.CurrentCulture);
            cell.Value = expected;
            Assert.That(cell.DataType, Is.EqualTo(XLDataType.DateTime));
            Assert.That(cell.Value, Is.EqualTo(date));

            var b = true;
            expected = b.ToString(CultureInfo.CurrentCulture);
            cell.Value = expected;
            Assert.That(cell.DataType, Is.EqualTo(XLDataType.Boolean));
            Assert.That(cell.Value, Is.EqualTo(b));

            var ts = new TimeSpan(8, 12, 4);
            expected = ts.ToString();
            cell.Value = expected;
            Assert.That(cell.DataType, Is.EqualTo(XLDataType.TimeSpan));
            Assert.That(cell.Value, Is.EqualTo(ts));
        }

        [Test]
        public void SetStringValueTooLong()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            ws.FirstCell().Value = new DateTime(2018, 5, 15);

            ws.FirstCell().SetValue(new string('A', 32767));

            Assert.Throws<ArgumentOutOfRangeException>(() => ws.FirstCell().Value = new string('A', 32768));
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.FirstCell().SetValue(new string('A', 32768)));
        }

        [Test]
        [Culture("en-GB")]
        public void SetDateTime_in_Regular_and_Strict_Mode()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            var cell = ws.FirstCell() as XLCell;

            //Test non-strict mode
            cell.SetDateValue("1/1/2000");

            Assert.That(cell.Value, Is.EqualTo("36526"));
            Assert.That(DateTime.FromOADate(double.Parse(cell.Value as string)), Is.EqualTo(DateTime.Parse("1/1/2000")));

            //Test strict mode
            cell.SetDateValue("30000");

            Assert.That(cell.Value, Is.EqualTo("30000"));
            Assert.That(DateTime.FromOADate(double.Parse(cell.Value as string)), Is.EqualTo(DateTime.Parse("2/18/1982")));
        }

        [Test]
        public void SetDateOutOfRange()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-ZA");

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            ws.FirstCell().Value = 5;

            var date = XLCell.BaseDate.AddDays(-1);
            ws.FirstCell().Value = date;

            // Should default to string representation using current culture's date format
            Assert.That(ws.FirstCell().DataType, Is.EqualTo(XLDataType.Text));
            Assert.That(ws.FirstCell().Value, Is.EqualTo(date.ToString()));

            Assert.Throws<ArgumentException>(() => ws.FirstCell().SetValue(XLCell.BaseDate.AddDays(-1)));
        }

        [Test]
        public void SetCellValueWipesFormulas()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            ws.FirstCell().FormulaA1 = "=TODAY()";
            ws.FirstCell().Value = "hello world";
            Assert.That(ws.FirstCell().HasFormula, Is.False);

            ws.FirstCell().FormulaA1 = "=TODAY()";
            ws.FirstCell().SetValue("hello world");
            Assert.That(ws.FirstCell().HasFormula, Is.False);
        }

        [Test]
        public void CellValueLineWrapping()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            ws.FirstCell().Value = "hello world";
            Assert.That(ws.FirstCell().Style.Alignment.WrapText, Is.False);

            ws.FirstCell().Value = "hello\r\nworld";
            Assert.That(ws.FirstCell().Style.Alignment.WrapText, Is.True);

            ws.FirstCell().Style.Alignment.WrapText = false;

            ws.FirstCell().SetValue("hello world");
            Assert.That(ws.FirstCell().Style.Alignment.WrapText, Is.False);

            ws.FirstCell().SetValue("hello\r\nworld");
            Assert.That(ws.FirstCell().Style.Alignment.WrapText, Is.True);
        }

        [Test]
        public void TestInvalidXmlCharacters()
        {
            byte[] data;

            using (var stream = new MemoryStream())
            {
                using var wb = new XLWorkbook();
                wb.AddWorksheet("Sheet1").FirstCell().SetValue("\u0018");
                wb.SaveAs(stream);
                data = stream.ToArray();
            }

            using (var stream = new MemoryStream(data))
            {
                using var wb = new XLWorkbook(stream);
                Assert.That(wb.Worksheets.First().FirstCell().Value, Is.EqualTo("\u0018"));
            }
        }

        [Test]
        public void CanClearCellValueBySettingNullValue()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var cell = ws.FirstCell();

            cell.Value = "Test";
            Assert.That(cell.Value, Is.EqualTo("Test"));
            Assert.That(cell.DataType, Is.EqualTo(XLDataType.Text));

            string s = null;
            cell.SetValue(s);
            Assert.That(cell.Value, Is.EqualTo(string.Empty));

            cell.Value = "Test";
            cell.Value = null;
            Assert.That(cell.Value, Is.EqualTo(string.Empty));
        }

        [Test]
        public void CanClearDateTimeCellValue()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                var c = ws.FirstCell();
                c.SetValue(new DateTime(2017, 10, 08));
                Assert.That(c.DataType, Is.EqualTo(XLDataType.DateTime));
                Assert.That(c.Value, Is.EqualTo(new DateTime(2017, 10, 08)));

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                var c = ws.FirstCell();
                Assert.That(c.DataType, Is.EqualTo(XLDataType.DateTime));
                Assert.That(c.Value, Is.EqualTo(new DateTime(2017, 10, 08)));

                c.Clear();
                wb.Save();
            }

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                var c = ws.FirstCell();
                Assert.That(c.DataType, Is.EqualTo(XLDataType.Text));
                Assert.That(c.IsEmpty(), Is.True);
            }
        }

        [Test]
        public void ClearCellRemovesSparkline()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.Worksheets.Add("Sheet1");
            ws.SparklineGroups.Add("B1:B3", "C1:E3");

            ws.Cell("B1").Clear(XLClearOptions.All);
            ws.Cell("B2").Clear(XLClearOptions.Sparklines);

            Assert.That(ws.SparklineGroups.Single().Count(), Is.EqualTo(1));
            Assert.That(ws.Cell("B1").HasSparkline, Is.False);
            Assert.That(ws.Cell("B2").HasSparkline, Is.False);
            Assert.That(ws.Cell("B3").HasSparkline, Is.True);
        }

        [Test]
        public void CurrentRegion()
        {
            // Partially based on sample in https://github.com/ClosedXML/ClosedXML/issues/120
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            ws.Cell("B1").SetValue("x")
                .CellBelow().SetValue("x")
                .CellBelow().SetValue("x");

            ws.Cell("C1").SetValue("x")
                .CellBelow().SetValue("x")
                .CellBelow().SetValue("x");

            //Deliberately D2
            ws.Cell("D2").SetValue("x")
                .CellBelow().SetValue("x");

            ws.Cell("G1").SetValue("x")
                .CellBelow() // skip a cell
                .CellBelow().SetValue("x")
                .CellBelow().SetValue("x");

            // Deliberately H2
            ws.Cell("H2").SetValue("x")
                .CellBelow().SetValue("x")
                .CellBelow().SetValue("x");

            // A diagonal
            ws.Cell("E8").SetValue("x")
                .CellBelow().CellRight().SetValue("x")
                .CellBelow().CellRight().SetValue("x")
                .CellBelow().CellRight().SetValue("x")
                .CellBelow().CellRight().SetValue("x");

            Assert.That(ws.Cell("A10").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("A10:A10"));
            Assert.That(ws.Cell("B5").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("B5:B5"));
            Assert.That(ws.Cell("P1").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("P1:P1"));

            Assert.That(ws.Cell("D3").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("B1:D3"));
            Assert.That(ws.Cell("D4").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("B1:D4"));
            Assert.That(ws.Cell("E4").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("B1:E4"));

            foreach (var c in ws.Range("B1:D3").Cells())
            {
                Assert.That(c.CurrentRegion.RangeAddress.ToString(), Is.EqualTo("B1:D3"));
            }

            foreach (var c in ws.Range("A1:A3").Cells())
            {
                Assert.That(c.CurrentRegion.RangeAddress.ToString(), Is.EqualTo("A1:D3"));
            }

            Assert.That(ws.Cell("A4").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("A1:D4"));

            foreach (var c in ws.Range("E1:E3").Cells())
            {
                Assert.That(c.CurrentRegion.RangeAddress.ToString(), Is.EqualTo("B1:E3"));
            }

            Assert.That(ws.Cell("E4").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("B1:E4"));

            //// SECOND REGION
            foreach (var c in ws.Range("F1:F4").Cells())
            {
                Assert.That(c.CurrentRegion.RangeAddress.ToString(), Is.EqualTo("F1:H4"));
            }

            Assert.That(ws.Cell("F5").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("F1:H5"));

            //// DIAGONAL
            Assert.That(ws.Cell("E8").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:I12"));
            Assert.That(ws.Cell("F9").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:I12"));
            Assert.That(ws.Cell("G10").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:I12"));
            Assert.That(ws.Cell("H11").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:I12"));
            Assert.That(ws.Cell("I12").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:I12"));

            Assert.That(ws.Cell("G9").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:I12"));
            Assert.That(ws.Cell("F10").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:I12"));

            Assert.That(ws.Cell("D7").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("D7:I12"));
            Assert.That(ws.Cell("J13").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:J13"));
        }

        // https://github.com/ClosedXML/ClosedXML/issues/630
        [Test]
        public void ConsiderEmptyValueAsNumericInSumFormula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            ws.Cell("A1").SetValue("Empty");
            ws.Cell("A2").SetValue("Numeric");
            ws.Cell("A3").SetValue("Copy of numeric");

            ws.Cell("B2").SetFormulaA1("=B1");
            ws.Cell("B3").SetFormulaA1("=B2");

            ws.Cell("C2").SetFormulaA1("=SUM(C1)");
            ws.Cell("C3").SetFormulaA1("=C2");

            var b1 = ws.Cell("B1").Value;
            var b2 = ws.Cell("B2").Value;
            var b3 = ws.Cell("B3").Value;

            Assert.That(b1, Is.EqualTo(""));
            Assert.That(b2, Is.EqualTo(0));
            Assert.That(b3, Is.EqualTo(0));

            var c1 = ws.Cell("C1").Value;
            var c2 = ws.Cell("C2").Value;
            var c3 = ws.Cell("C3").Value;

            Assert.That(c1, Is.EqualTo(""));
            Assert.That(c2, Is.EqualTo(0));
            Assert.That(c3, Is.EqualTo(0));
        }

        [Test]
        public void SetFormulaA1AffectsR1C1()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.FormulaR1C1 = "R[1]C";

            cell.FormulaA1 = "B2";

            Assert.That(cell.FormulaR1C1, Is.EqualTo("R[1]C[1]"));
        }

        [Test]
        public void SetFormulaR1C1AffectsA1()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.FormulaA1 = "A2";

            cell.FormulaR1C1 = "R[1]C[1]";

            Assert.That(cell.FormulaA1, Is.EqualTo("B2"));
        }

        [Test]
        public void FormulaWithCircularReferenceFails()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var A1 = ws.Cell("A1");
            var A2 = ws.Cell("A2");
            A1.FormulaA1 = "A2 + 1";
            A2.FormulaA1 = "A1 + 1";

            Assert.Throws<InvalidOperationException>(() =>
            {
                _ = A1.Value;
            });
            Assert.Throws<InvalidOperationException>(() =>
            {
                _ = A2.Value;
            });
        }

        [Test]
        public void InvalidFormulaShiftProducesREF()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = 1;
                ws.Cell("B1").Value = 2;
                ws.Cell("B2").FormulaA1 = "=A1+B1";

                Assert.That(ws.Cell("B2").Value, Is.EqualTo(3));

                ws.Range("A2").Value = ws.Range("B2");
                var fA2 = ws.Cell("A2").FormulaA1;

                wb.SaveAs(ms);

                Assert.That(fA2, Is.EqualTo("#REF!+A1"));
            }

            using (var wb2 = new XLWorkbook(ms))
            {
                var fA2 = wb2.Worksheets.First().Cell("A2").FormulaA1;
                Assert.That(fA2, Is.EqualTo("#REF!+A1"));
            }
        }

        [Test]
        public void FormulaWithCircularReferenceFails2()
        {
            using var xLWorkbook = new XLWorkbook();
            var cell = xLWorkbook.Worksheets.Add("Sheet1").FirstCell();
            cell.FormulaA1 = "A1";
            Assert.Throws<InvalidOperationException>(() =>
            {
                _ = cell.Value;
            });
        }

        [Test]
        public void TryGetValueFormulaEvaluation()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var A1 = ws.Cell("A1");
            var A2 = ws.Cell("A2");
            var A3 = ws.Cell("A3");
            A1.FormulaA1 = "A2 + 1";
            A2.FormulaA1 = "A1 + 1";

            Assert.That(A1.TryGetValue(out string _), Is.False);
            Assert.That(A2.TryGetValue(out string _), Is.False);
            Assert.That(A3.TryGetValue(out string _), Is.True);
        }

        [Test]
        public void SetValue_IEnumerable()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            object[] values = { "Text", 45, DateTime.Today, true, "More text" };

            ws.FirstCell().SetValue(values);

            Assert.That(ws.FirstCell().GetString(), Is.EqualTo("Text"));
            Assert.That(ws.Cell("A2").GetDouble(), Is.EqualTo(45));
            Assert.That(ws.Cell("A3").GetDateTime(), Is.EqualTo(DateTime.Today));
            Assert.That(ws.Cell("A4").GetBoolean(), Is.EqualTo(true));
            Assert.That(ws.Cell("A5").GetString(), Is.EqualTo("More text"));
            Assert.That(ws.Cell("A6").IsEmpty(), Is.True);
        }

        [Test]
        public void ToStringNoFormatString()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var c = ws.FirstCell().CellBelow(2).CellRight(3);

            Assert.That(c.ToString(), Is.EqualTo("D3"));
        }

        [Test]
        [TestCase("D3", "A")]
        [TestCase("YEAR(DATE(2018, 1, 1))", "F")]
        [TestCase("YEAR(DATE(2018, 1, 1))", "f")]
        [TestCase("0000.00", "NF")]
        [TestCase("0000.00", "nf")]
        [TestCase("FFFF0000", "fg")]
        [TestCase("Color Theme: Accent5, Tint: 0", "BG")]
        [TestCase("2018.00", "v")]
        public void ToStringFormatString(string expected, string format)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var c = ws.FirstCell().CellBelow(2).CellRight(3);

            var formula = "YEAR(DATE(2018, 1, 1))";
            c.FormulaA1 = formula;

            var numberFormat = "0000.00";
            c.Style.NumberFormat.Format = numberFormat;

            c.Style.Font.FontColor = XLColor.Red;
            c.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent5);

            Assert.That(c.ToString(format), Is.EqualTo(expected));

            Assert.Throws<FormatException>(() => c.ToString("dummy"));
        }

        [Test]
        public void ToStringInvalidFormat()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var c = ws.FirstCell();

            Assert.Throws<FormatException>(() => c.ToString("dummy"));
        }
    }
}