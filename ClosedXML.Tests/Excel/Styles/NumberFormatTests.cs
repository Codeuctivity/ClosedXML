using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.Styles
{
    public class NumberFormatTests
    {
        [Test]
        public void PreserveCellFormat()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            using var table = new DataTable();
            table.Columns.Add("Date", typeof(DateTime));

            for (var i = 0; i < 10; i++)
            {
                table.Rows.Add(new DateTime(2017, 1, 1).AddMonths(i));
            }

            ws.Column(1).Style.NumberFormat.Format = "yy-MM-dd";
            ws.Cell("A1").InsertData(table);
            Assert.That(ws.Cell("A5").Style.DateFormat.Format, Is.EqualTo("yy-MM-dd"));

            ws.Row(1).Style.NumberFormat.Format = "yy-MM-dd";
            ws.Cell("A1").InsertData(table.Rows, true);
            Assert.That(ws.Cell("E1").Style.DateFormat.Format, Is.EqualTo("yy-MM-dd"));
        }

        [Test]
        public void TestExcelNumberFormats()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var c = ws.FirstCell()
                .SetValue(41573.875)
                .SetDataType(XLDataType.DateTime);

            c.Style.NumberFormat.SetFormat("m/d/yy\\ h:mm;@");

            Assert.That(c.GetFormattedString(), Is.EqualTo("10/26/13 21:00"));
        }

        [Test]
        public void ReadAndWriteColumnNumberFormat()
        {
            using var memoryStream = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var sourceColumn = ws.Column(1);
                sourceColumn.Style.NumberFormat.Format = "0.000";
                wb.SaveAs(memoryStream);
            }

            memoryStream.Position = 0;

            using (var wb = new XLWorkbook(memoryStream))
            {
                var column = wb.Worksheets.Single().Column(1);
                Assert.That(column.Style.NumberFormat.Format, Is.EqualTo("0.000"));
            }
        }

        [Test]
        public void XLNumberFormatKey_GetHashCode_IsCaseSensitive()
        {
            var numberFormatKey1 = new XLNumberFormatKey { Format = "MM" };
            var numberFormatKey2 = new XLNumberFormatKey { Format = "mm" };

            Assert.That(numberFormatKey2.GetHashCode(), Is.Not.EqualTo(numberFormatKey1.GetHashCode()));
        }

        [Test]
        public void XLNumberFormatKey_Equals_IsCaseSensitive()
        {
            var numberFormatKey1 = new XLNumberFormatKey { Format = "MM" };
            var numberFormatKey2 = new XLNumberFormatKey { Format = "mm" };

            Assert.That(numberFormatKey1.Equals(numberFormatKey2), Is.False);
        }
    }
}