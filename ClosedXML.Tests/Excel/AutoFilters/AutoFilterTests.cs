using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;

namespace ClosedXML.Tests.Excel.AutoFilters
{
    [TestFixture]
    public class AutoFilterTests
    {
        [Test]
        public void AutoFilterExpandsWithTable()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.FirstCell().SetValue("Categories")
                .CellBelow().SetValue("1")
                .CellBelow().SetValue("2");

            var table = ws.RangeUsed().CreateTable();

            var listOfArr = new List<int>
            {
                3,
                4,
                5,
                6
            };

            table.DataRange.InsertRowsBelow(listOfArr.Count - table.DataRange.RowCount());
            table.DataRange.FirstCell().InsertData(listOfArr);

            Assert.That(table.AutoFilter.Range.RangeAddress.ToStringRelative(), Is.EqualTo("A1:A5"));
            Assert.That(table.AutoFilter.VisibleRows.Count(), Is.EqualTo(5));
        }

        [Test]
        public void AutoFilterSortWhenNotInFirstRow()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Cell(3, 3).SetValue("Names")
                .CellBelow().SetValue("Manuel")
                .CellBelow().SetValue("Carlos")
                .CellBelow().SetValue("Dominic");
            ws.RangeUsed().SetAutoFilter().Sort();
            Assert.That(ws.Cell(4, 3).GetString(), Is.EqualTo("Carlos"));
        }

        [Test]
        public void CanClearAutoFilter()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("AutoFilter");
            ws.Cell("A1").Value = "Names";
            ws.Cell("A2").Value = "John";
            ws.Cell("A3").Value = "Hank";
            ws.Cell("A4").Value = "Dagny";

            ws.AutoFilter.Clear(); // We should be able to clear a filter even if it hasn't been set.
            Assert.That(!ws.AutoFilter.IsEnabled);

            ws.RangeUsed().SetAutoFilter();
            Assert.That(ws.AutoFilter.IsEnabled);

            ws.AutoFilter.Clear();
            Assert.That(!ws.AutoFilter.IsEnabled);
        }

        [Test]
        public void CanClearAutoFilter2()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("AutoFilter");
            ws.Cell("A1").Value = "Names";
            ws.Cell("A2").Value = "John";
            ws.Cell("A3").Value = "Hank";
            ws.Cell("A4").Value = "Dagny";

            ws.SetAutoFilter(false);
            Assert.That(!ws.AutoFilter.IsEnabled);

            ws.RangeUsed().SetAutoFilter();
            Assert.That(ws.AutoFilter.IsEnabled);

            ws.RangeUsed().SetAutoFilter(false);
            Assert.That(!ws.AutoFilter.IsEnabled);
        }

        [Test]
        public void CanCopyAutoFilterToNewSheetOnNewWorkbook()
        {
            using var ms1 = new MemoryStream();
            using var ms2 = new MemoryStream();
            using (var wb1 = new XLWorkbook())
            using (var wb2 = new XLWorkbook())
            {
                var ws = wb1.Worksheets.Add("AutoFilter");
                ws.Cell("A1").Value = "Names";
                ws.Cell("A2").Value = "John";
                ws.Cell("A3").Value = "Hank";
                ws.Cell("A4").Value = "Dagny";

                ws.RangeUsed().SetAutoFilter();

                wb1.SaveAs(ms1);

                ws.CopyTo(wb2, ws.Name);
                wb2.SaveAs(ms2);
            }

            using (var wb2 = new XLWorkbook(ms2))
            {
                Assert.That(wb2.Worksheets.First().AutoFilter.IsEnabled, Is.True);
            }
        }

        [Test]
        public void CannotAddAutoFilterOverExistingTable()
        {
            using var wb = new XLWorkbook();

            var data = Enumerable.Range(1, 10).Select(i => new
            {
                Index = i,
                String = $"String {i}"
            });

            var ws = wb.AddWorksheet();
            ws.FirstCell().InsertTable(data);

            Assert.Throws<InvalidOperationException>(() => ws.RangeUsed().SetAutoFilter());
        }

        [Test]
        [TestCase("A1:A4")]
        [TestCase("A1:B4")]
        [TestCase("A1:C4")]
        public void AutoFilterRangeRemainsValidOnInsertColumn(string rangeAddress)
        {
            //Arrange
            using var ms1 = new MemoryStream();
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("AutoFilter");
            ws.Cell("A1").Value = "Ids";
            ws.Cell("B1").Value = "Names";
            ws.Cell("B2").Value = "John";
            ws.Cell("B3").Value = "Hank";
            ws.Cell("B4").Value = "Dagny";
            ws.Cell("C1").Value = "Phones";

            ws.Range("B1:B4").SetAutoFilter(true);

            //Act
            var range = ws.Range(rangeAddress);
            range.InsertColumnsBefore(1);

            //Assert
            Assert.That(ws.AutoFilter.Range.RangeAddress.IsValid, Is.True);
        }

        [Test]
        public void AutoFilterVisibleRows()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Cell(3, 3).SetValue("Names")
                .CellBelow().SetValue("Manuel")
                .CellBelow().SetValue("Carlos")
                .CellBelow().SetValue("Dominic");

            var autoFilter = ws.RangeUsed()
                .SetAutoFilter();

            autoFilter.Column(1).AddFilter("Carlos");

            Assert.That(ws.Cell(5, 3).GetString(), Is.EqualTo("Carlos"));
            Assert.That(autoFilter.VisibleRows.Count(), Is.EqualTo(2));
            Assert.That(autoFilter.VisibleRows.First().WorksheetRow().RowNumber(), Is.EqualTo(3));
            Assert.That(autoFilter.VisibleRows.Last().WorksheetRow().RowNumber(), Is.EqualTo(5));
        }

        [Test]
        public void ReapplyAutoFilter()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet1");

            ws.Cell(3, 3).SetValue("Names")
                .CellBelow().SetValue("Manuel")
                .CellBelow().SetValue("Carlos")
                .CellBelow().SetValue("Dominic")
                .CellBelow().SetValue("Jose");

            var autoFilter = ws.RangeUsed()
                .SetAutoFilter();

            autoFilter.Column(1).AddFilter("Carlos");

            Assert.That(autoFilter.HiddenRows.Count(), Is.EqualTo(3));

            // Unhide the rows so that the table is out of sync with the filter
            autoFilter.HiddenRows.ForEach(r => r.WorksheetRow().Unhide());
            Assert.That(autoFilter.HiddenRows.Any(), Is.False);

            autoFilter.Reapply();
            Assert.That(autoFilter.HiddenRows.Count(), Is.EqualTo(3));
        }

        [Test]
        public void CanLoadAutoFilterWithThousandsSeparator()
        {
            var backupCulture = Thread.CurrentThread.CurrentCulture;

            try
            {
                // Set thread culture to French, which should format numbers using a space as thousands separator
                var culture = CultureInfo.CreateSpecificCulture("fr-FR");
                // but use a period instead of a comma as for decimal separator and space as group separator
                culture.NumberFormat.CurrencyDecimalSeparator = ".";
                culture.NumberFormat.CurrencyGroupSeparator = " ";

                Thread.CurrentThread.CurrentCulture = culture;

                using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\AutoFilter\AutoFilterWithThousandsSeparator.xlsx")))
                using (var wb = new XLWorkbook(stream))
                {
                    var ws = wb.Worksheets.First();
                    Assert.That((ws.AutoFilter as XLAutoFilter).Filters.First().Value.FirstOrDefault().Value, Is.EqualTo(10000));
                    Assert.That(ws.AutoFilter.VisibleRows.Count(), Is.EqualTo(2));

                    ws.AutoFilter.Reapply();
                    Assert.That(ws.AutoFilter.VisibleRows.Count(), Is.EqualTo(2));
                }

                Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");

                using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\AutoFilter\AutoFilterWithThousandsSeparator.xlsx")))
                using (var wb = new XLWorkbook(stream))
                {
                    var ws = wb.Worksheets.First();
                    Assert.That((ws.AutoFilter as XLAutoFilter).Filters.First().Value.FirstOrDefault().Value, Is.EqualTo("10 000.00"));

                    _ = ws.AutoFilter.VisibleRows.Select(r => r.FirstCell().Value).ToList();
                    Assert.That(ws.AutoFilter.VisibleRows.Count(), Is.EqualTo(2));

                    ws.AutoFilter.Reapply();
                    Assert.That(ws.AutoFilter.VisibleRows.Count(), Is.EqualTo(1));
                }
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = backupCulture;
            }
        }
    }
}