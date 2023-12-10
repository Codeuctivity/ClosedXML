using ClosedXML.Excel;
using ClosedXML.Tests.Utils;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace ClosedXML.Tests.Excel.PivotTables
{
    [TestFixture]
    public class XLPivotTableTests
    {
        [Test]
        public void PivotTables()
        {
            Assert.DoesNotThrow(() =>
            {
                using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx"));
                using var wb = new XLWorkbook(stream);
                var ws = wb.Worksheet("PastrySalesData");
                var table = ws.Table("PastrySalesData");
                var ptSheet = wb.Worksheets.Add("BlankPivotTable");
                ptSheet.PivotTables.Add("pvt", ptSheet.Cell(1, 1), table);

                using var ms = new MemoryStream();
                wb.SaveAs(ms, true);
            });
        }

        [Test]
        public void TestPivotTableVersioningAttributes()
        {
            var expectedFilePath = @"Other\PivotTableReferenceFiles\VersioningAttributes\outputfile.xlsx";

            // Pivot table definitions in input file has created and refreshed versioning attributes = 3
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\VersioningAttributes\inputfile.xlsx"));
            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook(stream);

                var data = wb.Worksheet("Data");

                var pt = data.RangeUsed().CreatePivotTable(wb.AddWorksheet("pvt2").FirstCell(), "pvt2");

                pt.ColumnLabels.Add("Sex");
                pt.RowLabels.Add("FullName");
                pt.Values.Add("Id", "Count of Id").SetSummaryFormula(XLPivotSummary.Count);

                //Uncomment to replace expectation running.net6.0,
                //var expectedFileInVsSolution = Path.GetFullPath(Path.Combine("../../../", "Resource", expectedFilePath));
                //wb.SaveAs(expectedFileInVsSolution);

                return wb;
                // Pivot table definitions in output file has created and refreshed versioning attributes = 5
            }, expectedFilePath);
        }

        [Test]
        public void PivotTableOptionsSaveTest()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet("PastrySalesData");
            var table = ws.Table("PastrySalesData");
            var ptSheet = wb.Worksheets.Add("BlankPivotTable");
            var pt = ptSheet.PivotTables.Add("pvtOptionsTest", ptSheet.Cell(1, 1), table);

            pt.ColumnHeaderCaption = "clmn header";
            pt.RowHeaderCaption = "row header";

            pt.AutofitColumns = true;
            pt.PreserveCellFormatting = false;
            pt.ShowGrandTotalsColumns = true;
            pt.ShowGrandTotalsRows = true;
            pt.UseCustomListsForSorting = false;
            pt.ShowExpandCollapseButtons = false;
            pt.ShowContextualTooltips = false;
            pt.DisplayCaptionsAndDropdowns = false;
            pt.RepeatRowLabels = true;
            pt.SaveSourceData = false;
            pt.EnableShowDetails = false;
            pt.ShowColumnHeaders = false;
            pt.ShowRowHeaders = false;

            pt.MergeAndCenterWithLabels = true; // MergeItem
            pt.RowLabelIndent = 12; // Indent
            pt.FilterAreaOrder = XLFilterAreaOrder.OverThenDown; // PageOverThenDown
            pt.FilterFieldsPageWrap = 14; // PageWrap
            pt.ErrorValueReplacement = "error test"; // ErrorCaption
            pt.EmptyCellReplacement = "empty test"; // MissingCaption

            pt.FilteredItemsInSubtotals = true; // Subtotal filtered page items
            pt.AllowMultipleFilters = false; // MultipleFieldFilters

            pt.ShowPropertiesInTooltips = false;
            pt.ClassicPivotTableLayout = true;
            pt.ShowEmptyItemsOnRows = true;
            pt.ShowEmptyItemsOnColumns = true;
            pt.DisplayItemLabels = false;
            pt.SortFieldsAtoZ = true;

            pt.PrintExpandCollapsedButtons = true;
            pt.PrintTitles = true;

            // TODO pt.RefreshDataOnOpen = false;
            pt.ItemsToRetainPerField = XLItemsToRetain.Max;
            pt.EnableCellEditing = true;
            pt.ShowValuesRow = true;
            pt.ShowRowStripes = true;
            pt.ShowColumnStripes = true;
            pt.Theme = XLPivotTableTheme.PivotStyleDark13;

            pt.DataCaption = "Test Caption Values";
            pt.GrandTotalCaption = "Test Grand Total Caption";

            using var ms = new MemoryStream();
            wb.SaveAs(ms, true);

            ms.Position = 0;

            using var wbassert = new XLWorkbook(ms);
            var wsassert = wbassert.Worksheet("BlankPivotTable");
            var ptassert = wsassert.PivotTable("pvtOptionsTest");
            Assert.That(ptassert, Is.Not.EqualTo(null), "name save failure");
            Assert.That(ptassert.ColumnHeaderCaption, Is.EqualTo("clmn header"), "ColumnHeaderCaption save failure");
            Assert.That(ptassert.RowHeaderCaption, Is.EqualTo("row header"), "RowHeaderCaption save failure");
            Assert.That(ptassert.MergeAndCenterWithLabels, Is.EqualTo(true), "MergeAndCenterWithLabels save failure");
            Assert.That(ptassert.RowLabelIndent, Is.EqualTo(12), "RowLabelIndent save failure");
            Assert.That(ptassert.FilterAreaOrder, Is.EqualTo(XLFilterAreaOrder.OverThenDown), "FilterAreaOrder save failure");
            Assert.That(ptassert.FilterFieldsPageWrap, Is.EqualTo(14), "FilterFieldsPageWrap save failure");
            Assert.That(ptassert.ErrorValueReplacement, Is.EqualTo("error test"), "ErrorValueReplacement save failure");
            Assert.That(ptassert.EmptyCellReplacement, Is.EqualTo("empty test"), "EmptyCellReplacement save failure");
            Assert.That(ptassert.AutofitColumns, Is.EqualTo(true), "AutofitColumns save failure");
            Assert.That(ptassert.PreserveCellFormatting, Is.EqualTo(false), "PreserveCellFormatting save failure");
            Assert.That(ptassert.ShowGrandTotalsRows, Is.EqualTo(true), "ShowGrandTotalsRows save failure");
            Assert.That(ptassert.ShowGrandTotalsColumns, Is.EqualTo(true), "ShowGrandTotalsColumns save failure");
            Assert.That(ptassert.FilteredItemsInSubtotals, Is.EqualTo(true), "FilteredItemsInSubtotals save failure");
            Assert.That(ptassert.AllowMultipleFilters, Is.EqualTo(false), "AllowMultipleFilters save failure");
            Assert.That(ptassert.UseCustomListsForSorting, Is.EqualTo(false), "UseCustomListsForSorting save failure");
            Assert.That(ptassert.ShowExpandCollapseButtons, Is.EqualTo(false), "ShowExpandCollapseButtons save failure");
            Assert.That(ptassert.ShowContextualTooltips, Is.EqualTo(false), "ShowContextualTooltips save failure");
            Assert.That(ptassert.ShowPropertiesInTooltips, Is.EqualTo(false), "ShowPropertiesInTooltips save failure");
            Assert.That(ptassert.DisplayCaptionsAndDropdowns, Is.EqualTo(false), "DisplayCaptionsAndDropdowns save failure");
            Assert.That(ptassert.ClassicPivotTableLayout, Is.EqualTo(true), "ClassicPivotTableLayout save failure");
            Assert.That(ptassert.ShowEmptyItemsOnRows, Is.EqualTo(true), "ShowEmptyItemsOnRows save failure");
            Assert.That(ptassert.ShowEmptyItemsOnColumns, Is.EqualTo(true), "ShowEmptyItemsOnColumns save failure");
            Assert.That(ptassert.DisplayItemLabels, Is.EqualTo(false), "DisplayItemLabels save failure");
            Assert.That(ptassert.SortFieldsAtoZ, Is.EqualTo(true), "SortFieldsAtoZ save failure");
            Assert.That(ptassert.PrintExpandCollapsedButtons, Is.EqualTo(true), "PrintExpandCollapsedButtons save failure");
            Assert.That(ptassert.RepeatRowLabels, Is.EqualTo(true), "RepeatRowLabels save failure");
            Assert.That(ptassert.PrintTitles, Is.EqualTo(true), "PrintTitles save failure");
            Assert.That(ptassert.SaveSourceData, Is.EqualTo(false), "SaveSourceData save failure");
            Assert.That(ptassert.EnableShowDetails, Is.EqualTo(false), "EnableShowDetails save failure");
            // TODO Assert.AreEqual(false, ptassert.RefreshDataOnOpen, "RefreshDataOnOpen save failure");
            Assert.That(ptassert.ItemsToRetainPerField, Is.EqualTo(XLItemsToRetain.Max), "ItemsToRetainPerField save failure");
            Assert.That(ptassert.EnableCellEditing, Is.EqualTo(true), "EnableCellEditing save failure");
            Assert.That(ptassert.Theme, Is.EqualTo(XLPivotTableTheme.PivotStyleDark13), "Theme save failure");
            Assert.That(ptassert.ShowValuesRow, Is.EqualTo(true), "ShowValuesRow save failure");
            Assert.That(ptassert.ShowRowHeaders, Is.EqualTo(false), "ShowRowHeaders save failure");
            Assert.That(ptassert.ShowColumnHeaders, Is.EqualTo(false), "ShowColumnHeaders save failure");
            Assert.That(ptassert.ShowRowStripes, Is.EqualTo(true), "ShowRowStripes save failure");
            Assert.That(ptassert.ShowColumnStripes, Is.EqualTo(true), "ShowColumnStripes save failure");
            Assert.That(ptassert.DataCaption, Is.EqualTo("Test Caption Values"), "DataCaption save failure");
            Assert.That(ptassert.GrandTotalCaption, Is.EqualTo("Test Grand Total Caption"), "GrandTotalCaption save failure");
        }

        [Test]
        public void PivotTableOptionsSaveTest_CaptionsNotSet()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet("PastrySalesData");
            var table = ws.Table("PastrySalesData");
            var ptSheet = wb.Worksheets.Add("BlankPivotTable");
            var pt = ptSheet.PivotTables.Add("pvtOptionsTest", ptSheet.Cell(1, 1), table);

            pt.DataCaption = null;
            pt.GrandTotalCaption = null;

            using var ms = new MemoryStream();
            wb.SaveAs(ms, true);

            ms.Position = 0;

            using var wbassert = new XLWorkbook(ms);
            var wsassert = wbassert.Worksheet("BlankPivotTable");
            var ptassert = wsassert.PivotTable("pvtOptionsTest");

            Assert.That(ptassert.DataCaption, Is.EqualTo("Values"), "DataCaption save failure");
            Assert.That(ptassert.GrandTotalCaption, Is.EqualTo(null), "GrandTotalCaption save failure");
        }

        [TestCase(true)]
        [TestCase(false)]
        public void PivotFieldOptionsSaveTest(bool withDefaults)
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet("PastrySalesData");
            var table = ws.Table("PastrySalesData");

            var ptSheet = wb.Worksheets.Add("pvtFieldOptionsTest");
            var pt = ptSheet.PivotTables.Add("pvtFieldOptionsTest", ptSheet.Cell(1, 1), table);

            var field = pt.RowLabels.Add("Name")
                .SetSubtotalCaption("Test caption")
                .SetCustomName("Test name");
            SetFieldOptions(field, withDefaults);

            pt.ColumnLabels.Add("Month");
            pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

            using var ms = new MemoryStream();
            wb.SaveAs(ms, true);

            ms.Position = 0;

            using var wbassert = new XLWorkbook(ms);
            var wsassert = wbassert.Worksheet("pvtFieldOptionsTest");
            var ptassert = wsassert.PivotTable("pvtFieldOptionsTest");
            var pfassert = ptassert.RowLabels.Get("Name");
            Assert.That(pfassert, Is.Not.EqualTo(null), "name save failure");
            Assert.That(pfassert.SubtotalCaption, Is.EqualTo("Test caption"), "SubtotalCaption save failure");
            Assert.That(pfassert.CustomName, Is.EqualTo("Test name"), "CustomName save failure");
            AssertFieldOptions(pfassert, withDefaults);
        }

        [Test]
        public void PivotTableStyleFormatsTest()
        {
            using var ms = new MemoryStream();
            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx")))
            using (var wbSource = new XLWorkbook(stream))
            using (var wbDestination = new XLWorkbook())
            {
                var ws = wbSource.Worksheet("PastrySalesData");
                wbDestination.AddWorksheet(ws);
                ws = wbDestination.Worksheet("PastrySalesData");

                var table = ws.Table("PastrySalesData");
                var ptSheet = wbDestination.Worksheets.Add("PivotTableStyleFormats");
                var pt = ptSheet.PivotTables.Add("pvtStyleFormats", ptSheet.Cell(1, 1), table);
                pt.Layout = XLPivotLayout.Tabular;

                pt.SetSubtotals(XLPivotSubtotals.AtBottom);

                var monthPivotField = pt.ColumnLabels.Add("Month");

                var namePivotField = pt.RowLabels.Add("Name")
                    .SetSubtotalCaption("Test caption")
                    .SetCustomName("Test name")
                    .AddSubtotal(XLSubtotalFunction.Sum);

                ptSheet.SetTabActive();

                var numberOfOrdersPivotValue = pt.Values.Add("NumberOfOrders")
                    .SetSummaryFormula(XLPivotSummary.Sum);

                var qualityPivotValue = pt.Values.Add("Quality").SetSummaryFormula(XLPivotSummary.Sum);

                pt.StyleFormats.RowGrandTotalFormats.ForElement(XLPivotStyleFormatElement.All).Style.Font.FontColor = XLColor.VenetianRed;

                namePivotField.StyleFormats.Subtotal.Style.Fill.BackgroundColor = XLColor.Blue;
                monthPivotField.StyleFormats.Label.Style.Fill.BackgroundColor = XLColor.Amber;
                monthPivotField.StyleFormats.Header.Style.Font.FontColor = XLColor.Yellow;
                namePivotField.StyleFormats.DataValuesFormat
                    .AndWith(monthPivotField, v => v.ToString() == "May")
                    .ForValueField(numberOfOrdersPivotValue)
                    .Style.Font.FontColor = XLColor.Green;

                wbDestination.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheet("PivotTableStyleFormats");
                var pt = ws.PivotTable("pvtStyleFormats").CastTo<XLPivotTable>();

                Assert.That(pt.StyleFormats.ColumnGrandTotalFormats.Count(), Is.EqualTo(0));

                Assert.That(pt.StyleFormats.RowGrandTotalFormats, Is.Not.Null);
                Assert.That(pt.StyleFormats.RowGrandTotalFormats.Count(), Is.EqualTo(1));
                Assert.That(pt.StyleFormats.RowGrandTotalFormats.First().AppliesTo, Is.EqualTo(XLPivotStyleFormatElement.All));
                Assert.That(pt.StyleFormats.RowGrandTotalFormats.ForElement(XLPivotStyleFormatElement.All).Style.Font.FontColor, Is.EqualTo(XLColor.VenetianRed));

                var namePivotField = pt.RowLabels.Get("Name");
                var monthPivotField = pt.ColumnLabels.Get("Month");
                var numberOfOrdersPivotValue = pt.Values.Get("NumberOfOrders");

                Assert.That(namePivotField.StyleFormats.Label.Style, Is.EqualTo(XLStyle.Default));
                Assert.That(namePivotField.StyleFormats.Subtotal.Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Blue));

                Assert.That(monthPivotField.StyleFormats.Subtotal.Style, Is.EqualTo(XLStyle.Default));
                Assert.That(monthPivotField.StyleFormats.Label.Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Amber));
                Assert.That(monthPivotField.StyleFormats.Header.Style.Font.FontColor, Is.EqualTo(XLColor.Yellow));

                var nameDataValuesFormat = namePivotField.StyleFormats.DataValuesFormat as XLPivotValueStyleFormat;
                Assert.That(nameDataValuesFormat.FieldReferences.Count, Is.EqualTo(2));

                Assert.That(nameDataValuesFormat.FieldReferences.First().CastTo<PivotLabelFieldReference>().PivotField, Is.EqualTo(monthPivotField));

                Assert.That(nameDataValuesFormat.FieldReferences.Last().CastTo<PivotValueFieldReference>().Value, Is.EqualTo(numberOfOrdersPivotValue.CustomName));

                wb.Save();
            }
        }

        [Test]
        public void CopyPivotTableTests()
        {
            using var ms = new MemoryStream();
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx"));
            using var wb = new XLWorkbook(stream);
            var ws1 = wb.Worksheet("pvt1");
            var pt1 = ws1.PivotTables.First() as XLPivotTable;

            Assert.Throws<InvalidOperationException>(() => pt1.CopyTo(pt1.TargetCell));

            var pt2 = pt1.CopyTo(ws1.Cell("AB100")) as XLPivotTable;

            AssertPivotTablesAreEqual(pt1, pt2, compareName: false);

            var ws2 = wb.AddWorksheet("Copy Of pvt1");
            AssertPivotTablesAreEqual(pt1, pt1.CopyTo(ws2.FirstCell()) as XLPivotTable, compareName: true);

            using var wb2 = new XLWorkbook();
            wb.Worksheet("PastrySalesData").CopyTo(wb2);

            AssertPivotTablesAreEqual(pt1, pt1.CopyTo(wb2.AddWorksheet("pvt").FirstCell()) as XLPivotTable, compareName: true);
        }

        private void AssertPivotTablesAreEqual(XLPivotTable original, XLPivotTable copy, bool compareName)
        {
            Assert.That(copy.Guid, Is.Not.EqualTo(original.Guid));
            Assert.That(original.Name.Equals(copy.Name), Is.EqualTo(compareName));

            var comparer = new PivotTableComparer(compareName: compareName, compareRelId: false, compareTargetCellAddress: false);
            Assert.That(comparer.Equals(original, copy), Is.True);
        }

        private class Pastry
        {
            public Pastry(string name, int? code, int numberOfOrders, double quality, string month, DateTime? bakeDate)
            {
                Name = name;
                Code = code;
                NumberOfOrders = numberOfOrders;
                Quality = quality;
                Month = month;
                BakeDate = bakeDate;
            }

            public string Name { get; set; }
            public int? Code { get; }
            public int NumberOfOrders { get; set; }
            public double Quality { get; set; }
            public string Month { get; set; }
            public DateTime? BakeDate { get; set; }
        }

        [Test]
        public void BlankPivotTableField()
        {
            using var wb = new XLWorkbook();

            var expectedFilePath = @"Other\PivotTableReferenceFiles\BlankPivotTableField\BlankPivotTableField.xlsx";
            using var ms = new MemoryStream();
            TestHelper.CreateAndCompare(() =>
            {
                // Based on .\ClosedXML\ClosedXML.Examples\PivotTables\PivotTables.cs
                // But with empty column for Month
                var pastries = new List<Pastry>
                {
                        new Pastry("Croissant", 101, 150, 60.2, "", new DateTime(2016, 04, 21)),
                        new Pastry("Croissant", 101, 250, 50.42, "", new DateTime(2016, 05, 03)),
                        new Pastry("Croissant", 101, 134, 22.12, "", new DateTime(2016, 06, 24)),
                        new Pastry("Doughnut", 102, 250, 89.99, "", new DateTime(2017, 04, 23)),
                        new Pastry("Doughnut", 102, 225, 70, "", new DateTime(2016, 05, 24)),
                        new Pastry("Doughnut", 102, 210, 75.33, "", new DateTime(2016, 06, 02)),
                        new Pastry("Bearclaw", 103, 134, 10.24, "", new DateTime(2016, 04, 27)),
                        new Pastry("Bearclaw", 103, 184, 33.33, "", new DateTime(2016, 05, 20)),
                        new Pastry("Bearclaw", 103, 124, 25, "", new DateTime(2017, 06, 05)),
                        new Pastry("Danish", 104, 394, -20.24, "", null),
                        new Pastry("Danish", 104, 190, 60, "", new DateTime(2017, 05, 08)),
                        new Pastry("Danish", 104, 221, 24.76, "", new DateTime(2016, 06, 21)),

                        // Deliberately add different casings of same string to ensure pivot table doesn't duplicate it.
                        new Pastry("Scone", 105, 135, 0, "", new DateTime(2017, 04, 22)),
                        new Pastry("SconE", 105, 122, 5.19, "", new DateTime(2017, 05, 03)),
                        new Pastry("SCONE", 105, 243, 44.2, "", new DateTime(2017, 06, 14)),

                        // For ContainsBlank and integer rows/columns test
                        new Pastry("Scone", null, 255, 18.4, "", null),
                };

                var sheet = wb.Worksheets.Add("PastrySalesData");
                // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                var table = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
                sheet.Cell("F11").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                sheet.Columns().AdjustToContents();

                IXLWorksheet ptSheet;
                IXLPivotTable pt;

                for (var i = 1; i <= 5; i++)
                {
                    // Add a new sheet for our pivot table
                    ptSheet = wb.Worksheets.Add("pvt" + i);

                    // Create the pivot table, using the data from the "PastrySalesData" table
                    pt = ptSheet.PivotTables.Add("pvt" + i, ptSheet.Cell(1, 1), table);

                    if (i == 1 || i == 4 || i == 5)
                    {
                        pt.ColumnLabels.Add("Name");
                    }
                    else if (i == 2 || i == 3)
                    {
                        pt.RowLabels.Add("Name");
                    }

                    if (i == 1 || i == 3)
                    {
                        pt.RowLabels.Add("Month");
                    }
                    else if (i == 2 || i == 4)
                    {
                        pt.ColumnLabels.Add("Month");
                    }
                    else if (i == 5)
                    {
                        pt.RowLabels.Add("BakeDate");
                    }

                    // The values in our table will come from the "NumberOfOrders" field
                    // The default calculation setting is a total of each row/column
                    pt.Values.Add("NumberOfOrders", "NumberOfOrdersPercentageOfBearclaw")
                    .ShowAsPercentageFrom("Name").And("Bearclaw")
                    .NumberFormat.Format = "0%";

                    ptSheet.Columns().AdjustToContents();
                }

                //Uncomment to replace expectation running.net6.0,
                //var expectedFileInVsSolution = Path.GetFullPath(Path.Combine("../../../", "Resource", expectedFilePath));
                //wb.SaveAs(expectedFileInVsSolution);

                return wb;
            }, expectedFilePath, ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void SourceSheetWithWhitespace()
        {
            using var wb = new XLWorkbook();

            var expectedFilePath = @"Other\PivotTableReferenceFiles\SourceSheetWithWhitespace\outputfile.xlsx";

            using var ms = new MemoryStream();
            TestHelper.CreateAndCompare(() =>
            {
                // Based on .\ClosedXML\ClosedXML.Examples\PivotTables\PivotTables.cs
                // But with empty column for Month
                var pastries = new List<Pastry>
                {
                        new Pastry("Croissant", 101, 150, 60.2, "", new DateTime(2016, 04, 21)),
                        new Pastry("Croissant", 101, 250, 50.42, "", new DateTime(2016, 05, 03)),
                        new Pastry("Croissant", 101, 134, 22.12, "", new DateTime(2016, 06, 24)),
                        new Pastry("Doughnut", 102, 250, 89.99, "", new DateTime(2017, 04, 23)),
                        new Pastry("Doughnut", 102, 225, 70, "", new DateTime(2016, 05, 24)),
                        new Pastry("Doughnut", 102, 210, 75.33, "", new DateTime(2016, 06, 02)),
                        new Pastry("Bearclaw", 103, 134, 10.24, "", new DateTime(2016, 04, 27)),
                        new Pastry("Bearclaw", 103, 184, 33.33, "", new DateTime(2016, 05, 20)),
                        new Pastry("Bearclaw", 103, 124, 25, "", new DateTime(2017, 06, 05)),
                        new Pastry("Danish", 104, 394, -20.24, "", null),
                        new Pastry("Danish", 104, 190, 60, "", new DateTime(2017, 05, 08)),
                        new Pastry("Danish", 104, 221, 24.76, "", new DateTime(2016, 06, 21)),

                        // Deliberately add different casings of same string to ensure pivot table doesn't duplicate it.
                        new Pastry("Scone", 105, 135, 0, "", new DateTime(2017, 04, 22)),
                        new Pastry("SconE", 105, 122, 5.19, "", new DateTime(2017, 05, 03)),
                        new Pastry("SCONE", 105, 243, 44.2, "", new DateTime(2017, 06, 14)),

                        // For ContainsBlank and integer rows/columns test
                        new Pastry("Scone", null, 255, 18.4, "", null),
                };

                var sheet = wb.Worksheets.Add("Pastry Sales Data");
                // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                var table = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
                sheet.Cell("F11").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                sheet.Columns().AdjustToContents();

                IXLWorksheet ptSheet;
                IXLPivotTable pt;

                // Add a new sheet for our pivot table
                ptSheet = wb.Worksheets.Add("pvt");

                // Create the pivot table, using the data from the "PastrySalesData" table
                pt = ptSheet.PivotTables.Add("pvt", ptSheet.Cell(1, 1), table.AsRange());
                pt.ColumnLabels.Add("Name");
                pt.RowLabels.Add("Month");

                // The values in our table will come from the "NumberOfOrders" field
                // The default calculation setting is a total of each row/column
                pt.Values.Add("NumberOfOrders", "NumberOfOrdersPercentageOfBearclaw")
                .ShowAsPercentageFrom("Name").And("Bearclaw")
                .NumberFormat.Format = "0%";

                ptSheet.Columns().AdjustToContents();

                // Uncomment to replace expectation running .net6.0,
                //var expectedFileInVsSolution = Path.GetFullPath(Path.Combine("../../../", "Resource", expectedFilePath));
                //wb.SaveAs(expectedFileInVsSolution);
                // Uncomment to replace expectation running .net4.8,
                //var expectedFileInVsSolution = Path.GetFullPath(Path.Combine("ClosedXML.Tests", "Resource", expectedFilePath));
                //wb.SaveAs(expectedFileInVsSolution);

                return wb;
            }, expectedFilePath, ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void PivotTableWithNoneTheme()
        {
            var expectedFilePath = @"Other\PivotTableReferenceFiles\PivotTableWithNoneTheme\outputfile.xlsx";

            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\PivotTableWithNoneTheme\inputfile.xlsx"));
            using var ms = new MemoryStream();

            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook(stream);

                wb.SaveAs(ms);

                // Uncomment to replace expectation running .net6.0,
                //var expectedFileInVsSolution = Path.GetFullPath(Path.Combine("../../../", "Resource", expectedFilePath));
                //wb.SaveAs(expectedFileInVsSolution);

                return wb;
            }, expectedFilePath);
        }

        [Test]
        public void MaintainPivotTableLabelsOrder()
        {
            var pastries = new List<Pastry>
            {
                new Pastry("Croissant", 101, 150, 60.2, "", new DateTime(2016, 04, 21)),
                new Pastry("Croissant", 101, 250, 50.42, "", new DateTime(2016, 05, 03)),
                new Pastry("Croissant", 101, 134, 22.12, "", new DateTime(2016, 06, 24)),
                new Pastry("Doughnut", 102, 250, 89.99, "", new DateTime(2017, 04, 23)),
                new Pastry("Doughnut", 102, 225, 70, "", new DateTime(2016, 05, 24)),
                new Pastry("Doughnut", 102, 210, 75.33, "", new DateTime(2016, 06, 02)),
                new Pastry("Bearclaw", 103, 134, 10.24, "", new DateTime(2016, 04, 27)),
                new Pastry("Bearclaw", 103, 184, 33.33, "", new DateTime(2016, 05, 20)),
                new Pastry("Bearclaw", 103, 124, 25, "", new DateTime(2017, 06, 05)),
                new Pastry("Danish", 104, 394, -20.24, "", null),
                new Pastry("Danish", 104, 190, 60, "", new DateTime(2017, 05, 08)),
                new Pastry("Danish", 104, 221, 24.76, "", new DateTime(2016, 06, 21)),

                // Deliberately add different casings of same string to ensure pivot table doesn't duplicate it.
                new Pastry("Scone", 105, 135, 0, "", new DateTime(2017, 04, 22)),
                new Pastry("SconE", 105, 122, 5.19, "", new DateTime(2017, 05, 03)),
                new Pastry("SCONE", 105, 243, 44.2, "", new DateTime(2017, 06, 14)),

                // For ContainsBlank and integer rows/columns test
                new Pastry("Scone", null, 255, 18.4, "", null),
            };

            using (var ms = new MemoryStream())
            {
                // Page fields
                using (var wb = new XLWorkbook())
                {
                    var sheet = wb.Worksheets.Add("PastrySalesData");
                    // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                    var table = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
                    sheet.Cell("F11").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet.Columns().AdjustToContents();

                    IXLWorksheet ptSheet;
                    IXLPivotTable pt;

                    // Add a new sheet for our pivot table
                    ptSheet = wb.Worksheets.Add("pvt");

                    // Create the pivot table, using the data from the "PastrySalesData" table
                    pt = ptSheet.PivotTables.Add("PastryPivot", ptSheet.Cell(1, 1), table);

                    pt.ReportFilters.Add("Month");
                    pt.ReportFilters.Add("Name");

                    pt.RowLabels.Add("BakeDate");
                    pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    var pageFields = wb.Worksheets.SelectMany(ws => ws.PivotTables)
                        .First()
                        .ReportFilters
                        .ToArray();

                    Assert.That(pageFields[0].SourceName, Is.EqualTo("Month"));
                    Assert.That(pageFields[1].SourceName, Is.EqualTo("Name"));
                }
            }

            using (var ms = new MemoryStream())
            {
                // Column labels
                using (var wb = new XLWorkbook())
                {
                    var sheet = wb.Worksheets.Add("PastrySalesData");
                    // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                    var table = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
                    sheet.Cell("F11").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet.Columns().AdjustToContents();

                    IXLWorksheet ptSheet;
                    IXLPivotTable pt;

                    // Add a new sheet for our pivot table
                    ptSheet = wb.Worksheets.Add("pvt");

                    // Create the pivot table, using the data from the "PastrySalesData" table
                    pt = ptSheet.PivotTables.Add("PastryPivot", ptSheet.Cell(1, 1), table);

                    pt.ColumnLabels.Add("Month");
                    pt.ColumnLabels.Add("Name");

                    pt.RowLabels.Add("BakeDate");
                    pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    var columnLabels = wb.Worksheets.SelectMany(ws => ws.PivotTables)
                        .First()
                        .ColumnLabels
                        .ToArray();

                    Assert.That(columnLabels[0].SourceName, Is.EqualTo("Month"));
                    Assert.That(columnLabels[1].SourceName, Is.EqualTo("Name"));
                }
            }

            using (var ms = new MemoryStream())
            {
                // Row labels
                using (var wb = new XLWorkbook())
                {
                    var sheet = wb.Worksheets.Add("PastrySalesData");
                    // Insert our list of pastry data into the "PastrySalesData" sheet at cell 1,1
                    var table = sheet.Cell(1, 1).InsertTable(pastries, "PastrySalesData", true);
                    sheet.Cell("F11").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet.Columns().AdjustToContents();

                    IXLWorksheet ptSheet;
                    IXLPivotTable pt;

                    // Add a new sheet for our pivot table
                    ptSheet = wb.Worksheets.Add("pvt");

                    // Create the pivot table, using the data from the "PastrySalesData" table
                    pt = ptSheet.PivotTables.Add("PastryPivot", ptSheet.Cell(1, 1), table);

                    pt.RowLabels.Add("Month");
                    pt.RowLabels.Add("Name");
                    pt.RowLabels.Add(XLConstants.PivotTable.ValuesSentinalLabel);

                    pt.ColumnLabels.Add("BakeDate");
                    pt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

                    wb.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    var rowLabels = wb.Worksheets.SelectMany(ws => ws.PivotTables)
                        .First()
                        .RowLabels
                        .ToArray();

                    Assert.That(rowLabels[0].SourceName, Is.EqualTo("Month"));
                    Assert.That(rowLabels[1].SourceName, Is.EqualTo("Name"));
                    Assert.That(rowLabels[2].SourceName, Is.EqualTo("{{Values}}"));
                }
            }
        }

        [Test]
        public void MaintainPivotTableIntegrityOnMultipleSaves()
        {
            var pastries = new List<Pastry>
            {
                new Pastry("Croissant", 101, 150, 60.2, "", new DateTime(2016, 04, 21)),
                new Pastry("Croissant", 101, 250, 50.42, "", new DateTime(2016, 05, 03)),
                new Pastry("Croissant", 101, 134, 22.12, "", new DateTime(2016, 06, 24)),
                new Pastry("Doughnut", 102, 250, 89.99, "", new DateTime(2017, 04, 23)),
                new Pastry("Doughnut", 102, 225, 70, "", new DateTime(2016, 05, 24)),
                new Pastry("Doughnut", 102, 210, 75.33, "", new DateTime(2016, 06, 02)),
                new Pastry("Bearclaw", 103, 134, 10.24, "", new DateTime(2016, 04, 27)),
                new Pastry("Bearclaw", 103, 184, 33.33, "", new DateTime(2016, 05, 20)),
                new Pastry("Bearclaw", 103, 124, 25, "", new DateTime(2017, 06, 05)),
                new Pastry("Danish", 104, 394, -20.24, "", null),
                new Pastry("Danish", 104, 190, 60, "", new DateTime(2017, 05, 08)),
                new Pastry("Danish", 104, 221, 24.76, "", new DateTime(2016, 06, 21)),

                // Deliberately add different casings of same string to ensure pivot table doesn't duplicate it.
                new Pastry("Scone", 105, 135, 0, "", new DateTime(2017, 04, 22)),
                new Pastry("SconE", 105, 122, 5.19, "", new DateTime(2017, 05, 03)),
                new Pastry("SCONE", 105, 243, 44.2, "", new DateTime(2017, 06, 14)),

                // For ContainsBlank and integer rows/columns test
                new Pastry("Scone", null, 255, 18.4, "", null),
            };

            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("PastrySalesData");
                var table = ws.FirstCell().InsertTable(pastries, "PastrySalesData", true);

                var pvtSheet = wb.Worksheets.Add("pvt");
                var pvt = table.CreatePivotTable(pvtSheet.FirstCell(), "PastryPvt");

                pvt.ColumnLabels.Add("Month");
                pvt.RowLabels.Add("Name");
                pvt.Values.Add("NumberOfOrders").SetSummaryFormula(XLPivotSummary.Sum);

                //Deliberately try to save twice
                wb.SaveAs(ms);
                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                Assert.That(wb.Worksheets.SelectMany(ws => ws.PivotTables).Count(), Is.EqualTo(1));
            }
        }

        [Test]
        public void TwoPivotWithOneSourceTest()
        {
            var expectedFilePath = @"Other\PivotTableReferenceFiles\TwoPivotTablesWithSingleSource\output.xlsx";

            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\TwoPivotTablesWithSingleSource\input.xlsx"));
            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook(stream);
                var srcRange = wb.Range("Sheet1!$B$2:$H$207");

                foreach (var pt in wb.Worksheets.SelectMany(ws => ws.PivotTables))
                {
                    pt.SourceRange = srcRange;
                }

                //Uncomment to replace expectation running.net6.0,
                //var expectedFileInVsSolution = Path.GetFullPath(Path.Combine("../../../", "Resource", expectedFilePath));
                //wb.SaveAs(expectedFileInVsSolution);

                return wb;
            }, expectedFilePath);
        }

        [Test]
        public void PivotSubtotalsLoadingTest()
        {
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\PivotSubtotalsSource\input.xlsx"));
            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook(stream);
                return wb;
            }, @"Other\PivotTableReferenceFiles\PivotSubtotalsSource\input.xlsx");
        }

        [Test]
        public void ClearPivotTableTenderedTange()
        {
            // https://github.com/ClosedXML/ClosedXML/pull/856
            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\ClearPivotTableRenderedRangeWhenLoading\inputfile.xlsx"));
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook(stream))
            {
                var ws = wb.Worksheet("Sheet1");
                Assert.That(ws.Cell("B1").IsEmpty(), Is.True);
                Assert.That(ws.Cell("C2").IsEmpty(), Is.True);
                Assert.That(ws.Cell("D5").IsEmpty(), Is.True);
                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheet("Sheet1");
                Assert.That(ws.Cell("B1").IsEmpty(), Is.True);
                Assert.That(ws.Cell("C2").IsEmpty(), Is.True);
                Assert.That(ws.Cell("D5").IsEmpty(), Is.True);
            }
        }

        private static void SetFieldOptions(IXLPivotField field, bool withDefaults)
        {
            field.SubtotalsAtTop = !withDefaults;
            field.ShowBlankItems = !withDefaults;
            field.Outline = !withDefaults;
            field.Compact = !withDefaults;
            field.Collapsed = withDefaults;
            field.InsertBlankLines = withDefaults;
            field.RepeatItemLabels = withDefaults;
            field.InsertPageBreaks = withDefaults;
            field.IncludeNewItemsInFilter = withDefaults;
        }

        private static void AssertFieldOptions(IXLPivotField field, bool withDefaults)
        {
            Assert.That(field.SubtotalsAtTop, Is.EqualTo(!withDefaults), "SubtotalsAtTop save failure");
            Assert.That(field.ShowBlankItems, Is.EqualTo(!withDefaults), "ShowBlankItems save failure");
            Assert.That(field.Outline, Is.EqualTo(!withDefaults), "Outline save failure");
            Assert.That(field.Compact, Is.EqualTo(!withDefaults), "Compact save failure");
            Assert.That(field.Collapsed, Is.EqualTo(withDefaults), "Collapsed save failure");
            Assert.That(field.InsertBlankLines, Is.EqualTo(withDefaults), "InsertBlankLines save failure");
            Assert.That(field.RepeatItemLabels, Is.EqualTo(withDefaults), "RepeatItemLabels save failure");
            Assert.That(field.InsertPageBreaks, Is.EqualTo(withDefaults), "InsertPageBreaks save failure");
            Assert.That(field.IncludeNewItemsInFilter, Is.EqualTo(withDefaults), "IncludeNewItemsInFilter save failure");
        }
    }
}