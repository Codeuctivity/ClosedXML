
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.NamedRanges
{
    [TestFixture]
    public class NamedRangesTests
    {
        [Test]
        public void CanEvaluateNamedMultiRange()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet1");
            ws1.Range("A1:C1").Value = 1;
            ws1.Range("A3:C3").Value = 3;
            wb.NamedRanges.Add("TEST", ws1.Ranges("A1:C1,A3:C3"));

            ws1.Cell(2, 1).FormulaA1 = "=SUM(TEST)";

            Assert.That((double)ws1.Cell(2, 1).Value, Is.EqualTo(12.0).Within(XLHelper.Epsilon));
        }

        [Test]
        public void CanGetNamedFromAnother()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            ws1.Cell("A1").SetValue(1).AddToNamed("value1");

            Assert.That(wb.Cell("value1").GetValue<int>(), Is.EqualTo(1));
            Assert.That(wb.Range("value1").FirstCell().GetValue<int>(), Is.EqualTo(1));

            Assert.That(ws1.Cell("value1").GetValue<int>(), Is.EqualTo(1));
            Assert.That(ws1.Range("value1").FirstCell().GetValue<int>(), Is.EqualTo(1));

            var ws2 = wb.Worksheets.Add("Sheet2");

            ws2.Cell("A1").SetFormulaA1("=value1").AddToNamed("value2");

            Assert.That(wb.Cell("value2").GetValue<int>(), Is.EqualTo(1));
            Assert.That(wb.Range("value2").FirstCell().GetValue<int>(), Is.EqualTo(1));

            Assert.That(ws2.Cell("value1").GetValue<int>(), Is.EqualTo(1));
            Assert.That(ws2.Range("value1").FirstCell().GetValue<int>(), Is.EqualTo(1));

            Assert.That(ws2.Cell("value2").GetValue<int>(), Is.EqualTo(1));
            Assert.That(ws2.Range("value2").FirstCell().GetValue<int>(), Is.EqualTo(1));
        }

        [Test]
        public void CanGetValidNamedRanges()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet 1");
            var ws2 = wb.Worksheets.Add("Sheet 2");
            var ws3 = wb.Worksheets.Add("Sheet'3");

            ws1.Range("A1:D1").AddToNamed("Named range 1", XLScope.Worksheet);
            ws1.Range("A2:D2").AddToNamed("Named range 2", XLScope.Workbook);
            ws2.Range("A3:D3").AddToNamed("Named range 3", XLScope.Worksheet);
            ws2.Range("A4:D4").AddToNamed("Named range 4", XLScope.Workbook);
            wb.NamedRanges.Add("Named range 5", new XLRanges
                {
                    ws1.Range("A5:D5"),
                    ws3.Range("A5:D5")
                });

            ws2.Delete();
            ws3.Delete();

            var globalValidRanges = wb.NamedRanges.ValidNamedRanges();
            var globalInvalidRanges = wb.NamedRanges.InvalidNamedRanges();
            var localValidRanges = ws1.NamedRanges.ValidNamedRanges();
            var localInvalidRanges = ws1.NamedRanges.InvalidNamedRanges();

            Assert.That(globalValidRanges.Count(), Is.EqualTo(1));
            Assert.That(globalValidRanges.First().Name, Is.EqualTo("Named range 2"));

            Assert.That(globalInvalidRanges.Count(), Is.EqualTo(2));
            Assert.That(globalInvalidRanges.First().Name, Is.EqualTo("Named range 4"));
            Assert.That(globalInvalidRanges.Last().Name, Is.EqualTo("Named range 5"));

            Assert.That(localValidRanges.Count(), Is.EqualTo(1));
            Assert.That(localValidRanges.First().Name, Is.EqualTo("Named range 1"));

            Assert.That(localInvalidRanges.Count(), Is.EqualTo(0));
        }

        [Test]
        public void CanRenameNamedRange()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet1");
            var nr1 = wb.NamedRanges.Add("TEST", "=0.1");

            Assert.That(wb.NamedRanges.TryGetValue("TEST", out var _), Is.True);
            Assert.That(wb.NamedRanges.TryGetValue("TEST1", out var _), Is.False);

            nr1.Name = "TEST1";

            Assert.That(wb.NamedRanges.TryGetValue("TEST", out var _), Is.False);
            Assert.That(wb.NamedRanges.TryGetValue("TEST1", out var _), Is.True);

            var nr2 = wb.NamedRanges.Add("TEST2", "=TEST1*2");

            ws1.Cell(1, 1).FormulaA1 = "TEST1";
            ws1.Cell(2, 1).FormulaA1 = "TEST1*10";
            ws1.Cell(3, 1).FormulaA1 = "TEST2";
            ws1.Cell(4, 1).FormulaA1 = "TEST2*3";

            Assert.That((double)ws1.Cell(1, 1).Value, Is.EqualTo(0.1).Within(XLHelper.Epsilon));
            Assert.That((double)ws1.Cell(2, 1).Value, Is.EqualTo(1.0).Within(XLHelper.Epsilon));
            Assert.That((double)ws1.Cell(3, 1).Value, Is.EqualTo(0.2).Within(XLHelper.Epsilon));
            Assert.That((double)ws1.Cell(4, 1).Value, Is.EqualTo(0.6).Within(XLHelper.Epsilon));
        }

        [Test]
        public void CanSaveAndLoadNamedRanges()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var sheet1 = wb.Worksheets.Add("Sheet1");
                var sheet2 = wb.Worksheets.Add("Sheet2");

                wb.NamedRanges.Add("wbNamedRange",
                    "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
                sheet1.NamedRanges.Add("sheet1NamedRange",
                    "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
                sheet2.NamedRanges.Add("sheet2NamedRange", "Sheet1!A1,Sheet2!A1");

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var sheet1 = wb.Worksheet("Sheet1");
                var sheet2 = wb.Worksheet("Sheet2");

                Assert.That(wb.NamedRanges.Count(), Is.EqualTo(1));
                Assert.That(wb.NamedRanges.Single().Name, Is.EqualTo("wbNamedRange"));
                Assert.That(wb.NamedRanges.Single().RefersTo, Is.EqualTo("Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G"));
                Assert.That(wb.NamedRanges.Single().Ranges.Count, Is.EqualTo(5));

                Assert.That(sheet1.NamedRanges.Count(), Is.EqualTo(1));
                Assert.That(sheet1.NamedRanges.Single().Name, Is.EqualTo("sheet1NamedRange"));
                Assert.That(sheet1.NamedRanges.Single().RefersTo, Is.EqualTo("Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G"));
                Assert.That(sheet1.NamedRanges.Single().Ranges.Count, Is.EqualTo(5));

                Assert.That(sheet2.NamedRanges.Count(), Is.EqualTo(1));
                Assert.That(sheet2.NamedRanges.Single().Name, Is.EqualTo("sheet2NamedRange"));
                Assert.That(sheet2.NamedRanges.Single().RefersTo, Is.EqualTo("Sheet1!A1,Sheet2!A1"));
                Assert.That(sheet2.NamedRanges.Single().Ranges.Count, Is.EqualTo(2));
            }
        }

        [Test]
        public void CopyNamedRangeDifferentWorksheets()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            var ws2 = wb.Worksheets.Add("Sheet2");
            var ranges = new XLRanges
            {
                ws1.Range("B2:E6"),
                ws2.Range("D1:E2")
            };
            var original = ws1.NamedRanges.Add("Named range", ranges);

            var copy = original.CopyTo(ws2);

            Assert.That(ws1.NamedRanges.Count(), Is.EqualTo(1));
            Assert.That(ws2.NamedRanges.Count(), Is.EqualTo(1));
            Assert.That(original.Ranges.Count, Is.EqualTo(2));
            Assert.That(copy.Ranges.Count, Is.EqualTo(2));
            Assert.That(copy.Name, Is.EqualTo(original.Name));
            Assert.That(copy.Scope, Is.EqualTo(original.Scope));
            Assert.That(original.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("Sheet1!B2:E6"));
            Assert.That(original.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("Sheet2!D1:E2"));
            Assert.That(copy.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("Sheet2!D1:E2"));
            Assert.That(copy.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("Sheet2!B2:E6"));
        }

        [Test]
        public void CopyNamedRangeSameWorksheet()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet1");
            ws1.Range("B2:E6").AddToNamed("Named range", XLScope.Worksheet);
            var nr = ws1.NamedRange("Named range");

            void action() => nr.CopyTo(ws1);

            Assert.Throws(typeof(InvalidOperationException), action);
        }

        [Test]
        public void DeleteColumnUsedInNamedRange()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Column1");
            ws.FirstCell().CellRight().SetValue("Column2").Style.Font.SetBold();
            ws.FirstCell().CellRight(2).SetValue("Column3");
            ws.NamedRanges.Add("MyRange", "A1:C1");

            ws.Column(1).Delete();

            Assert.That(ws.Cell("A1").Style.Font.Bold, Is.True);
            Assert.That(ws.Cell("B1").GetValue<string>(), Is.EqualTo("Column3"));
            Assert.That(ws.Cell("C1").GetValue<string>(), Is.Empty);
        }

        [Test]
        public void MovingRanges()
        {
            using var wb = new XLWorkbook();

            var sheet1 = wb.Worksheets.Add("Sheet1");
            var sheet2 = wb.Worksheets.Add("Sheet2");

            wb.NamedRanges.Add("wbNamedRange",
                "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
            sheet1.NamedRanges.Add("sheet1NamedRange",
                "Sheet1!$B$2,Sheet1!$B$3:$C$3,Sheet2!$D$3:$D$4,Sheet1!$6:$7,Sheet1!$F:$G");
            sheet2.NamedRanges.Add("sheet2NamedRange", "Sheet1!A1,Sheet2!A1");

            sheet1.Row(1).InsertRowsAbove(2);
            sheet1.Row(1).Delete();
            sheet1.Column(1).InsertColumnsBefore(2);
            sheet1.Column(1).Delete();

            Assert.That(wb.NamedRanges.First().RefersTo, Is.EqualTo("Sheet1!$C$3,Sheet1!$C$4:$D$4,Sheet2!$D$3:$D$4,Sheet1!$7:$8,Sheet1!$G:$H"));
            Assert.That(sheet1.NamedRanges.First().RefersTo, Is.EqualTo("Sheet1!$C$3,Sheet1!$C$4:$D$4,Sheet2!$D$3:$D$4,Sheet1!$7:$8,Sheet1!$G:$H"));
            Assert.That(sheet2.NamedRanges.First().RefersTo, Is.EqualTo("Sheet1!B2,Sheet2!A1"));

            wb.NamedRanges.ForEach(nr => Assert.That(nr.Scope, Is.EqualTo(XLNamedRangeScope.Workbook)));
            sheet1.NamedRanges.ForEach(nr => Assert.That(nr.Scope, Is.EqualTo(XLNamedRangeScope.Worksheet)));
            sheet2.NamedRanges.ForEach(nr => Assert.That(nr.Scope, Is.EqualTo(XLNamedRangeScope.Worksheet)));
        }

        [Test, Ignore("Muted until shifting is fixed (see #880)")]
        public void NamedRangeBecomesInvalidOnRangeAndWorksheetDeleting()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet 1");
            var ws2 = wb.Worksheets.Add("Sheet 2");
            ws1.Range("A1:B2").AddToNamed("Simple", XLScope.Workbook);
            wb.NamedRanges.Add("Compound", new XLRanges
                {
                    ws1.Range("C1:D2"),
                    ws2.Range("A10:D15")
                });

            ws1.Rows(1, 5).Delete();
            ws1.Delete();

            Assert.That(wb.NamedRanges.Count(), Is.EqualTo(2));
            Assert.That(wb.NamedRanges.ValidNamedRanges().Count(), Is.EqualTo(0));
            Assert.That(wb.NamedRanges.ElementAt(0).RefersTo, Is.EqualTo("#REF!#REF!"));
            Assert.That(wb.NamedRanges.ElementAt(0).RefersTo, Is.EqualTo("#REF!#REF!,'Sheet 2'!A10:D15"));
        }

        [Test, Ignore("Muted until shifting is fixed (see #880)")]
        public void NamedRangeBecomesInvalidOnRangeDeleting()
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Sheet 1");
            ws.Range("A1:B2").AddToNamed("Simple", XLScope.Workbook);
            wb.NamedRanges.Add("Compound", new XLRanges
                {
                    ws.Range("C1:D2"),
                    ws.Range("A10:D15")
                });

            ws.Rows(1, 5).Delete();

            Assert.That(wb.NamedRanges.Count(), Is.EqualTo(2));
            Assert.That(wb.NamedRanges.ValidNamedRanges().Count(), Is.EqualTo(0));
            Assert.That(wb.NamedRanges.ElementAt(0).RefersTo, Is.EqualTo("'Sheet 1'!#REF!"));
            Assert.That(wb.NamedRanges.ElementAt(0).RefersTo, Is.EqualTo("'Sheet 1'!#REF!,'Sheet 1'!A5:D10"));
        }

        [Test]
        public void NamedRangeMayReferToExpression()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet1");
                wb.NamedRanges.Add("TEST", "=0.1");
                wb.NamedRanges.Add("TEST2", "=TEST*2");

                ws1.Cell(1, 1).FormulaA1 = "TEST";
                ws1.Cell(2, 1).FormulaA1 = "TEST*10";
                ws1.Cell(3, 1).FormulaA1 = "TEST2";
                ws1.Cell(4, 1).FormulaA1 = "TEST2*3";

                Assert.That((double)ws1.Cell(1, 1).Value, Is.EqualTo(0.1).Within(XLHelper.Epsilon));
                Assert.That((double)ws1.Cell(2, 1).Value, Is.EqualTo(1.0).Within(XLHelper.Epsilon));
                Assert.That((double)ws1.Cell(3, 1).Value, Is.EqualTo(0.2).Within(XLHelper.Epsilon));
                Assert.That((double)ws1.Cell(4, 1).Value, Is.EqualTo(0.6).Within(XLHelper.Epsilon));

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var ws1 = wb.Worksheets.First();

                Assert.That((double)ws1.Cell(1, 1).Value, Is.EqualTo(0.1).Within(XLHelper.Epsilon));
                Assert.That((double)ws1.Cell(2, 1).Value, Is.EqualTo(1.0).Within(XLHelper.Epsilon));
                Assert.That((double)ws1.Cell(3, 1).Value, Is.EqualTo(0.2).Within(XLHelper.Epsilon));
                Assert.That((double)ws1.Cell(4, 1).Value, Is.EqualTo(0.6).Within(XLHelper.Epsilon));
            }
        }

        [Test]
        public void NamedRangeReferringToMultipleRangesCanBeSavedAndLoaded()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet 1");

                wb.NamedRanges.Add("Multirange named range", new XLRanges
                    {
                        ws.Range("A5:D5"),
                        ws.Range("A15:D15")
                    });

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                Assert.That(wb.NamedRanges.Count(), Is.EqualTo(1));
                var nr = wb.NamedRanges.Single() as XLNamedRange;
                Assert.That(nr.RefersTo, Is.EqualTo("'Sheet 1'!$A$5:$D$5,'Sheet 1'!$A$15:$D$15"));
                Assert.That(nr.Ranges.Count, Is.EqualTo(2));
                Assert.That(nr.Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!A5:D5"));
                Assert.That(nr.Ranges.Last().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!A15:D15"));
                Assert.That(nr.RangeList.Count, Is.EqualTo(2));
                Assert.That(nr.RangeList.First(), Is.EqualTo("'Sheet 1'!$A$5:$D$5"));
                Assert.That(nr.RangeList.Last(), Is.EqualTo("'Sheet 1'!$A$15:$D$15"));
            }
        }

        [Test]
        public void NamedRangesBecomeInvalidOnWorksheetDeleting()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("Sheet 1");
            var ws2 = wb.Worksheets.Add("Sheet 2");
            var ws3 = wb.Worksheets.Add("Sheet'3");

            ws1.Range("A1:D1").AddToNamed("Named range 1", XLScope.Worksheet);
            ws1.Range("A2:D2").AddToNamed("Named range 2", XLScope.Workbook);
            ws2.Range("A3:D3").AddToNamed("Named range 3", XLScope.Worksheet);
            ws2.Range("A4:D4").AddToNamed("Named range 4", XLScope.Workbook);
            wb.NamedRanges.Add("Named range 5", new XLRanges
                {
                    ws1.Range("A5:D5"),
                    ws3.Range("A5:D5")
                });

            ws2.Delete();
            ws3.Delete();

            Assert.That(ws1.NamedRanges.Count(), Is.EqualTo(1));
            Assert.That(ws1.NamedRanges.First().Name, Is.EqualTo("Named range 1"));
            Assert.That(ws1.NamedRanges.First().Scope, Is.EqualTo(XLNamedRangeScope.Worksheet));
            Assert.That(ws1.NamedRanges.First().RefersTo, Is.EqualTo("'Sheet 1'!$A$1:$D$1"));
            Assert.That(ws1.NamedRanges.First().Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!A1:D1"));

            Assert.That(wb.NamedRanges.Count(), Is.EqualTo(3));

            Assert.That(wb.NamedRanges.ElementAt(0).Name, Is.EqualTo("Named range 2"));
            Assert.That(wb.NamedRanges.ElementAt(0).Scope, Is.EqualTo(XLNamedRangeScope.Workbook));
            Assert.That(wb.NamedRanges.ElementAt(0).RefersTo, Is.EqualTo("'Sheet 1'!$A$2:$D$2"));
            Assert.That(wb.NamedRanges.ElementAt(0).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!A2:D2"));

            Assert.That(wb.NamedRanges.ElementAt(1).Name, Is.EqualTo("Named range 4"));
            Assert.That(wb.NamedRanges.ElementAt(1).Scope, Is.EqualTo(XLNamedRangeScope.Workbook));
            Assert.That(wb.NamedRanges.ElementAt(1).RefersTo, Is.EqualTo("#REF!$A$4:$D$4"));
            Assert.That(wb.NamedRanges.ElementAt(1).Ranges.Any(), Is.False);

            Assert.That(wb.NamedRanges.ElementAt(2).Name, Is.EqualTo("Named range 5"));
            Assert.That(wb.NamedRanges.ElementAt(2).Scope, Is.EqualTo(XLNamedRangeScope.Workbook));
            Assert.That(wb.NamedRanges.ElementAt(2).RefersTo, Is.EqualTo("'Sheet 1'!$A$5:$D$5,#REF!$A$5:$D$5"));
            Assert.That(wb.NamedRanges.ElementAt(2).Ranges.Count, Is.EqualTo(1));
            Assert.That(wb.NamedRanges.ElementAt(2).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!A5:D5"));
        }

        [Test]
        public void NamedRangesFromDeletedSheetAreSavedWithoutAddress()
        {
            // Range address referring to the deleted sheet look like #REF!A1:B2.
            // But workbooks with such references in named ranges Excel considers as broken files.
            // It requires #REF!

            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                wb.Worksheets.Add("Sheet 1");
                var ws2 = wb.Worksheets.Add("Sheet 2");
                ws2.Range("A4:D4").AddToNamed("Test named range", XLScope.Workbook);
                ws2.Delete();
                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                Assert.That(wb.NamedRanges.Single().RefersTo, Is.EqualTo("#REF!"));
            }
        }

        [Test]
        public void NamedRangesWhenCopyingWorksheets()
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet("Sheet1");
            ws1.FirstCell().Value = Enumerable.Range(1, 10);
            wb.NamedRanges.Add("wbNamedRange", ws1.Range("A1:A10"));
            ws1.NamedRanges.Add("wsNamedRange", ws1.Range("A3"));

            var ws2 = wb.AddWorksheet("Sheet2");
            ws2.FirstCell().Value = Enumerable.Range(101, 10);
            ws1.NamedRanges.Add("wsNamedRangeAcrossSheets", ws2.Range("A4"));

            ws1.Cell("C1").FormulaA1 = "=wbNamedRange";
            ws1.Cell("C2").FormulaA1 = "=wsNamedRange";
            ws1.Cell("C3").FormulaA1 = "=wsNamedRangeAcrossSheets";

            Assert.That(ws1.Cell("C1").Value, Is.EqualTo(1));
            Assert.That(ws1.Cell("C2").Value, Is.EqualTo(3));
            Assert.That(ws1.Cell("C3").Value, Is.EqualTo(104));

            var wsCopy = ws1.CopyTo("Copy");
            Assert.That(wsCopy.Cell("C1").Value, Is.EqualTo(1));
            Assert.That(wsCopy.Cell("C2").Value, Is.EqualTo(3));
            Assert.That(wsCopy.Cell("C3").Value, Is.EqualTo(104));

            Assert.That(wb.NamedRange("wbNamedRange").Ranges.First().RangeAddress.ToStringRelative(true), Is.EqualTo("Sheet1!A1:A10"));
            Assert.That(wsCopy.NamedRange("wsNamedRange").Ranges.First().RangeAddress.ToStringRelative(true), Is.EqualTo("Copy!A3:A3"));
            Assert.That(wsCopy.NamedRange("wsNamedRangeAcrossSheets").Ranges.First().RangeAddress.ToStringRelative(true), Is.EqualTo("Sheet2!A4:A4"));
        }

        [Test]
        public void SavedNamedRangesBecomeInvalidOnWorksheetDeleting()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.Worksheets.Add("Sheet 1");
                var ws2 = wb.Worksheets.Add("Sheet2");
                var ws3 = wb.Worksheets.Add("Sheet'3");

                ws1.Range("A1:D1").AddToNamed("Named range 1", XLScope.Worksheet);
                ws1.Range("A2:D2").AddToNamed("Named range 2", XLScope.Workbook);
                ws2.Range("A3:D3").AddToNamed("Named range 3", XLScope.Worksheet);
                ws2.Range("A4:D4").AddToNamed("Named range 4", XLScope.Workbook);
                wb.NamedRanges.Add("Named range 5", new XLRanges
                    {
                        ws1.Range("A5:D5"),
                        ws3.Range("A5:D5")
                    });

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                wb.Worksheet("Sheet2").Delete();
                wb.Worksheet("Sheet'3").Delete();
                wb.Save();
            }

            using (var wb = new XLWorkbook(ms))
            {
                var ws1 = wb.Worksheet("Sheet 1");
                Assert.That(ws1.NamedRanges.Count(), Is.EqualTo(1));
                Assert.That(ws1.NamedRanges.First().Name, Is.EqualTo("Named range 1"));
                Assert.That(ws1.NamedRanges.First().Scope, Is.EqualTo(XLNamedRangeScope.Worksheet));
                Assert.That(ws1.NamedRanges.First().RefersTo, Is.EqualTo("'Sheet 1'!$A$1:$D$1"));
                Assert.That(ws1.NamedRanges.First().Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!A1:D1"));

                Assert.That(wb.NamedRanges.Count(), Is.EqualTo(3));

                Assert.That(wb.NamedRanges.ElementAt(0).Name, Is.EqualTo("Named range 2"));
                Assert.That(wb.NamedRanges.ElementAt(0).Scope, Is.EqualTo(XLNamedRangeScope.Workbook));
                Assert.That(wb.NamedRanges.ElementAt(0).RefersTo, Is.EqualTo("'Sheet 1'!$A$2:$D$2"));
                Assert.That(wb.NamedRanges.ElementAt(0).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!A2:D2"));

                Assert.That(wb.NamedRanges.ElementAt(1).Name, Is.EqualTo("Named range 4"));
                Assert.That(wb.NamedRanges.ElementAt(1).Scope, Is.EqualTo(XLNamedRangeScope.Workbook));
                Assert.That(wb.NamedRanges.ElementAt(1).RefersTo, Is.EqualTo("#REF!"));
                Assert.That(wb.NamedRanges.ElementAt(1).Ranges.Any(), Is.False);

                Assert.That(wb.NamedRanges.ElementAt(2).Name, Is.EqualTo("Named range 5"));
                Assert.That(wb.NamedRanges.ElementAt(2).Scope, Is.EqualTo(XLNamedRangeScope.Workbook));
                Assert.That(wb.NamedRanges.ElementAt(2).RefersTo, Is.EqualTo("'Sheet 1'!$A$5:$D$5,#REF!"));
                Assert.That(wb.NamedRanges.ElementAt(2).Ranges.Count, Is.EqualTo(1));
                Assert.That(wb.NamedRanges.ElementAt(2).Ranges.Single().RangeAddress.ToString(XLReferenceStyle.A1, true), Is.EqualTo("'Sheet 1'!A5:D5"));
            }
        }

        [Test]
        public void TestInvalidNamedRangeOnWorkbookScope()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().SetValue("Column1");
            ws.FirstCell().CellRight().SetValue("Column2").Style.Font.SetBold();
            ws.FirstCell().CellRight(2).SetValue("Column3");

            Assert.Throws<ArgumentException>(() => wb.NamedRanges.Add("MyRange", "A1:C1"));
        }

        [Test]
        public void WbContainsWsNamedRange()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddToNamed("Name", XLScope.Worksheet);

            Assert.That(wb.NamedRanges.Contains("Sheet1!Name"), Is.True);
            Assert.That(wb.NamedRanges.Contains("Sheet1!NameX"), Is.False);

            Assert.That(wb.NamedRange("Sheet1!Name"), Is.Not.Null);
            Assert.That(wb.NamedRange("Sheet1!NameX"), Is.Null);

            var result1 = wb.NamedRanges.TryGetValue("Sheet1!Name", out var range1);
            Assert.That(result1, Is.True);
            Assert.That(range1, Is.Not.Null);
            Assert.That(range1.Scope, Is.EqualTo(XLNamedRangeScope.Worksheet));

            var result2 = wb.NamedRanges.TryGetValue("Sheet1!NameX", out var range2);
            Assert.That(result2, Is.False);
            Assert.That(range2, Is.Null);
        }

        [Test]
        public void WorkbookContainsNamedRange()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddToNamed("Name");

            Assert.That(wb.NamedRanges.Contains("Name"), Is.True);
            Assert.That(wb.NamedRanges.Contains("NameX"), Is.False);

            Assert.That(wb.NamedRange("Name"), Is.Not.Null);
            Assert.That(wb.NamedRange("NameX"), Is.Null);

            var result1 = wb.NamedRanges.TryGetValue("Name", out var range1);
            Assert.That(result1, Is.True);
            Assert.That(range1, Is.Not.Null);

            var result2 = wb.NamedRanges.TryGetValue("NameX", out var range2);
            Assert.That(result2, Is.False);
            Assert.That(range2, Is.Null);
        }

        [Test]
        public void WorksheetContainsNamedRange()
        {
            using var xLWorkbook = new XLWorkbook(); var ws = xLWorkbook.AddWorksheet("Sheet1");
            ws.FirstCell().AddToNamed("Name", XLScope.Worksheet);

            Assert.That(ws.NamedRanges.Contains("Name"), Is.True);
            Assert.That(ws.NamedRanges.Contains("NameX"), Is.False);

            Assert.That(ws.NamedRange("Name"), Is.Not.Null);
            Assert.That(ws.NamedRange("NameX"), Is.Null);

            var result1 = ws.NamedRanges.TryGetValue("Name", out var range1);
            Assert.That(result1, Is.True);
            Assert.That(range1, Is.Not.Null);

            var result2 = ws.NamedRanges.TryGetValue("NameX", out var range2);
            Assert.That(result2, Is.False);
            Assert.That(range2, Is.Null);
        }

        [Test]
        public void NamedRangeWithSameNameAsAFunction()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var a1 = ws.FirstCell();
            var a2 = a1.CellBelow();

            a1.SetValue(5).AddToNamed("RAND");
            a2.FormulaA1 = "=RAND * 10";

            Assert.That(a2.GetDouble(), Is.EqualTo(50));
        }
    }
}