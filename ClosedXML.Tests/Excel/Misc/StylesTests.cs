using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Misc
{
    [TestFixture]
    public class StylesTests
    {
        private static void SetupBorders(IXLRange range)
        {
            range.FirstRow().Cell(1).Style.Border.TopBorder = XLBorderStyleValues.None;
            range.FirstRow().Cell(2).Style.Border.TopBorder = XLBorderStyleValues.Thick;
            range.FirstRow().Cell(3).Style.Border.TopBorder = XLBorderStyleValues.Double;

            range.LastRow().Cell(1).Style.Border.BottomBorder = XLBorderStyleValues.None;
            range.LastRow().Cell(2).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
            range.LastRow().Cell(3).Style.Border.BottomBorder = XLBorderStyleValues.Double;

            range.FirstColumn().Cell(1).Style.Border.LeftBorder = XLBorderStyleValues.None;
            range.FirstColumn().Cell(2).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
            range.FirstColumn().Cell(3).Style.Border.LeftBorder = XLBorderStyleValues.Double;

            range.LastColumn().Cell(1).Style.Border.RightBorder = XLBorderStyleValues.None;
            range.LastColumn().Cell(2).Style.Border.RightBorder = XLBorderStyleValues.Thick;
            range.LastColumn().Cell(3).Style.Border.RightBorder = XLBorderStyleValues.Double;
        }

        [Test]
        public void InsideBorderTest()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var range = ws.Range("B2:D4");

            SetupBorders(range);

            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorderColor = XLColor.Red;

            var center = range.Cell(2, 2);

            Assert.That(center.Style.Border.TopBorderColor, Is.EqualTo(XLColor.Red));
            Assert.That(center.Style.Border.BottomBorderColor, Is.EqualTo(XLColor.Red));
            Assert.That(center.Style.Border.LeftBorderColor, Is.EqualTo(XLColor.Red));
            Assert.That(center.Style.Border.RightBorderColor, Is.EqualTo(XLColor.Red));

            Assert.That(range.FirstRow().Cell(1).Style.Border.TopBorder, Is.EqualTo(XLBorderStyleValues.None));
            Assert.That(range.FirstRow().Cell(2).Style.Border.TopBorder, Is.EqualTo(XLBorderStyleValues.Thick));
            Assert.That(range.FirstRow().Cell(3).Style.Border.TopBorder, Is.EqualTo(XLBorderStyleValues.Double));

            Assert.That(range.LastRow().Cell(1).Style.Border.BottomBorder, Is.EqualTo(XLBorderStyleValues.None));
            Assert.That(range.LastRow().Cell(2).Style.Border.BottomBorder, Is.EqualTo(XLBorderStyleValues.Thick));
            Assert.That(range.LastRow().Cell(3).Style.Border.BottomBorder, Is.EqualTo(XLBorderStyleValues.Double));

            Assert.That(range.FirstColumn().Cell(1).Style.Border.LeftBorder, Is.EqualTo(XLBorderStyleValues.None));
            Assert.That(range.FirstColumn().Cell(2).Style.Border.LeftBorder, Is.EqualTo(XLBorderStyleValues.Thick));
            Assert.That(range.FirstColumn().Cell(3).Style.Border.LeftBorder, Is.EqualTo(XLBorderStyleValues.Double));

            Assert.That(range.LastColumn().Cell(1).Style.Border.RightBorder, Is.EqualTo(XLBorderStyleValues.None));
            Assert.That(range.LastColumn().Cell(2).Style.Border.RightBorder, Is.EqualTo(XLBorderStyleValues.Thick));
            Assert.That(range.LastColumn().Cell(3).Style.Border.RightBorder, Is.EqualTo(XLBorderStyleValues.Double));
        }

        [Test]
        public void ResolveThemeColors()
        {
            using var wb = new XLWorkbook();
            string color;
            color = wb.Theme.ResolveThemeColor(XLThemeColor.Accent1).Color.ToHex();
            Assert.That(color, Is.EqualTo("FF4F81BD"));

            color = wb.Theme.ResolveThemeColor(XLThemeColor.Background1).Color.ToHex();
            Assert.That(color, Is.EqualTo("FFFFFFFF"));
        }

        [Theory]
        public void CanResolveAllThemeColors(XLThemeColor themeColor)
        {
            using var xLWorkbook = new XLWorkbook();
            var theme = xLWorkbook.Theme;
            var color = theme.ResolveThemeColor(themeColor);
            Assert.That(color, Is.Not.Null);
        }

        [Test]
        public void SetStyleViaRowReference()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Style
               .Font.SetFontSize(8)
               .Font.SetFontColor(XLColor.Green)
               .Font.SetBold(true);

            var row = ws.Row(1);
            ws.Cell(1, 1).Value = "Test";
            row.Cell(2).Value = "Test";
            row.Cells(3, 3).Value = "Test";

            foreach (var cell in ws.CellsUsed())
            {
                Assert.That(ws.Cell("A1").Style.Font.FontSize, Is.EqualTo(8));
                Assert.That(ws.Cell("B1").Style.Font.FontColor, Is.EqualTo(XLColor.Green));
                Assert.That(ws.Cell("C1").Style.Font.Bold, Is.EqualTo(true));
            }
        }
    }
}