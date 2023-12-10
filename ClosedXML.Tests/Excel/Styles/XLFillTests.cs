using ClosedXML.Excel;
using NUnit.Framework;
using System.IO;

namespace ClosedXML.Tests.Excel.Styles
{
    [TestFixture]
    public class XLFillTests
    {
        [Test]
        public void BackgroundColorSetsPattern()
        {
            var fill = new XLFill { BackgroundColor = XLColor.Blue };
            Assert.That(fill.PatternType, Is.EqualTo(XLFillPatternValues.Solid));
        }

        [Test]
        public void BackgroundNoColorSetsPatternNone()
        {
            var fill = new XLFill { BackgroundColor = XLColor.NoColor };
            Assert.That(fill.PatternType, Is.EqualTo(XLFillPatternValues.None));
        }

        [Test]
        public void BackgroundPatternEqualCheck()
        {
            var fill1 = new XLFill { BackgroundColor = XLColor.Blue };
            var fill2 = new XLFill { BackgroundColor = XLColor.Blue };
            Assert.That(fill1.Equals(fill2), Is.True);
            Assert.That(fill2.GetHashCode(), Is.EqualTo(fill1.GetHashCode()));
        }

        [Test]
        public void BackgroundPatternNotEqualCheck()
        {
            var fill1 = new XLFill { PatternType = XLFillPatternValues.Solid, BackgroundColor = XLColor.Blue };
            var fill2 = new XLFill { PatternType = XLFillPatternValues.Solid, BackgroundColor = XLColor.Red };
            Assert.That(fill1.Equals(fill2), Is.False);
            Assert.That(fill2.GetHashCode(), Is.Not.EqualTo(fill1.GetHashCode()));
        }

        [Test]
        public void FillsWithTransparentColorEqual()
        {
            var fill1 = new XLFill { BackgroundColor = XLColor.ElectricUltramarine, PatternType = XLFillPatternValues.None };
            var fill2 = new XLFill { BackgroundColor = XLColor.EtonBlue, PatternType = XLFillPatternValues.None };
            var fill3 = new XLFill { BackgroundColor = XLColor.FromIndex(64) };
            var fill4 = new XLFill { BackgroundColor = XLColor.NoColor };

            Assert.That(fill1.Equals(fill2), Is.True);
            Assert.That(fill1.Equals(fill3), Is.True);
            Assert.That(fill1.Equals(fill4), Is.True);
            Assert.That(fill2.GetHashCode(), Is.EqualTo(fill1.GetHashCode()));
            Assert.That(fill3.GetHashCode(), Is.EqualTo(fill1.GetHashCode()));
            Assert.That(fill4.GetHashCode(), Is.EqualTo(fill1.GetHashCode()));
        }

        [Test]
        public void SolidFillsWithDifferentPatternColorEqual()
        {
            var fill1 = new XLFill
            {
                PatternType = XLFillPatternValues.Solid,
                BackgroundColor = XLColor.Red,
                PatternColor = XLColor.Blue
            };

            var fill2 = new XLFill
            {
                PatternType = XLFillPatternValues.Solid,
                BackgroundColor = XLColor.Red,
                PatternColor = XLColor.Green
            };

            Assert.That(fill1.Equals(fill2), Is.True);
            Assert.That(fill2.GetHashCode(), Is.EqualTo(fill1.GetHashCode()));
        }

        [Test]
        public void BackgroundWithConditionalFormat()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Test");
            worksheet.Cell(2, 2).SetValue("Text");
            var cf = worksheet.Cell(2, 2).AddConditionalFormat();
            var style = cf.WhenNotBlank();
            style = style
                .Border.SetOutsideBorder(XLBorderStyleValues.Thick)
                .Border.SetOutsideBorderColor(XLColor.Blue);

            Assert.That(XLBorderStyleValues.Thick, Is.EqualTo(style.Border.BottomBorder));
            Assert.That(XLBorderStyleValues.Thick, Is.EqualTo(style.Border.TopBorder));
            Assert.That(XLBorderStyleValues.Thick, Is.EqualTo(style.Border.LeftBorder));
            Assert.That(XLBorderStyleValues.Thick, Is.EqualTo(style.Border.RightBorder));

            Assert.That(XLColor.Blue, Is.EqualTo(style.Border.BottomBorderColor));
            Assert.That(XLColor.Blue, Is.EqualTo(style.Border.TopBorderColor));
            Assert.That(XLColor.Blue, Is.EqualTo(style.Border.LeftBorderColor));
            Assert.That(XLColor.Blue, Is.EqualTo(style.Border.RightBorderColor));
        }

        [Test]
        public void LoadAndSaveTransparentBackgroundFill()
        {
            var expectedFilePath = @"Other\StyleReferenceFiles\TransparentBackgroundFill\TransparentBackgroundFill.xlsx";

            using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\StyleReferenceFiles\TransparentBackgroundFill\inputfile.xlsx"));
            using var ms = new MemoryStream();

            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook(stream);

                wb.SaveAs(ms);

                //Uncomment to replace expectation running.net6.0,
                //var expectedFileInVsSolution = Path.GetFullPath(Path.Combine("../../../", "Resource", expectedFilePath));
                //wb.SaveAs(expectedFileInVsSolution);
                return wb;
            }, expectedFilePath);
        }
    }
}