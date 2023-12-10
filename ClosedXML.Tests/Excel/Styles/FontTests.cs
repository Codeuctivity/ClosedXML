using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Styles
{
    public class FontTests
    {
        [Test]
        public void XLFontKey_GetHashCode_IsCaseInsensitive()
        {
            var fontKey1 = new XLFontKey { FontName = "Arial" };
            var fontKey2 = new XLFontKey { FontName = "Times New Roman" };
            var fontKey3 = new XLFontKey { FontName = "TIMES NEW ROMAN" };

            Assert.That(fontKey2.GetHashCode(), Is.Not.EqualTo(fontKey1.GetHashCode()));
            Assert.That(fontKey3.GetHashCode(), Is.EqualTo(fontKey2.GetHashCode()));
        }

        [Test]
        public void XLFontKey_Equals_IsCaseInsensitive()
        {
            var fontKey1 = new XLFontKey { FontName = "Arial" };
            var fontKey2 = new XLFontKey { FontName = "Times New Roman" };
            var fontKey3 = new XLFontKey { FontName = "TIMES NEW ROMAN" };

            Assert.That(fontKey1.Equals(fontKey2), Is.False);
            Assert.That(fontKey2.Equals(fontKey3), Is.True);
        }
    }
}
