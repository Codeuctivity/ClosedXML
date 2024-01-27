using ClosedXML.Utils;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Misc
{
    [TestFixture]
    public class XmlEncoderTest
    {
        [Test]
        public void TestControlChars()
        {
            Assert.That(XmlEncoder.EncodeString("\u0001 \u0002 \u0003 \u0004"), Is.EqualTo("_x0001_ _x0002_ _x0003_ _x0004_"));
            Assert.That(XmlEncoder.EncodeString("\u0005 \u0006 \u0007 \u0008"), Is.EqualTo("_x0005_ _x0006_ _x0007_ _x0008_"));

            Assert.That(XmlEncoder.DecodeString("_x0001_ _x0002_ _x0003_ _x0004_"), Is.EqualTo("\u0001 \u0002 \u0003 \u0004"));
            Assert.That(XmlEncoder.DecodeString("_x0005_ _x0006_ _x0007_ _x0008_"), Is.EqualTo("\u0005 \u0006 \u0007 \u0008"));
            Assert.That(XmlEncoder.DecodeString("_xaaBB_ _xAAbb_"), Is.EqualTo("\uAABB \uAABB"));

            // https://github.com/ClosedXML/ClosedXML/issues/1154
            Assert.That(XmlEncoder.DecodeString("_Xceed_Something"), Is.EqualTo("_Xceed_Something"));
        }
    }
}
