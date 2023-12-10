using ClosedXML.Excel.CalcEngine.Functions;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class XLMathTests
    {
        [Test]
        public void IsEven()
        {
            Assert.That(XLMath.IsEven(2), Is.True);
            Assert.That(XLMath.IsEven(3), Is.False);
        }

        [Test]
        public void IsOdd()
        {
            Assert.That(XLMath.IsOdd(3), Is.True);
            Assert.That(XLMath.IsOdd(2), Is.False);
        }
    }
}
