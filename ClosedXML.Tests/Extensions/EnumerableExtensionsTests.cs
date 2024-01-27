using ClosedXML.Excel;
using ClosedXML.Tests.Excel.Tables;
using NUnit.Framework;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Tests.Extensions
{
    public class EnumerableExtensionsTests
    {
        [Test]
        public void CanGetItemType()
        {
            var array = new int[0];
            Assert.That(array.GetItemType(), Is.EqualTo(typeof(int)));

            var list = new List<double>();
            Assert.That(list.GetItemType(), Is.EqualTo(typeof(double)));
            Assert.That(list.AsEnumerable().GetItemType(), Is.EqualTo(typeof(double)));

            IEnumerable<IEnumerable> enumerable = new List<string>();
            Assert.That(enumerable.GetItemType(), Is.EqualTo(typeof(string)));

            enumerable = new List<List<string>>();
            Assert.That(enumerable.GetItemType(), Is.EqualTo(typeof(List<string>)));

            enumerable = new List<int[]>();
            Assert.That(enumerable.GetItemType(), Is.EqualTo(typeof(int[])));

            var anonymousIterator = new List<TablesTests.TestObjectWithoutAttributes>()
                .Select(o => new { FirstName = o.Column1, LastName = o.Column2 });

            //expectedType can be something like <>f__AnonymousType9`2[System.String,System.String]
            //but since that `9` may differ with new anonymous types declare in the assembly
            //check the beginning and the ending of the actual type
            var expectedTypeStart = "<>f__AnonymousType";
            var expectedTypeEnd = "`2[System.String,System.String]";
            var actualType = anonymousIterator.GetItemType().ToString();
            Assert.That(actualType.StartsWith(expectedTypeStart), Is.True);
            Assert.That(actualType.EndsWith(expectedTypeEnd), Is.True);

            IEnumerable<object> obj = anonymousIterator;
            actualType = obj.GetItemType().ToString();
            Assert.That(actualType.StartsWith(expectedTypeStart), Is.True);
            Assert.That(actualType.EndsWith(expectedTypeEnd), Is.True);
        }
    }
}
