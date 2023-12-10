using ClosedXML.Excel.InsertData;
using ClosedXML.Tests.Excel.Tables;
using NUnit.Framework;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Tests.Excel.InsertData
{
    public class ObjectReaderTests
    {
        private static readonly TablesTests.TestObjectWithAttributes[] ObjectWithAttributes =
        {
            new TablesTests.TestObjectWithAttributes
            {
                Column1 = "Value 1",
                Column2 = "Value 2",
                UnOrderedColumn = 3,
                MyField = 4,
            },
            new TablesTests.TestObjectWithAttributes
            {
                Column1 = "Value 5",
                Column2 = "Value 6",
                UnOrderedColumn = 7,
                MyField = 8,
            }
        };

        private static readonly TablesTests.TestObjectWithoutAttributes[] ObjectWithoutAttributes =
        {
            new TablesTests.TestObjectWithoutAttributes
            {
                Column1 = "Value 9",
                Column2 = "Value 10"
            },
            new TablesTests.TestObjectWithoutAttributes
            {
                Column1 = "Value 11",
                Column2 = "Value 12"
            }
        };

        private static readonly TestPoint[] Structs =
        {
            new TestPoint
            {
                X = 1,
                Y = 2,
                Z = 3
            },
            new TestPoint(),
        };

        private static readonly TestPoint?[] NullableStructs =
        {
            new TestPoint
            {
                X = 1,
                Y = 2,
                Z = 3
            },
            new TestPoint(),
            null
        };

        [TestCaseSource(nameof(ObjectSourceNames))]
        public string CanGetPropertyName<T>(IEnumerable<T> data, int propertyIndex)
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(data);
            return reader.GetPropertyName(propertyIndex);
        }

        private static IEnumerable<TestCaseData> ObjectSourceNames
        {
            get
            {
                IEnumerable data = ObjectWithoutAttributes;
                yield return new TestCaseData(data, 0).Returns("Column1");
                yield return new TestCaseData(data, 1).Returns("Column2");

                data = ObjectWithAttributes;
                yield return new TestCaseData(data, 0).Returns("FirstColumn");
                yield return new TestCaseData(data, 1).Returns("SecondColumn");
                yield return new TestCaseData(data, 2).Returns("SomeFieldNotProperty");
                yield return new TestCaseData(data, 3).Returns("UnOrderedColumn");

                data = Structs;
                yield return new TestCaseData(data, 0).Returns("X");
                yield return new TestCaseData(data, 1).Returns("Y");
                yield return new TestCaseData(data, 2).Returns("Z");

                data = NullableStructs;
                yield return new TestCaseData(data, 0).Returns("X");
                yield return new TestCaseData(data, 1).Returns("Y");
                yield return new TestCaseData(data, 2).Returns("Z");
            }
        }

        [TestCaseSource(nameof(PropertyCounts))]
        public int CanGetPropertiesCount(IEnumerable data)
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(data);
            return reader.GetPropertiesCount();
        }

        private static IEnumerable<TestCaseData> PropertyCounts
        {
            get
            {
                IEnumerable data = ObjectWithoutAttributes;
                yield return new TestCaseData(data).Returns(2);

                data = ObjectWithAttributes;
                yield return new TestCaseData(data).Returns(4);

                data = Structs;
                yield return new TestCaseData(data).Returns(3);

                data = NullableStructs;
                yield return new TestCaseData(data).Returns(3);
            }
        }

        [Test]
        public void CanGetRecordsCount()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(ObjectWithAttributes);
            Assert.That(reader.GetRecordsCount(), Is.EqualTo(2));
        }

        [Test]
        public void CanReadValues_FromObject()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(ObjectWithAttributes);
            var result = reader.GetData();

            var firstRecord = result.First().ToArray();
            var lastRecord = result.Last().ToArray();

            Assert.That(firstRecord[0], Is.EqualTo("Value 2"));
            Assert.That(firstRecord[1], Is.EqualTo("Value 1"));
            Assert.That(firstRecord[2], Is.EqualTo(4));
            Assert.That(firstRecord[3], Is.EqualTo(3));

            Assert.That(lastRecord[0], Is.EqualTo("Value 6"));
            Assert.That(lastRecord[1], Is.EqualTo("Value 5"));
            Assert.That(lastRecord[2], Is.EqualTo(8));
            Assert.That(lastRecord[3], Is.EqualTo(7));
        }

        [Test]
        public void CanReadValues_FromStruct()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(Structs);
            var result = reader.GetData();

            var firstRecord = result.First().ToArray();
            var lastRecord = result.Last().ToArray();

            Assert.That(firstRecord[0], Is.EqualTo(1));
            Assert.That(firstRecord[1], Is.EqualTo(2));
            Assert.That(firstRecord[2], Is.EqualTo(3));

            Assert.That(lastRecord[0], Is.EqualTo(0));
            Assert.That(lastRecord[1], Is.EqualTo(0));
            Assert.That(lastRecord[2], Is.EqualTo(null));
        }

        [Test]
        public void CanReadValues_FromNullableStruct()
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(NullableStructs);
            var result = reader.GetData();

            var firstRecord = result.First().ToArray();
            var lastRecord = result.Last().ToArray();

            Assert.That(firstRecord[0], Is.EqualTo(1));
            Assert.That(firstRecord[1], Is.EqualTo(2));
            Assert.That(firstRecord[2], Is.EqualTo(3));

            Assert.That(lastRecord[0], Is.EqualTo(null));
            Assert.That(lastRecord[1], Is.EqualTo(null));
            Assert.That(lastRecord[2], Is.EqualTo(null));
        }

        private struct TestPoint
        {
            public double X { get; set; }
            public double Y { get; set; }
            public double? Z { get; set; }
        }
    }
}
