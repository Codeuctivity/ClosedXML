using ClosedXML.Examples;
using ClosedXML.Examples.Misc;
using ClosedXML.Examples.Ranges;
using NUnit.Framework;
using System.Runtime.InteropServices;

namespace ClosedXML.Tests.Examples
{
    [TestFixture]
    public class RangesTests
    {
        [Test]
        public void ClearingRanges()
        {
            TestHelper.RunTestExample<ClearingRanges>(@"Ranges\ClearingRanges.xlsx");
        }

        [Test]
        public void CopyingRanges()
        {
            TestHelper.RunTestExample<CopyingRanges>(@"Ranges\CopyingRanges.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void CurrentRowColumn()
        {
            TestHelper.RunTestExample<CurrentRowColumn>(@"Ranges\CurrentRowColumn.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void DefiningRanges()
        {
            TestHelper.RunTestExample<DefiningRanges>(@"Ranges\DefiningRanges.xlsx");
        }

        [Test]
        public void DeletingRanges()
        {
            TestHelper.RunTestExample<DeletingRanges>(@"Ranges\DeletingRanges.xlsx");
        }

        [Test]
        public void InsertingDeletingColumns()
        {
            TestHelper.RunTestExample<InsertingDeletingColumns>(@"Ranges\InsertingDeletingColumns.xlsx");
        }

        [Test]
        public void InsertingDeletingRows()
        {
            TestHelper.RunTestExample<InsertingDeletingRows>(@"Ranges\InsertingDeletingRows.xlsx");
        }

        [Test]
        public void MultipleRanges()
        {
            TestHelper.RunTestExample<MultipleRanges>(@"Ranges\MultipleRanges.xlsx");
        }

        [Test]
        public void NamedRanges()
        {
            TestHelper.RunTestExample<NamedRanges>(@"Ranges\NamedRanges.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void SelectingRanges()
        {
            TestHelper.RunTestExample<SelectingRanges>(@"Ranges\SelectingRanges.xlsx");
        }

        [Test]
        public void ShiftingRanges()
        {
            TestHelper.RunTestExample<ShiftingRanges>(@"Ranges\ShiftingRanges.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void SortExample()
        {
            TestHelper.RunTestExample<SortExample>(@"Ranges\SortExample.xlsx");
        }

        [Test]
        public void Sorting()
        {
            TestHelper.RunTestExample<Sorting>(@"Ranges\Sorting.xlsx");
        }

        [Test]
        public void TransposeRanges()
        {
            TestHelper.RunTestExample<TransposeRanges>(@"Ranges\TransposeRanges.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void TransposeRangesPlus()
        {
            TestHelper.RunTestExample<TransposeRangesPlus>(@"Ranges\TransposeRangesPlus.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void AddingRowToTables()
        {
            TestHelper.RunTestExample<AddingRowToTables>(@"Ranges\AddingRowToTables.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void WalkingRanges()
        {
            TestHelper.RunTestExample<WalkingRanges>(@"Ranges\WalkingRanges.xlsx");
        }
    }
}
