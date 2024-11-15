
using System;

namespace ClosedXML.Excel
{
    public interface IXLTableRange : IXLRange
    {
        IXLTable Table { get; }

        IXLTableRow FirstRow(Func<IXLTableRow, bool> predicate = null);

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLTableRow FirstRowUsed(bool includeFormats, Func<IXLTableRow, bool> predicate = null);

        IXLTableRow FirstRowUsed(XLCellsUsedOptions options, Func<IXLTableRow, bool> predicate = null);

        IXLTableRow FirstRowUsed(Func<IXLTableRow, bool> predicate = null);

        new IXLTableRows InsertRowsAbove(int numberOfRows);

        new IXLTableRows InsertRowsBelow(int numberOfRows);

        IXLTableRow LastRow(Func<IXLTableRow, bool> predicate = null);

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLTableRow LastRowUsed(bool includeFormats, Func<IXLTableRow, bool> predicate = null);

        IXLTableRow LastRowUsed(XLCellsUsedOptions options, Func<IXLTableRow, bool> predicate = null);

        IXLTableRow LastRowUsed(Func<IXLTableRow, bool> predicate = null);

        /// <summary>
        /// Rows the specified row.
        /// </summary>
        /// <param name="row">1-based row number relative to the first row of this range.</param>
        /// <returns></returns>
        new IXLTableRow Row(int row);

        IXLTableRows Rows(Func<IXLTableRow, bool> predicate = null);

        /// <summary>
        /// Returns a subset of the rows
        /// </summary>
        /// <param name="firstRow">The first row to return. 1-based row number relative to the first row of this range.</param>
        /// <param name="lastRow">The last row to return. 1-based row number relative to the first row of this range.</param>
        /// <returns></returns>
        new IXLTableRows Rows(int firstRow, int lastRow);

        new IXLTableRows Rows(string rows);

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLTableRows RowsUsed(bool includeFormats, Func<IXLTableRow, bool> predicate = null);

        IXLTableRows RowsUsed(XLCellsUsedOptions options, Func<IXLTableRow, bool> predicate = null);

        IXLTableRows RowsUsed(Func<IXLTableRow, bool> predicate = null);
    }
}
