using ClosedXML.Excel.Ranges;
using ClosedXML.Excel.Tables;
using System;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLTableRange : XLRange, IXLTableRange
    {
        private readonly XLTable _table;
        private readonly XLRange _range;
        public XLTableRange(XLRange range, XLTable table)
            : base(new XLRangeParameters(range.RangeAddress, range.Style))
        {
            _table = table;
            _range = range;
        }

        IXLTableRow IXLTableRange.FirstRow(Func<IXLTableRow, bool> predicate)
        {
            return FirstRow(predicate);
        }
        public XLTableRow FirstRow(Func<IXLTableRow, bool> predicate = null)
        {
            if (predicate == null)
            {
                return new XLTableRow(this, _range.FirstRow());
            }

            var rowCount = _range.RowCount();

            for (var ro = 1; ro <= rowCount; ro++)
            {
                var row = new XLTableRow(this, _range.Row(ro));
                if (predicate(row))
                {
                    return row;
                }
            }

            return null;
        }

        IXLTableRow IXLTableRange.FirstRowUsed(Func<IXLTableRow, bool> predicate)
        {
            return FirstRowUsed(XLCellsUsedOptions.AllContents, predicate);
        }
        public XLTableRow FirstRowUsed(Func<IXLTableRow, bool> predicate = null)
        {
            return FirstRowUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLTableRow IXLTableRange.FirstRowUsed(bool includeFormats, Func<IXLTableRow, bool> predicate)
        {
            return FirstRowUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents,
                predicate);
        }

        IXLTableRow IXLTableRange.FirstRowUsed(XLCellsUsedOptions options, Func<IXLTableRow, bool> predicate)
        {
            return FirstRowUsed(options, predicate);
        }

        internal XLTableRow FirstRowUsed(XLCellsUsedOptions options, Func<IXLTableRow, bool> predicate = null)
        {
            if (predicate == null)
            {
                return new XLTableRow(this, _range.FirstRowUsed(options));
            }

            var rowCount = _range.RowCount();

            for (var ro = 1; ro <= rowCount; ro++)
            {
                var row = new XLTableRow(this, _range.Row(ro));

                if (!row.IsEmpty(options) && predicate(row))
                {
                    return row;
                }
            }

            return null;
        }


        IXLTableRow IXLTableRange.LastRow(Func<IXLTableRow, bool> predicate)
        {
            return LastRow(predicate);
        }
        public XLTableRow LastRow(Func<IXLTableRow, bool> predicate = null)
        {
            if (predicate == null)
            {
                return new XLTableRow(this, _range.LastRow());
            }

            var rowCount = _range.RowCount();

            for (var ro = rowCount; ro >= 1; ro--)
            {
                var row = new XLTableRow(this, _range.Row(ro));
                if (predicate(row))
                {
                    return row;
                }
            }
            return null;
        }

        IXLTableRow IXLTableRange.LastRowUsed(Func<IXLTableRow, bool> predicate)
        {
            return LastRowUsed(XLCellsUsedOptions.AllContents, predicate);
        }
        public XLTableRow LastRowUsed(Func<IXLTableRow, bool> predicate = null)
        {
            return LastRowUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLTableRow IXLTableRange.LastRowUsed(bool includeFormats, Func<IXLTableRow, bool> predicate)
        {
            return LastRowUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents,
                predicate);
        }

        IXLTableRow IXLTableRange.LastRowUsed(XLCellsUsedOptions options, Func<IXLTableRow, bool> predicate)
        {
            return LastRowUsed(options, predicate);
        }


        internal XLTableRow LastRowUsed(XLCellsUsedOptions options, Func<IXLTableRow, bool> predicate = null)
        {
            if (predicate == null)
            {
                return new XLTableRow(this, _range.LastRowUsed(options));
            }

            var rowCount = _range.RowCount();

            for (var ro = rowCount; ro >= 1; ro--)
            {
                var row = new XLTableRow(this, _range.Row(ro));

                if (!row.IsEmpty(options) && predicate(row))
                {
                    return row;
                }
            }

            return null;
        }

        IXLTableRow IXLTableRange.Row(int row)
        {
            return Row(row);
        }
        public new XLTableRow Row(int row)
        {
            if (row <= 0 || row > XLHelper.MaxRowNumber + RangeAddress.FirstAddress.RowNumber - 1)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(row),
                    string.Format("Row number must be between 1 and {0}", XLHelper.MaxRowNumber + RangeAddress.FirstAddress.RowNumber - 1)
                );
            }

            return new XLTableRow(this, base.Row(row));
        }

        public IXLTableRows Rows(Func<IXLTableRow, bool> predicate = null)
        {
            var retVal = new XLTableRows(Worksheet.Style);
            var rowCount = _range.RowCount();

            for (var r = 1; r <= rowCount; r++)
            {
                var row = Row(r);
                if (predicate == null || predicate(row))
                {
                    retVal.Add(row);
                }
            }
            return retVal;
        }

        public new IXLTableRows Rows(int firstRow, int lastRow)
        {
            var retVal = new XLTableRows(Worksheet.Style);

            for (var rowNumber = firstRow; rowNumber <= lastRow; rowNumber++)
            {
                retVal.Add(Row(rowNumber));
            }

            return retVal;
        }

        public new IXLTableRows Rows(string rows)
        {
            var retVal = new XLTableRows(Worksheet.Style);
            var rowPairs = rows.Split(',');
            foreach (var tPair in rowPairs.Select(pair => pair.Trim()))
            {
                string firstRow;
                string lastRow;
                if (tPair.Contains(':') || tPair.Contains('-'))
                {
                    var rowRange = XLHelper.SplitRange(tPair);

                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = tPair;
                    lastRow = tPair;
                }
                foreach (var row in Rows(int.Parse(firstRow), int.Parse(lastRow)))
                {
                    retVal.Add(row);
                }
            }
            return retVal;
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public IXLTableRows RowsUsed(bool includeFormats, Func<IXLTableRow, bool> predicate = null)
        {
            return RowsUsed(includeFormats
                ? XLCellsUsedOptions.AllContents
                : XLCellsUsedOptions.All,
                predicate);
        }

        IXLTableRows IXLTableRange.RowsUsed(XLCellsUsedOptions options, Func<IXLTableRow, bool> predicate)
        {
            return RowsUsed(options, predicate);
        }

        internal XLTableRows RowsUsed(XLCellsUsedOptions options, Func<IXLTableRow, bool> predicate = null)
        {
            var rows = new XLTableRows(Worksheet.Style);
            var rowCount = RowCount();

            for (var ro = 1; ro <= rowCount; ro++)
            {
                var row = Row(ro);

                if (!row.IsEmpty(options) && (predicate == null || predicate(row)))
                {
                    rows.Add(row);
                }
            }
            return rows;
        }

        IXLTableRows IXLTableRange.RowsUsed(Func<IXLTableRow, bool> predicate)
        {
            return RowsUsed(predicate);
        }
        public IXLTableRows RowsUsed(Func<IXLTableRow, bool> predicate = null)
        {
            return RowsUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        IXLTable IXLTableRange.Table => _table;
        public XLTable Table => _table;

        public new IXLTableRows InsertRowsAbove(int numberOfRows)
        {
            return XLHelper.InsertRowsWithoutEvents(InsertRowsAbove, this, numberOfRows, !Table.ShowTotalsRow );
        }
        public new IXLTableRows InsertRowsBelow(int numberOfRows)
        {
            return XLHelper.InsertRowsWithoutEvents(InsertRowsBelow, this, numberOfRows, !Table.ShowTotalsRow);
        }


        public new IXLRangeColumn Column(string column)
        {
            if (XLHelper.IsValidColumn(column))
            {
                var coNum = XLHelper.GetColumnNumberFromLetter(column);
                return coNum > ColumnCount() ? Column(_table.GetFieldIndex(column) + 1) : Column(coNum);
            }

            return Column(_table.GetFieldIndex(column) + 1);
        }
    }
}
