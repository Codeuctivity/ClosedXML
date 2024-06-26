using ClosedXML.Excel.Ranges;
using System;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRangeRow : XLRangeBase, IXLRangeRow
    {
        #region Constructor

        /// <summary>
        /// The direct contructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
        /// </summary>
        public XLRangeRow(XLRangeParameters rangeParameters)
            : base(rangeParameters.RangeAddress, (rangeParameters.DefaultStyle as XLStyle).Value)
        {
        }

        #endregion Constructor

        #region IXLRangeRow Members

        public IXLCell Cell(int column)
        {
            return Cell(1, column);
        }

        public override XLCell Cell(string columnLetter)
        {
            return Cell(1, columnLetter);
        }

        IXLCell IXLRangeRow.Cell(string columnLetter)
        {
            return Cell(columnLetter);
        }

        public void Delete()
        {
            Delete(XLShiftDeletedCells.ShiftCellsUp);
        }

        public IXLCells InsertCellsAfter(int numberOfColumns)
        {
            return InsertCellsAfter(numberOfColumns, true);
        }

        public IXLCells InsertCellsAfter(int numberOfColumns, bool expandRange)
        {
            return InsertColumnsAfter(numberOfColumns, expandRange).Cells();
        }

        public IXLCells InsertCellsBefore(int numberOfColumns)
        {
            return InsertCellsBefore(numberOfColumns, false);
        }

        public IXLCells InsertCellsBefore(int numberOfColumns, bool expandRange)
        {
            return InsertColumnsBefore(numberOfColumns, expandRange).Cells();
        }

        public override IXLCells Cells(string cellsInRow)
        {
            var retVal = new XLCells(false, XLCellsUsedOptions.AllContents);
            var rangePairs = cellsInRow.Split(',');
            foreach (var pair in rangePairs)
            {
                retVal.Add(Range(pair.Trim()).RangeAddress);
            }

            return retVal;
        }

        public IXLCells Cells(int firstColumn, int lastColumn)
        {
            return Cells(firstColumn + ":" + lastColumn);
        }

        public IXLCells Cells(string firstColumn, string lastColumn)
        {
            return Cells(XLHelper.GetColumnNumberFromLetter(firstColumn) + ":"
                         + XLHelper.GetColumnNumberFromLetter(lastColumn));
        }

        public int CellCount()
        {
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public new IXLRangeRow Sort()
        {
            return SortLeftToRight();
        }

        public new IXLRangeRow SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true)
        {
            base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
            return this;
        }

        public new IXLRangeRow CopyTo(IXLCell target)
        {
            base.CopyTo(target);

            var lastRowNumber = target.Address.RowNumber + RowCount() - 1;
            if (lastRowNumber > XLHelper.MaxRowNumber)
            {
                lastRowNumber = XLHelper.MaxRowNumber;
            }

            var lastColumnNumber = target.Address.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > XLHelper.MaxColumnNumber)
            {
                lastColumnNumber = XLHelper.MaxColumnNumber;
            }

            return target.Worksheet.Range(
                target.Address.RowNumber,
                target.Address.ColumnNumber,
                lastRowNumber,
                lastColumnNumber)
                .Row(1);
        }

        public new IXLRangeRow CopyTo(IXLRangeBase target)
        {
            base.CopyTo(target);
            var lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
            if (lastRowNumber > XLHelper.MaxRowNumber)
            {
                lastRowNumber = XLHelper.MaxRowNumber;
            }

            var lastColumnNumber = target.RangeAddress.LastAddress.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > XLHelper.MaxColumnNumber)
            {
                lastColumnNumber = XLHelper.MaxColumnNumber;
            }

            return target.Worksheet.Range(
                target.RangeAddress.FirstAddress.RowNumber,
                target.RangeAddress.LastAddress.ColumnNumber,
                lastRowNumber,
                lastColumnNumber)
                .Row(1);
        }

        public IXLRangeRow Row(int start, int end)
        {
            return Range(1, start, 1, end).Row(1);
        }

        public IXLRangeRow Row(IXLCell start, IXLCell end)
        {
            return Row(start.Address.ColumnNumber, end.Address.ColumnNumber);
        }

        public IXLRangeRows Rows(string rows)
        {
            var retVal = new XLRangeRows();
            var columnPairs = rows.Split(',');
            foreach (var trimmedPair in columnPairs.Select(pair => pair.Trim()))
            {
                string firstColumn;
                string lastColumn;
                if (trimmedPair.Contains(':') || trimmedPair.Contains('-'))
                {
                    var columnRange = trimmedPair.Contains('-')
                                          ? trimmedPair.Replace('-', ':').Split(':')
                                          : trimmedPair.Split(':');
                    firstColumn = columnRange[0];
                    lastColumn = columnRange[1];
                }
                else
                {
                    firstColumn = trimmedPair;
                    lastColumn = trimmedPair;
                }

                retVal.Add(Range(firstColumn, lastColumn).FirstRow());
            }

            return retVal;
        }

        public IXLRangeRow SetDataType(XLDataType dataType)
        {
            DataType = dataType;
            return this;
        }

        public IXLRow WorksheetRow()
        {
            return Worksheet.Row(RangeAddress.FirstAddress.RowNumber);
        }

        #endregion IXLRangeRow Members
        public override XLRangeType RangeType => XLRangeType.RangeRow;

        internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
        {
            RangeAddress = (XLRangeAddress)ShiftColumns(RangeAddress, range, columnsShifted);
        }

        internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
        {
            RangeAddress = (XLRangeAddress)ShiftRows(RangeAddress, range, rowsShifted);
        }

        public IXLRange Range(int firstColumn, int lastColumn)
        {
            return Range(1, firstColumn, 1, lastColumn);
        }

        public override XLRange Range(string rangeAddressStr)
        {
            string rangeAddressToUse;
            if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
            {
                if (rangeAddressStr.Contains('-'))
                {
                    rangeAddressStr = rangeAddressStr.Replace('-', ':');
                }

                var arrRange = rangeAddressStr.Split(':');
                var firstPart = arrRange[0];
                var secondPart = arrRange[1];
                rangeAddressToUse = FixRowAddress(firstPart) + ":" + FixRowAddress(secondPart);
            }
            else
            {
                rangeAddressToUse = FixRowAddress(rangeAddressStr);
            }

            var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressToUse);
            return Range(rangeAddress);
        }

        public int CompareTo(XLRangeRow otherRow, IXLSortElements columnsToSort)
        {
            foreach (var e in columnsToSort)
            {
                var thisCell = (XLCell)Cell(e.ElementNumber);
                var otherCell = (XLCell)otherRow.Cell(e.ElementNumber);
                int comparison;
                var thisCellIsBlank = thisCell.IsEmpty();
                var otherCellIsBlank = otherCell.IsEmpty();
                if (e.IgnoreBlanks && (thisCellIsBlank || otherCellIsBlank))
                {
                    if (thisCellIsBlank && otherCellIsBlank)
                    {
                        comparison = 0;
                    }
                    else
                    {
                        if (thisCellIsBlank)
                        {
                            comparison = e.SortOrder == XLSortOrder.Ascending ? 1 : -1;
                        }
                        else
                        {
                            comparison = e.SortOrder == XLSortOrder.Ascending ? -1 : 1;
                        }
                    }
                }
                else
                {
                    if (thisCell.DataType == otherCell.DataType)
                    {
                        switch (thisCell.DataType)
                        {
                            case XLDataType.Text:
                                comparison = e.MatchCase
                                                 ? thisCell.InnerText.CompareTo(otherCell.InnerText)
                                                 : string.Compare(thisCell.InnerText, otherCell.InnerText, true);
                                break;

                            case XLDataType.TimeSpan:
                                comparison = thisCell.GetTimeSpan().CompareTo(otherCell.GetTimeSpan());
                                break;

                            case XLDataType.DateTime:
                                comparison = thisCell.GetDateTime().CompareTo(otherCell.GetDateTime());
                                break;

                            case XLDataType.Number:
                                comparison = thisCell.GetDouble().CompareTo(otherCell.GetDouble());
                                break;

                            case XLDataType.Boolean:
                                comparison = thisCell.GetBoolean().CompareTo(otherCell.GetBoolean());
                                break;

                            default:
                                throw new NotImplementedException();
                        }
                    }
                    else if (e.MatchCase)
                    {
                        comparison = string.Compare(thisCell.GetString(), otherCell.GetString(), true);
                    }
                    else
                    {
                        comparison = thisCell.GetString().CompareTo(otherCell.GetString());
                    }
                }

                if (comparison != 0)
                {
                    return e.SortOrder == XLSortOrder.Ascending ? comparison : -comparison;
                }
            }

            return 0;
        }

        private XLRangeRow RowShift(int rowsToShift)
        {
            var rowNum = RowNumber() + rowsToShift;

            var range = Worksheet.Range(
                rowNum,
                RangeAddress.FirstAddress.ColumnNumber,
                rowNum,
                RangeAddress.LastAddress.ColumnNumber);

            return range.FirstRow();
        }

        #region XLRangeRow Above

        IXLRangeRow IXLRangeRow.RowAbove()
        {
            return RowAbove();
        }

        IXLRangeRow IXLRangeRow.RowAbove(int step)
        {
            return RowAbove(step);
        }

        public XLRangeRow RowAbove()
        {
            return RowAbove(1);
        }

        public XLRangeRow RowAbove(int step)
        {
            return RowShift(step * -1);
        }

        #endregion XLRangeRow Above

        #region XLRangeRow Below

        IXLRangeRow IXLRangeRow.RowBelow()
        {
            return RowBelow();
        }

        IXLRangeRow IXLRangeRow.RowBelow(int step)
        {
            return RowBelow(step);
        }

        public XLRangeRow RowBelow()
        {
            return RowBelow(1);
        }

        public XLRangeRow RowBelow(int step)
        {
            return RowShift(step);
        }

        #endregion XLRangeRow Below

        public new IXLRangeRow Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            base.Clear(clearOptions);
            return this;
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public IXLRangeRow RowUsed(bool includeFormats)
        {
            return RowUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents);
        }

        public IXLRangeRow RowUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents)
        {
            return Row((this as IXLRangeBase).FirstCellUsed(options),
                       (this as IXLRangeBase).LastCellUsed(options));
        }
    }
}
