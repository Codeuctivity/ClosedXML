using ClosedXML.Excel.Ranges;
using ClosedXML.Excel.Ranges.Sort;
using ClosedXML.Excel.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal abstract class XLRangeBase : XLStylizedBase, IXLRangeBase, IXLStylized
    {
        #region Fields

        private XLSortElements _sortRows;
        private XLSortElements _sortColumns;

        #endregion Fields

        protected IXLStyle GetStyle()
        {
            return Style;
        }

        #region Constructor

        private static int IdCounter;
        private readonly int Id;

        protected XLRangeBase(XLRangeAddress rangeAddress, XLStyleValue styleValue)
            : base(styleValue)
        {
            Id = ++IdCounter;

            _rangeAddress = rangeAddress;
        }

        #endregion Constructor

        protected virtual void OnRangeAddressChanged(XLRangeAddress oldAddress, XLRangeAddress newAddress)
        {
            Worksheet.RellocateRange(RangeType, oldAddress, newAddress);
        }

        #region Public properties

        private XLRangeAddress _rangeAddress;

        public XLRangeAddress RangeAddress
        {
            get { return _rangeAddress; }
            protected set
            {
                if (_rangeAddress != value)
                {
                    var oldAddress = _rangeAddress;
                    _rangeAddress = value;
                    OnRangeAddressChanged(oldAddress, _rangeAddress);
                }
            }
        }

        public XLWorksheet Worksheet => RangeAddress.Worksheet;

        public IXLDataValidation CreateDataValidation()
        {
            var newRange = AsRange();
            var dataValidation = new XLDataValidation(newRange);
            Worksheet.DataValidations.Add(dataValidation);
            return dataValidation;
        }

        public IXLDataValidation GetDataValidation()
        {
            Worksheet.DataValidations.TryGet(RangeAddress, out var existingDataValidation);
            return existingDataValidation;
        }

        #region IXLRangeBase Members

        IXLRangeAddress IXLAddressable.RangeAddress => RangeAddress;

        IXLWorksheet IXLRangeBase.Worksheet => RangeAddress.Worksheet;

        public string FormulaA1
        {
            set
            {
                Cells().ForEach(c =>
                                    {
                                        c.FormulaA1 = value;
                                        c.FormulaReference = RangeAddress;
                                    });
            }
        }

        public string FormulaR1C1
        {
            set
            {
                Cells().ForEach(c =>
                {
                    c.FormulaR1C1 = value;
                    c.FormulaReference = RangeAddress;
                });
            }
        }

        public bool ShareString
        {
            set { Cells().ForEach(c => c.ShareString = value); }
        }

        public IXLHyperlinks Hyperlinks
        {
            get
            {
                var hyperlinks = new XLHyperlinks();
                var hls = from hl in Worksheet.Hyperlinks
                          where RangeAddress.Contains(hl.Cell.Address)
                          select hl;
                hls.ForEach(hyperlinks.Add);
                return hyperlinks;
            }
        }

        public object Value
        {
            set { Cells().ForEach(c => c.Value = value); }
        }

        public XLDataType DataType
        {
            set { Cells().ForEach(c => c.DataType = value); }
        }

        #endregion IXLRangeBase Members

        #region IXLStylized Members

        public override IXLRanges RangesUsed
        {
            get
            {
                var retVal = new XLRanges { AsRange() };
                return retVal;
            }
        }

        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                foreach (var cell in Cells().OfType<XLCell>())
                {
                    yield return cell;
                }
            }
        }

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                foreach (var cell in Cells())
                {
                    yield return cell.Style;
                }
            }
        }

        #endregion IXLStylized Members

        #endregion Public properties

        #region IXLRangeBase Members

        IXLCell IXLRangeBase.FirstCell()
        {
            return FirstCell();
        }

        IXLCell IXLRangeBase.LastCell()
        {
            return LastCell();
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLCell IXLRangeBase.FirstCellUsed()
        {
            return FirstCellUsed(false);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLCell IXLRangeBase.FirstCellUsed(bool includeFormats)
        {
            return FirstCellUsed(includeFormats);
        }

        IXLCell IXLRangeBase.FirstCellUsed(XLCellsUsedOptions options)
        {
            return FirstCellUsed(options, null);
        }

        IXLCell IXLRangeBase.FirstCellUsed(Func<IXLCell, bool> predicate)
        {
            return FirstCellUsed(predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLCell IXLRangeBase.FirstCellUsed(bool includeFormats, Func<IXLCell, bool> predicate)
        {
            return FirstCellUsed(includeFormats, predicate);
        }

        IXLCell IXLRangeBase.FirstCellUsed(XLCellsUsedOptions options, Func<IXLCell, bool> predicate)
        {
            return FirstCellUsed(options, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLCell IXLRangeBase.LastCellUsed()
        {
            return LastCellUsed(false);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLCell IXLRangeBase.LastCellUsed(bool includeFormats)
        {
            return LastCellUsed(includeFormats);
        }

        IXLCell IXLRangeBase.LastCellUsed(XLCellsUsedOptions options)
        {
            return LastCellUsed(options, null);
        }

        IXLCell IXLRangeBase.LastCellUsed(Func<IXLCell, bool> predicate)
        {
            return LastCellUsed(predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLCell IXLRangeBase.LastCellUsed(bool includeFormats, Func<IXLCell, bool> predicate)
        {
            return LastCellUsed(includeFormats, predicate);
        }

        IXLCell IXLRangeBase.LastCellUsed(XLCellsUsedOptions options, Func<IXLCell, bool> predicate)
        {
            return LastCellUsed(options, predicate);
        }

        public virtual IXLCells Cells()
        {
            return Cells(false);
        }

        public virtual IXLCells Cells(bool usedCellsOnly)
        {
            return Cells(usedCellsOnly, XLCellsUsedOptions.AllContents);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public IXLCells Cells(bool usedCellsOnly, bool includeFormats)
        {
            return Cells(usedCellsOnly, includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents
            );
        }

        public IXLCells Cells(bool usedCellsOnly, XLCellsUsedOptions options)
        {
            var cells = new XLCells(usedCellsOnly, options) { RangeAddress };
            return cells;
        }

        public virtual IXLCells Cells(string cells)
        {
            return Ranges(cells).Cells();
        }

        public IXLCells Cells(Func<IXLCell, bool> predicate)
        {
            var cells = new XLCells(false, XLCellsUsedOptions.AllContents, predicate) { RangeAddress };
            return cells;
        }

        public IXLCells CellsUsed()
        {
            return Cells(true);
        }

        /// <summary>
        /// Return the collection of cell values not initializing empty cells.
        /// </summary>
        public IEnumerable CellValues()
        {
            for (var ro = RangeAddress.FirstAddress.RowNumber; ro <= RangeAddress.LastAddress.RowNumber; ro++)
            {
                for (var co = RangeAddress.FirstAddress.ColumnNumber; co <= RangeAddress.LastAddress.ColumnNumber; co++)
                {
                    yield return Worksheet.GetCellValue(ro, co);
                }
            }
        }

        public IXLRange Merge()
        {
            return Merge(true);
        }

        public IXLRange Merge(bool checkIntersect)
        {
            if (RangeAddress.FirstAddress.Equals(RangeAddress.LastAddress))
            {
                return Worksheet.Range(RangeAddress);
            }

            var asRange = AsRange();

            if (checkIntersect)
            {
                var intersectedMergedRanges =
                    Worksheet.Internals.MergedRanges.GetIntersectedRanges(RangeAddress).ToList();
                foreach (var intersectedMergedRange in intersectedMergedRanges)
                {
                    Worksheet.Internals.MergedRanges.Remove(intersectedMergedRange);
                }

                var firstCell = FirstCell();
                var firstCellStyleKey = (firstCell.Style as XLStyle).Key;
                var firstCellStyle = firstCell.Style;
                var defaultStyleKey = XLStyle.Default.Key;
                var cellsUsed =
                    CellsUsed(XLCellsUsedOptions.All & ~XLCellsUsedOptions.MergedRanges, c => c != firstCell).ToList();
                cellsUsed.ForEach(c => c.Clear(XLClearOptions.All
                                               & ~XLClearOptions.MergedRanges
                                               & ~XLClearOptions.NormalFormats));

                if (firstCellStyleKey.Alignment != defaultStyleKey.Alignment)
                {
                    asRange.Style.Alignment = firstCellStyle.Alignment;
                }
                else
                {
                    cellsUsed.ForEach(c => c.Style.Alignment = firstCellStyle.Alignment);
                }

                if (firstCellStyleKey.Fill != defaultStyleKey.Fill)
                {
                    asRange.Style.Fill = firstCellStyle.Fill;
                }
                else
                {
                    cellsUsed.ForEach(c => c.Style.Fill = firstCellStyle.Fill);
                }

                if (firstCellStyleKey.Font != defaultStyleKey.Font)
                {
                    asRange.Style.Font = firstCellStyle.Font;
                }
                else
                {
                    cellsUsed.ForEach(c => c.Style.Font = firstCellStyle.Font);
                }

                if (firstCellStyleKey.IncludeQuotePrefix != defaultStyleKey.IncludeQuotePrefix)
                {
                    asRange.Style.IncludeQuotePrefix = firstCellStyle.IncludeQuotePrefix;
                }
                else
                {
                    cellsUsed.ForEach(c => c.Style.IncludeQuotePrefix = firstCellStyle.IncludeQuotePrefix);
                }

                if (firstCellStyleKey.NumberFormat != defaultStyleKey.NumberFormat)
                {
                    asRange.Style.NumberFormat = firstCellStyle.NumberFormat;
                }
                else
                {
                    cellsUsed.ForEach(c => c.Style.NumberFormat = firstCellStyle.NumberFormat);
                }

                if (firstCellStyleKey.Protection != defaultStyleKey.Protection)
                {
                    asRange.Style.Protection = firstCellStyle.Protection;
                }
                else
                {
                    cellsUsed.ForEach(c => c.Style.Protection = firstCellStyle.Protection);
                }

                if (cellsUsed.Any(c => (c.Style as XLStyle).Key.Border != defaultStyleKey.Border))
                {
                    asRange.Style.Border.SetInsideBorder(XLBorderStyleValues.None);
                }
            }

            Worksheet.Internals.MergedRanges.Add(asRange);
            return asRange;
        }

        public IXLRange Unmerge()
        {
            var tAddress = RangeAddress.ToString();
            var asRange = AsRange();
            if (Worksheet.Internals.MergedRanges.Select(m => m.RangeAddress.ToString()).Any(mAddress => mAddress == tAddress))
            {
                Worksheet.Internals.MergedRanges.Remove(asRange);
            }

            return asRange;
        }

        public IXLRangeBase Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            var cellClearOptions = clearOptions
                    & ~XLClearOptions.ConditionalFormats
                    & ~XLClearOptions.DataValidation
                    & ~XLClearOptions.MergedRanges;
            var cellUsedOptions = cellClearOptions.ToCellsUsedOptions();
            foreach (var cell in CellsUsed(cellUsedOptions))
            {
                // We'll clear the conditional formatting, data validations
                // and merged ranges later down.
                (cell as XLCell).Clear(cellClearOptions, true);
            }

            if (clearOptions.HasFlag(XLClearOptions.ConditionalFormats))
            {
                RemoveConditionalFormatting();
            }

            if (clearOptions.HasFlag(XLClearOptions.DataValidation))
            {
                var validation = CreateDataValidation();
                Worksheet.DataValidations.Delete(validation);
            }

            if (clearOptions.HasFlag(XLClearOptions.MergedRanges))
            {
                ClearMerged();
            }

            if (clearOptions.HasFlag(XLClearOptions.Sparklines))
            {
                RemoveSparklines();
            }

            if (clearOptions == XLClearOptions.All)
            {
                Worksheet.Internals.CellsCollection.RemoveAll(
                    RangeAddress.FirstAddress.RowNumber,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber
                );
            }
            return this;
        }

        public IXLRangeBase Relative(IXLRangeBase sourceBaseRange, IXLRangeBase targetBaseRange)
        {
            var xlSourceBaseRangeAddress = (XLRangeAddress)sourceBaseRange.RangeAddress;
            var xlTargetBaseRangeAddress = (XLRangeAddress)targetBaseRange.RangeAddress;
            var xlRangeAddress = RangeAddress.Relative(in xlSourceBaseRangeAddress, in xlTargetBaseRangeAddress);

            return ((XLRangeBase)targetBaseRange).Range(in xlRangeAddress);
        }

        internal void RemoveConditionalFormatting()
        {
            var mf = RangeAddress.FirstAddress;
            var ml = RangeAddress.LastAddress;
            foreach (var format in Worksheet.ConditionalFormats.Where(x => x.Ranges.GetIntersectedRanges(RangeAddress).Any()).ToList())
            {
                var cfRanges = format.Ranges.ToList();
                format.Ranges.RemoveAll();

                foreach (var cfRange in cfRanges)
                {
                    if (!cfRange.Intersects(this))
                    {
                        format.Ranges.Add(cfRange);
                        continue;
                    }

                    var f = cfRange.RangeAddress.FirstAddress;
                    var l = cfRange.RangeAddress.LastAddress;
                    bool byWidth = false, byHeight = false;
                    XLRange rng1 = null, rng2 = null;
                    if (mf.ColumnNumber <= f.ColumnNumber && ml.ColumnNumber >= l.ColumnNumber)
                    {
                        if (mf.RowNumber.Between(f.RowNumber, l.RowNumber) || ml.RowNumber.Between(f.RowNumber, l.RowNumber))
                        {
                            if (mf.RowNumber > f.RowNumber)
                            {
                                rng1 = Worksheet.Range(f.RowNumber, f.ColumnNumber, mf.RowNumber - 1, l.ColumnNumber);
                            }

                            if (ml.RowNumber < l.RowNumber)
                            {
                                rng2 = Worksheet.Range(ml.RowNumber + 1, f.ColumnNumber, l.RowNumber, l.ColumnNumber);
                            }
                        }
                        byWidth = true;
                    }

                    if (mf.RowNumber <= f.RowNumber && ml.RowNumber >= l.RowNumber)
                    {
                        if (mf.ColumnNumber.Between(f.ColumnNumber, l.ColumnNumber) || ml.ColumnNumber.Between(f.ColumnNumber, l.ColumnNumber))
                        {
                            if (mf.ColumnNumber > f.ColumnNumber)
                            {
                                rng1 = Worksheet.Range(f.RowNumber, f.ColumnNumber, l.RowNumber, mf.ColumnNumber - 1);
                            }

                            if (ml.ColumnNumber < l.ColumnNumber)
                            {
                                rng2 = Worksheet.Range(f.RowNumber, ml.ColumnNumber + 1, l.RowNumber, l.ColumnNumber);
                            }
                        }
                        byHeight = true;
                    }

                    if (rng1 != null)
                    {
                        format.Ranges.Add(rng1);
                    }
                    if (rng2 != null)
                    {
                        //TODO: reflect the formula for a new range
                        format.Ranges.Add(rng2);
                    }

                    if (!byWidth && !byHeight)
                    {
                        format.Ranges.Add(cfRange); // Not split, preserve original
                    }
                }
                if (!format.Ranges.Any())
                {
                    Worksheet.ConditionalFormats.Remove(x => x == format);
                }
            }
        }

        internal void RemoveSparklines()
        {
            Worksheet.SparklineGroups.GetSparklines(this).ToList()
                .ForEach(sl => Worksheet.SparklineGroups.Remove(sl.Location));
        }

        public void DeleteComments()
        {
            Cells().DeleteComments();
        }

        public bool Contains(string rangeAddress)
        {
            var addressToUse = rangeAddress.Contains("!")
                                      ? rangeAddress.Substring(rangeAddress.IndexOf("!") + 1)
                                      : rangeAddress;

            XLAddress firstAddress;
            XLAddress lastAddress;
            if (addressToUse.Contains(':'))
            {
                var arrRange = addressToUse.Split(':');
                firstAddress = XLAddress.Create(Worksheet, arrRange[0]);
                lastAddress = XLAddress.Create(Worksheet, arrRange[1]);
            }
            else
            {
                firstAddress = XLAddress.Create(Worksheet, addressToUse);
                lastAddress = XLAddress.Create(Worksheet, addressToUse);
            }
            return Contains(firstAddress, lastAddress);
        }

        public bool Contains(IXLRangeBase range)
        {
            return Contains((XLAddress)range.RangeAddress.FirstAddress, (XLAddress)range.RangeAddress.LastAddress);
        }

        public bool Intersects(string rangeAddress)
        {
            return Intersects(Worksheet.Range(rangeAddress));
        }

        public bool Intersects(IXLRangeBase range)
        {
            if (!range.RangeAddress.IsValid || !RangeAddress.IsValid)
            {
                return false;
            }

            var ma = range.RangeAddress;
            var ra = RangeAddress;
            return ra.Intersects(ma);
        }

        IXLRange IXLRangeBase.AsRange()
        {
            return AsRange();
        }

        public virtual XLRange AsRange()
        {
            return Worksheet.Range(RangeAddress);
        }

        public IXLRange AddToNamed(string rangeName)
        {
            return AddToNamed(rangeName, XLScope.Workbook);
        }

        public IXLRange AddToNamed(string rangeName, XLScope scope)
        {
            return AddToNamed(rangeName, scope, null);
        }

        public IXLRange AddToNamed(string rangeName, XLScope scope, string comment)
        {
            var namedRanges = scope == XLScope.Workbook
                                  ? Worksheet.Workbook.NamedRanges
                                  : Worksheet.NamedRanges;

            if (namedRanges.TryGetValue(rangeName, out var namedRange))
            {
                namedRange.Add(Worksheet.Workbook, RangeAddress.ToStringFixed(XLReferenceStyle.A1, true));
            }
            else
            {
                namedRanges.Add(rangeName, RangeAddress.ToStringFixed(XLReferenceStyle.A1, true), comment);
            }

            return AsRange();
        }

        public IXLRangeBase SetValue<T>(T value)
        {
            Cells().ForEach(c => c.SetValue(value));
            return this;
        }

        public bool IsMerged()
        {
            return Cells().Any(c => c.IsMerged());
        }

        public virtual bool IsEmpty()
        {
            return !CellsUsed().Any() || CellsUsed().Any(c => c.IsEmpty());
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public virtual bool IsEmpty(bool includeFormats)
        {
            return IsEmpty(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents);
        }

        public virtual bool IsEmpty(XLCellsUsedOptions options)
        {
            return CellsUsed(options).Cast<XLCell>().All(c => c.IsEmpty(options));
        }

        public virtual bool IsEntireRow()
        {
            return RangeAddress.IsEntireRow();
        }

        public virtual bool IsEntireColumn()
        {
            return RangeAddress.IsEntireColumn();
        }

        public bool IsEntireSheet()
        {
            return RangeAddress.IsEntireSheet();
        }

        #endregion IXLRangeBase Members

        public IXLCells Search(string searchText, CompareOptions compareOptions = CompareOptions.Ordinal, bool searchFormulae = false)
        {
            var culture = CultureInfo.CurrentCulture;
            return CellsUsed(XLCellsUsedOptions.AllContents, c =>
            {
                try
                {
                    if (searchFormulae)
                    {
                        return c.HasFormula
                               && culture.CompareInfo.IndexOf(c.FormulaA1, searchText, compareOptions) >= 0
                               || culture.CompareInfo.IndexOf(c.Value.ToString(), searchText, compareOptions) >= 0;
                    }
                    else
                    {
                        return culture.CompareInfo.IndexOf(c.GetFormattedString(), searchText, compareOptions) >= 0;
                    }
                }
                catch
                {
                    return false;
                }
            });
        }

        internal XLCell FirstCell()
        {
            return Cell(1, 1);
        }

        internal XLCell LastCell()
        {
            return Cell(RowCount(), ColumnCount());
        }

        internal XLCell FirstCellUsed()
        {
            return FirstCellUsed(XLCellsUsedOptions.AllContents, predicate: null);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        internal XLCell FirstCellUsed(bool includeFormats)
        {
            return FirstCellUsed(includeFormats, null);
        }

        internal XLCell FirstCellUsed(Func<IXLCell, bool> predicate)
        {
            return FirstCellUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        internal XLCell FirstCellUsed(bool includeFormats, Func<IXLCell, bool> predicate)
        {
            return FirstCellUsed(includeFormats
                    ? XLCellsUsedOptions.All
                    : XLCellsUsedOptions.AllContents,
                predicate);
        }

        internal XLCell FirstCellUsed(XLCellsUsedOptions options, Func<IXLCell, bool> predicate = null)
        {
            var cellsUsed = CellsUsedInternal(options, r => r.FirstCell(), predicate).ToList();

            if (!cellsUsed.Any())
            {
                return null;
            }

            var firstRow = cellsUsed.Min(c => c.Address.RowNumber);
            var firstColumn = cellsUsed.Min(c => c.Address.ColumnNumber);

            if (firstRow < RangeAddress.FirstAddress.RowNumber)
            {
                firstRow = RangeAddress.FirstAddress.RowNumber;
            }

            if (firstColumn < RangeAddress.FirstAddress.ColumnNumber)
            {
                firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            }

            return Worksheet.Cell(firstRow, firstColumn);
        }

        internal XLCell LastCellUsed()
        {
            return LastCellUsed(XLCellsUsedOptions.AllContents, predicate: null);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        internal XLCell LastCellUsed(bool includeFormats)
        {
            return LastCellUsed(includeFormats, null);
        }

        internal XLCell LastCellUsed(Func<IXLCell, bool> predicate)
        {
            return LastCellUsed(XLCellsUsedOptions.AllContents, predicate);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        internal XLCell LastCellUsed(bool includeFormats, Func<IXLCell, bool> predicate)
        {
            return LastCellUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents,
                predicate);
        }

        internal XLCell LastCellUsed(XLCellsUsedOptions options, Func<IXLCell, bool> predicate = null)
        {
            var cellsUsed = CellsUsedInternal(options, r => r.LastCell(), predicate).ToList();

            if (!cellsUsed.Any())
            {
                return null;
            }

            var lastRow = cellsUsed.Max(c => c.Address.RowNumber);
            var lastColumn = cellsUsed.Max(c => c.Address.ColumnNumber);

            if (lastRow > RangeAddress.LastAddress.RowNumber)
            {
                lastRow = RangeAddress.LastAddress.RowNumber;
            }

            if (lastColumn > RangeAddress.LastAddress.ColumnNumber)
            {
                lastColumn = RangeAddress.LastAddress.ColumnNumber;
            }

            return Worksheet.Cell(lastRow, lastColumn);
        }

        public XLCell Cell(int row, int column)
        {
            return Cell(new XLAddress(Worksheet, row, column, false, false));
        }

        public virtual XLCell Cell(string cellAddressInRange)
        {
            if (XLHelper.IsValidA1Address(cellAddressInRange))
            {
                return Cell(XLAddress.Create(Worksheet, cellAddressInRange));
            }

            if (Worksheet.NamedRanges.TryGetValue(cellAddressInRange, out var namedRange))
            {
                return namedRange.Ranges.First().FirstCell().CastTo<XLCell>();
            }

            return null;
        }

        public XLCell Cell(int row, string column)
        {
            return Cell(new XLAddress(Worksheet, row, column, false, false));
        }

        public XLCell Cell(IXLAddress cellAddressInRange)
        {
            return Cell(cellAddressInRange.RowNumber, cellAddressInRange.ColumnNumber);
        }

        public XLCell Cell(in XLAddress cellAddressInRange)
        {
            var absRow = cellAddressInRange.RowNumber + RangeAddress.FirstAddress.RowNumber - 1;
            var absColumn = cellAddressInRange.ColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1;

            if (absRow <= 0 || absRow > XLHelper.MaxRowNumber)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(cellAddressInRange),
                    string.Format("Row number must be between 1 and {0}", XLHelper.MaxRowNumber)
                );
            }

            if (absColumn <= 0 || absColumn > XLHelper.MaxColumnNumber)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(cellAddressInRange),
                    string.Format("Column number must be between 1 and {0}", XLHelper.MaxColumnNumber)
                );
            }

            var cell = Worksheet.Internals.CellsCollection.GetCell(absRow,
                                                                   absColumn);

            if (cell != null)
            {
                return cell;
            }

            var styleValue = StyleValue;

            if (styleValue == Worksheet.StyleValue)
            {
                if (Worksheet.Internals.RowsCollection.TryGetValue(absRow, out var row)
                    && row.StyleValue != Worksheet.StyleValue)
                {
                    styleValue = row.StyleValue;
                }
                else if (Worksheet.Internals.ColumnsCollection.TryGetValue(absColumn, out var column)
                    && column.StyleValue != Worksheet.StyleValue)
                {
                    styleValue = column.StyleValue;
                }
            }
            var absoluteAddress = new XLAddress(Worksheet,
                                 absRow,
                                 absColumn,
                                 cellAddressInRange.FixedRow,
                                 cellAddressInRange.FixedColumn);

            // If the default style for this range base is empty, but the worksheet
            // has a default style, use the worksheet's default style
            var newCell = new XLCell(Worksheet, absoluteAddress, styleValue);

            Worksheet.Internals.CellsCollection.Add(absRow, absColumn, newCell);
            return newCell;
        }

        public int RowCount()
        {
            return RangeAddress.LastAddress.RowNumber - RangeAddress.FirstAddress.RowNumber + 1;
        }

        public int RowCount(XLCellsUsedOptions cellsUsedOptions)
        {
            var lcu = LastCellUsed(cellsUsedOptions);
            if (lcu == null)
            {
                return 0;
            }

            var fcu = FirstCellUsed(cellsUsedOptions);
            if (fcu == null)
            {
                return 0;
            }

            return lcu.Address.RowNumber - fcu.Address.RowNumber + 1;
        }

        public int RowNumber()
        {
            return RangeAddress.FirstAddress.RowNumber;
        }

        public int ColumnCount()
        {
            return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
        }

        public int ColumnCount(XLCellsUsedOptions cellsUsedOptions)
        {
            var lcu = LastCellUsed(cellsUsedOptions);
            if (lcu == null)
            {
                return 0;
            }

            var fcu = FirstCellUsed(cellsUsedOptions);
            if (fcu == null)
            {
                return 0;
            }

            return lcu.Address.ColumnNumber - fcu.Address.ColumnNumber + 1;
        }

        public int ColumnNumber()
        {
            return RangeAddress.FirstAddress.ColumnNumber;
        }

        public string ColumnLetter()
        {
            return RangeAddress.FirstAddress.ColumnLetter;
        }

        public virtual XLRange Range(string rangeAddressStr)
        {
            var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressStr);
            return Range(rangeAddress);
        }

        internal abstract void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted);

        internal abstract void WorksheetRangeShiftedRows(XLRange range, int rowsShifted);

        public abstract XLRangeType RangeType { get; }

        public XLRange Range(IXLCell firstCell, IXLCell lastCell)
        {
            var newFirstCellAddress = (XLAddress)firstCell.Address;
            var newLastCellAddress = (XLAddress)lastCell.Address;

            return GetRange(newFirstCellAddress, newLastCellAddress);
        }

        private XLRange GetRange(XLAddress newFirstCellAddress, XLAddress newLastCellAddress)
        {
            if (!Worksheet.Equals(newFirstCellAddress.Worksheet))
            {
                throw new ArgumentException("The address refers to a different worksheet.", nameof(newFirstCellAddress));
            }

            if (!Worksheet.Equals(newLastCellAddress.Worksheet))
            {
                throw new ArgumentException("The address refers to a different worksheet.", nameof(newLastCellAddress));
            }

            var newRangeAddress = new XLRangeAddress(newFirstCellAddress, newLastCellAddress);
            var xlRangeParameters = new XLRangeParameters(newRangeAddress, Style);
            if (
                newFirstCellAddress.RowNumber < RangeAddress.FirstAddress.RowNumber
                || newFirstCellAddress.RowNumber > RangeAddress.LastAddress.RowNumber
                || newLastCellAddress.RowNumber > RangeAddress.LastAddress.RowNumber
                || newFirstCellAddress.ColumnNumber < RangeAddress.FirstAddress.ColumnNumber
                || newFirstCellAddress.ColumnNumber > RangeAddress.LastAddress.ColumnNumber
                || newLastCellAddress.ColumnNumber > RangeAddress.LastAddress.ColumnNumber
            )
            {
                throw new ArgumentOutOfRangeException(string.Format(
                    "The cells {0} and {1} are outside the range '{2}'.",
                    newFirstCellAddress,
                    newLastCellAddress,
                    ToString()));
            }

            if (newFirstCellAddress.Worksheet != null)
            {
                return newFirstCellAddress.Worksheet.GetOrCreateRange(xlRangeParameters);
            }
            else if (Worksheet != null)
            {
                return Worksheet.GetOrCreateRange(xlRangeParameters);
            }
            else
            {
                return new XLRange(xlRangeParameters);
            }
        }

        public XLRange Range(string firstCellAddress, string lastCellAddress)
        {
            var rangeAddress = new XLRangeAddress(XLAddress.Create(Worksheet, firstCellAddress),
                                                  XLAddress.Create(Worksheet, lastCellAddress));
            return Range(rangeAddress);
        }

        public XLRange Range(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
        {
            var rangeAddress = new XLRangeAddress
            (
                new XLAddress
                (
                    Worksheet,
                    firstCellRow + RangeAddress.FirstAddress.RowNumber - 1,
                    firstCellColumn + RangeAddress.FirstAddress.ColumnNumber - 1,
                    fixedRow: false,
                    fixedColumn: false
                ),
                new XLAddress
                (
                    Worksheet,
                    lastCellRow + RangeAddress.FirstAddress.RowNumber - 1,
                    lastCellColumn + RangeAddress.FirstAddress.ColumnNumber - 1,
                    fixedRow: false,
                    fixedColumn: false
                )
            );
            return Range(rangeAddress);
        }

        public XLRange Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
        {
            var rangeAddress = new XLRangeAddress((XLAddress)firstCellAddress, (XLAddress)lastCellAddress);
            return Range(rangeAddress);
        }

        public XLRange Range(IXLRangeAddress rangeAddress)
        {
            var xlRangeAddress = (XLRangeAddress)rangeAddress;
            return Range(in xlRangeAddress);
        }

        internal XLRange Range(in XLRangeAddress rangeAddress)
        {
            var ws = rangeAddress.FirstAddress.Worksheet ??
                     rangeAddress.LastAddress.Worksheet ??
                     Worksheet;

            var newFirstCellAddress = new XLAddress(ws,
                                 rangeAddress.FirstAddress.RowNumber,
                                 rangeAddress.FirstAddress.ColumnNumber,
                                 rangeAddress.FirstAddress.FixedRow,
                                 rangeAddress.FirstAddress.FixedColumn);

            var newLastCellAddress = new XLAddress(ws,
                                rangeAddress.LastAddress.RowNumber,
                                rangeAddress.LastAddress.ColumnNumber,
                                rangeAddress.LastAddress.FixedRow,
                                rangeAddress.LastAddress.FixedColumn);

            return GetRange(newFirstCellAddress, newLastCellAddress);
        }

        public virtual IXLRanges Ranges(string ranges)
        {
            var retVal = new XLRanges();
            var rangePairs = ranges.Split(',');
            foreach (var pair in rangePairs)
            {
                retVal.Add(Range(pair.Trim()));
            }

            return retVal;
        }

        public IXLRanges Ranges(params string[] ranges)
        {
            var retVal = new XLRanges();
            foreach (var pair in ranges)
            {
                retVal.Add(Range(pair));
            }

            return retVal;
        }

        protected string FixColumnAddress(string address)
        {
            if (int.TryParse(address, out var rowNumber))
            {
                return RangeAddress.FirstAddress.ColumnLetter + (rowNumber + RangeAddress.FirstAddress.RowNumber - 1).ToInvariantString();
            }

            return address;
        }

        protected string FixRowAddress(string address)
        {
            if (int.TryParse(address, out var columnNumber))
            {
                return XLHelper.GetColumnLetterFromNumber(columnNumber + RangeAddress.FirstAddress.ColumnNumber - 1) + RangeAddress.FirstAddress.RowNumber.ToInvariantString();
            }

            return address;
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public IXLCells CellsUsed(bool includeFormats)
        {
            return CellsUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents);
        }

        public IXLCells CellsUsed(XLCellsUsedOptions options)
        {
            var cells = new XLCells(true, options) { RangeAddress };
            return cells;
        }

        public IXLCells CellsUsed(Func<IXLCell, bool> predicate)
        {
            var cells = new XLCells(true, XLCellsUsedOptions.AllContents, predicate) { RangeAddress };
            return cells;
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public IXLCells CellsUsed(bool includeFormats, Func<IXLCell, bool> predicate)
        {
            return CellsUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents,
                predicate);
        }

        public IXLCells CellsUsed(XLCellsUsedOptions options, Func<IXLCell, bool> predicate)
        {
            var cells = new XLCells(true, options, predicate) { RangeAddress };
            return cells;
        }

        public IXLRangeColumns InsertColumnsAfter(int numberOfColumns)
        {
            return InsertColumnsAfter(numberOfColumns, true);
        }

        public IXLRangeColumns InsertColumnsAfter(int numberOfColumns, bool expandRange)
        {
            var retVal = InsertColumnsAfter(false, numberOfColumns);
            // Adjust the range
            if (expandRange)
            {
                RangeAddress = new XLRangeAddress(
                    new XLAddress(Worksheet,
                                  RangeAddress.FirstAddress.RowNumber,
                                  RangeAddress.FirstAddress.ColumnNumber,
                                  RangeAddress.FirstAddress.FixedRow,
                                  RangeAddress.FirstAddress.FixedColumn),
                    new XLAddress(Worksheet,
                                  RangeAddress.LastAddress.RowNumber,
                                  RangeAddress.LastAddress.ColumnNumber + numberOfColumns,
                                  RangeAddress.LastAddress.FixedRow,
                                  RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }

        public IXLRangeColumns InsertColumnsAfter(bool onlyUsedCells, int numberOfColumns, bool formatFromLeft = true)
        {
            return InsertColumnsAfterInternal(onlyUsedCells, numberOfColumns, formatFromLeft);
        }

        public void InsertColumnsAfterVoid(bool onlyUsedCells, int numberOfColumns, bool formatFromLeft = true)
        {
            InsertColumnsAfterInternal(onlyUsedCells, numberOfColumns, formatFromLeft, nullReturn: true);
        }

        private IXLRangeColumns InsertColumnsAfterInternal(bool onlyUsedCells, int numberOfColumns, bool formatFromLeft = true, bool nullReturn = false)
        {
            var columnCount = ColumnCount();
            var firstColumn = RangeAddress.FirstAddress.ColumnNumber + columnCount;
            if (firstColumn > XLHelper.MaxColumnNumber)
            {
                firstColumn = XLHelper.MaxColumnNumber;
            }

            var lastColumn = firstColumn + ColumnCount() - 1;
            if (lastColumn > XLHelper.MaxColumnNumber)
            {
                lastColumn = XLHelper.MaxColumnNumber;
            }

            var firstRow = RangeAddress.FirstAddress.RowNumber;
            var lastRow = firstRow + RowCount() - 1;
            if (lastRow > XLHelper.MaxRowNumber)
            {
                lastRow = XLHelper.MaxRowNumber;
            }

            var newRange = Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
            return newRange.InsertColumnsBeforeInternal(onlyUsedCells, numberOfColumns, formatFromLeft, nullReturn);
        }

        public IXLRangeColumns InsertColumnsBefore(int numberOfColumns)
        {
            return InsertColumnsBefore(numberOfColumns, false);
        }

        public IXLRangeColumns InsertColumnsBefore(int numberOfColumns, bool expandRange)
        {
            var retVal = InsertColumnsBefore(false, numberOfColumns);
            // Adjust the range
            if (expandRange)
            {
                RangeAddress = new XLRangeAddress(
                    new XLAddress(Worksheet,
                                  RangeAddress.FirstAddress.RowNumber,
                                  RangeAddress.FirstAddress.ColumnNumber - numberOfColumns,
                                  RangeAddress.FirstAddress.FixedRow,
                                  RangeAddress.FirstAddress.FixedColumn),
                    new XLAddress(Worksheet,
                                  RangeAddress.LastAddress.RowNumber,
                                  RangeAddress.LastAddress.ColumnNumber,
                                  RangeAddress.LastAddress.FixedRow,
                                  RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }

        public IXLRangeColumns InsertColumnsBefore(bool onlyUsedCells, int numberOfColumns, bool formatFromLeft = true)
        {
            return InsertColumnsBeforeInternal(onlyUsedCells, numberOfColumns, formatFromLeft);
        }

        public void InsertColumnsBeforeVoid(bool onlyUsedCells, int numberOfColumns, bool formatFromLeft = true)
        {
            InsertColumnsBeforeInternal(onlyUsedCells, numberOfColumns, formatFromLeft, nullReturn: true);
        }

        private IXLRangeColumns InsertColumnsBeforeInternal(bool onlyUsedCells, int numberOfColumns, bool formatFromLeft = true, bool nullReturn = false)
        {
            if (numberOfColumns <= 0 || numberOfColumns > XLHelper.MaxColumnNumber)
            {
                throw new ArgumentOutOfRangeException(nameof(numberOfColumns),
                    $"Number of columns to insert must be a positive number no more than {XLHelper.MaxColumnNumber}");
            }

            foreach (var ws in Worksheet.Workbook.WorksheetsInternal)
            {
                foreach (var cell in ws.Internals.CellsCollection.GetCells(c => !string.IsNullOrWhiteSpace(c.FormulaA1)))
                {
                    cell.ShiftFormulaColumns(AsRange(), numberOfColumns);
                }
            }

            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
            var cellsToDelete = new List<IXLAddress>();
            var firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            var firstRow = RangeAddress.FirstAddress.RowNumber;
            var lastRow = RangeAddress.FirstAddress.RowNumber + RowCount() - 1;

            if (!onlyUsedCells)
            {
                var lastColumn = Worksheet.Internals.CellsCollection.MaxColumnUsed;
                if (lastColumn > 0)
                {
                    for (var co = lastColumn; co >= firstColumn; co--)
                    {
                        var newColumn = co + numberOfColumns;
                        for (var ro = lastRow; ro >= firstRow; ro--)
                        {
                            var oldCell = Worksheet.Internals.CellsCollection.GetCell(ro, co);
                            if (oldCell == null)
                            {
                                continue;
                            }

                            var oldKey = new XLAddress(Worksheet, ro, co, false, false);
                            var newKey = new XLAddress(Worksheet, ro, newColumn, false, false);

                            oldCell.Address = newKey;
                            if (newKey.IsValid)
                            {
                                cellsToInsert.Add(newKey, oldCell);
                            }

                            cellsToDelete.Add(oldKey);
                        }

                        if (IsEntireColumn())
                        {
                            Worksheet.Column(newColumn).Width = Worksheet.Column(co).Width;
                        }
                    }
                }
            }
            else
            {
                foreach (
                    var c in
                        Worksheet.Internals.CellsCollection.GetCells(firstRow, firstColumn, lastRow,
                                                                     Worksheet.Internals.CellsCollection.MaxColumnUsed))
                {
                    var newColumn = c.Address.ColumnNumber + numberOfColumns;
                    var newKey = new XLAddress(Worksheet, c.Address.RowNumber, newColumn, false, false);

                    cellsToDelete.Add(c.Address);
                    c.Address = newKey;
                    if (newKey.IsValid)
                    {
                        cellsToInsert.Add(newKey, c);
                    }
                }
            }

            cellsToDelete.ForEach(c => Worksheet.Internals.CellsCollection.Remove(c.RowNumber, c.ColumnNumber));
            cellsToInsert.ForEach(
                c => Worksheet.Internals.CellsCollection.Add(c.Key.RowNumber, c.Key.ColumnNumber, c.Value));

            var firstRowReturn = RangeAddress.FirstAddress.RowNumber;
            var lastRowReturn = RangeAddress.LastAddress.RowNumber;
            var firstColumnReturn = RangeAddress.FirstAddress.ColumnNumber;
            var lastColumnReturn = RangeAddress.FirstAddress.ColumnNumber + numberOfColumns - 1;

            Worksheet.NotifyRangeShiftedColumns(AsRange(), numberOfColumns);

            var rangeToReturn = Worksheet.Range(firstRowReturn, firstColumnReturn, lastRowReturn, lastColumnReturn);

            // We deliberately ignore conditional formats and data validation here. Their shifting is handled elsewhere
            var contentFlags = XLCellsUsedOptions.All
                & ~XLCellsUsedOptions.ConditionalFormats
                & ~XLCellsUsedOptions.DataValidation;

            if (formatFromLeft && rangeToReturn.RangeAddress.FirstAddress.ColumnNumber > 1)
            {
                var firstColumnUsed = rangeToReturn.FirstColumn();
                var model = firstColumnUsed.ColumnLeft();
                var modelFirstRow = (model as IXLRangeBase).FirstCellUsed(contentFlags);
                var modelLastRow = (model as IXLRangeBase).LastCellUsed(contentFlags);
                if (modelLastRow != null)
                {
                    var firstRoReturned = modelFirstRow.Address.RowNumber
                                            - model.RangeAddress.FirstAddress.RowNumber + 1;
                    var lastRoReturned = modelLastRow.Address.RowNumber
                                           - model.RangeAddress.FirstAddress.RowNumber + 1;
                    for (var ro = firstRoReturned; ro <= lastRoReturned; ro++)
                    {
                        rangeToReturn.Row(ro).Style = model.Cell(ro).Style;
                    }
                }
            }
            else
            {
                var lastRoUsed = rangeToReturn.LastRowUsed(contentFlags);
                if (lastRoUsed != null)
                {
                    var lastRoReturned = lastRoUsed.RowNumber();
                    for (var ro = 1; ro <= lastRoReturned; ro++)
                    {
                        var styleToUse =
                            Worksheet.Internals.RowsCollection.TryGetValue(ro, out var row)
                                ? row.Style
                                : Worksheet.Style;

                        rangeToReturn.Row(ro).Style = styleToUse;
                    }
                }
            }

            if (nullReturn)
            {
                return null;
            }

            return rangeToReturn.Columns();
        }

        public IXLRangeRows InsertRowsBelow(int numberOfRows)
        {
            return InsertRowsBelow(numberOfRows, true);
        }

        public IXLRangeRows InsertRowsBelow(int numberOfRows, bool expandRange)
        {
            var retVal = InsertRowsBelow(false, numberOfRows);
            // Adjust the range
            if (expandRange)
            {
                RangeAddress = new XLRangeAddress(
                    new XLAddress(Worksheet,
                                  RangeAddress.FirstAddress.RowNumber,
                                  RangeAddress.FirstAddress.ColumnNumber,
                                  RangeAddress.FirstAddress.FixedRow,
                                  RangeAddress.FirstAddress.FixedColumn),
                    new XLAddress(Worksheet,
                                  RangeAddress.LastAddress.RowNumber + numberOfRows,
                                  RangeAddress.LastAddress.ColumnNumber,
                                  RangeAddress.LastAddress.FixedRow,
                                  RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }

        public IXLRangeRows InsertRowsBelow(bool onlyUsedCells, int numberOfRows, bool formatFromAbove = true)
        {
            return InsertRowsBelowInternal(onlyUsedCells, numberOfRows, formatFromAbove, nullReturn: false);
        }

        public void InsertRowsBelowVoid(bool onlyUsedCells, int numberOfRows, bool formatFromAbove = true)
        {
            InsertRowsBelowInternal(onlyUsedCells, numberOfRows, formatFromAbove, nullReturn: true);
        }

        private IXLRangeRows InsertRowsBelowInternal(bool onlyUsedCells, int numberOfRows, bool formatFromAbove, bool nullReturn)
        {
            var rowCount = RowCount();
            var firstRow = RangeAddress.FirstAddress.RowNumber + rowCount;
            if (firstRow > XLHelper.MaxRowNumber)
            {
                firstRow = XLHelper.MaxRowNumber;
            }

            var lastRow = firstRow + RowCount() - 1;
            if (lastRow > XLHelper.MaxRowNumber)
            {
                lastRow = XLHelper.MaxRowNumber;
            }

            var firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            var lastColumn = firstColumn + ColumnCount() - 1;
            if (lastColumn > XLHelper.MaxColumnNumber)
            {
                lastColumn = XLHelper.MaxColumnNumber;
            }

            var newRange = Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
            return newRange.InsertRowsAboveInternal(onlyUsedCells, numberOfRows, formatFromAbove, nullReturn);
        }

        public IXLRangeRows InsertRowsAbove(int numberOfRows)
        {
            return InsertRowsAbove(numberOfRows, false);
        }

        public IXLRangeRows InsertRowsAbove(int numberOfRows, bool expandRange)
        {
            var retVal = InsertRowsAbove(false, numberOfRows);
            // Adjust the range
            if (expandRange)
            {
                RangeAddress = new XLRangeAddress(
                    new XLAddress(Worksheet,
                                  RangeAddress.FirstAddress.RowNumber - numberOfRows,
                                  RangeAddress.FirstAddress.ColumnNumber,
                                  RangeAddress.FirstAddress.FixedRow,
                                  RangeAddress.FirstAddress.FixedColumn),
                    new XLAddress(Worksheet,
                                  RangeAddress.LastAddress.RowNumber,
                                  RangeAddress.LastAddress.ColumnNumber,
                                  RangeAddress.LastAddress.FixedRow,
                                  RangeAddress.LastAddress.FixedColumn));
            }
            return retVal;
        }

        public void InsertRowsAboveVoid(bool onlyUsedCells, int numberOfRows, bool formatFromAbove = true)
        {
            InsertRowsAboveInternal(onlyUsedCells, numberOfRows, formatFromAbove, nullReturn: true);
        }

        public IXLRangeRows InsertRowsAbove(bool onlyUsedCells, int numberOfRows, bool formatFromAbove = true)
        {
            return InsertRowsAboveInternal(onlyUsedCells, numberOfRows, formatFromAbove, nullReturn: false);
        }

        private IXLRangeRows InsertRowsAboveInternal(bool onlyUsedCells, int numberOfRows, bool formatFromAbove, bool nullReturn)
        {
            if (numberOfRows <= 0 || numberOfRows > XLHelper.MaxRowNumber)
            {
                throw new ArgumentOutOfRangeException(nameof(numberOfRows),
                    $"Number of rows to insert must be a positive number no more than {XLHelper.MaxRowNumber}");
            }

            var asRange = AsRange();
            foreach (var ws in Worksheet.Workbook.WorksheetsInternal)
            {
                foreach (var cell in ws.Internals.CellsCollection.GetCells(c => !string.IsNullOrWhiteSpace(c.FormulaA1)))
                {
                    cell.ShiftFormulaRows(asRange, numberOfRows);
                }
            }

            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
            var cellsToDelete = new List<IXLAddress>();
            var firstRow = RangeAddress.FirstAddress.RowNumber;
            var firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            var lastColumn = Math.Min(
                RangeAddress.FirstAddress.ColumnNumber + ColumnCount() - 1,
                Worksheet.Internals.CellsCollection.MaxColumnUsed);

            if (!onlyUsedCells)
            {
                var lastRow = Worksheet.Internals.CellsCollection.MaxRowUsed;
                if (lastRow > 0)
                {
                    for (var ro = lastRow; ro >= firstRow; ro--)
                    {
                        var newRow = ro + numberOfRows;

                        for (var co = lastColumn; co >= firstColumn; co--)
                        {
                            var oldCell = Worksheet.Internals.CellsCollection.GetCell(ro, co);
                            if (oldCell == null)
                            {
                                continue;
                            }

                            var oldKey = new XLAddress(Worksheet, ro, co, false, false);
                            var newKey = new XLAddress(Worksheet, newRow, co, false, false);

                            oldCell.Address = newKey;
                            if (newKey.IsValid)
                            {
                                cellsToInsert.Add(newKey, oldCell);
                            }

                            cellsToDelete.Add(oldKey);
                        }
                        if (IsEntireRow())
                        {
                            Worksheet.Row(newRow).Height = Worksheet.Row(ro).Height;
                        }
                    }
                }
            }
            else
            {
                foreach (
                    var c in
                        Worksheet.Internals.CellsCollection.GetCells(firstRow, firstColumn,
                                                                     Worksheet.Internals.CellsCollection.MaxRowUsed,
                                                                     lastColumn))
                {
                    var newRow = c.Address.RowNumber + numberOfRows;
                    var newKey = new XLAddress(Worksheet, newRow, c.Address.ColumnNumber, false, false);
                    cellsToDelete.Add(c.Address);
                    c.Address = newKey;
                    if (newKey.IsValid)
                    {
                        cellsToInsert.Add(newKey, c);
                    }
                }
            }

            cellsToDelete.ForEach(c => Worksheet.Internals.CellsCollection.Remove(c.RowNumber, c.ColumnNumber));
            cellsToInsert.ForEach(c => Worksheet.Internals.CellsCollection.Add(c.Key.RowNumber, c.Key.ColumnNumber, c.Value));

            var firstRowReturn = RangeAddress.FirstAddress.RowNumber;
            var lastRowReturn = RangeAddress.FirstAddress.RowNumber + numberOfRows - 1;
            var firstColumnReturn = RangeAddress.FirstAddress.ColumnNumber;
            var lastColumnReturn = RangeAddress.LastAddress.ColumnNumber;

            Worksheet.NotifyRangeShiftedRows(AsRange(), numberOfRows);

            var rangeToReturn = Worksheet.Range(firstRowReturn, firstColumnReturn, lastRowReturn, lastColumnReturn);

            // We deliberately ignore conditional formats and data validation here. Their shifting is handled elsewhere
            var contentFlags = XLCellsUsedOptions.All
                & ~XLCellsUsedOptions.ConditionalFormats
                & ~XLCellsUsedOptions.DataValidation;

            if (formatFromAbove && rangeToReturn.RangeAddress.FirstAddress.RowNumber > 1)
            {
                var fr = rangeToReturn.FirstRow();
                var model = fr.RowAbove();
                var modelFirstColumn = (model as IXLRangeBase).FirstCellUsed(contentFlags);
                var modelLastColumn = (model as IXLRangeBase).LastCellUsed(contentFlags);
                if (modelFirstColumn != null && modelLastColumn != null)
                {
                    var firstCoReturned = modelFirstColumn.Address.ColumnNumber
                                            - model.RangeAddress.FirstAddress.ColumnNumber + 1;
                    var lastCoReturned = modelLastColumn.Address.ColumnNumber
                                            - model.RangeAddress.FirstAddress.ColumnNumber + 1;
                    for (var co = firstCoReturned; co <= lastCoReturned; co++)
                    {
                        rangeToReturn.Column(co).Style = model.Cell(co).Style;
                    }
                }
            }
            else
            {
                var lastCoUsed = rangeToReturn.LastColumnUsed(contentFlags);
                if (lastCoUsed != null)
                {
                    var lastCoReturned = lastCoUsed.ColumnNumber();
                    for (var co = 1; co <= lastCoReturned; co++)
                    {
                        var styleToUse =
                            Worksheet.Internals.ColumnsCollection.TryGetValue(co, out var column)
                                ? column.Style
                                : Worksheet.Style;

                        rangeToReturn.Style = styleToUse;
                    }
                }
            }

            // Skip calling .Rows() for performance reasons if required.
            if (nullReturn)
            {
                return null;
            }

            return rangeToReturn.Rows();
        }

        private void ClearMerged()
        {
            var mergeToDelete = Worksheet.Internals.MergedRanges.GetIntersectedRanges(RangeAddress).ToList();
            mergeToDelete.ForEach(m => Worksheet.Internals.MergedRanges.Remove(m));
        }

        public bool Contains(IXLCell cell)
        {
            return Contains((XLAddress)cell.Address);
        }

        public bool Contains(XLAddress first, XLAddress last)
        {
            return Contains(first) && Contains(last);
        }

        public bool Contains(XLAddress address)
        {
            return RangeAddress.Contains(in address);
        }

        public void Delete(XLShiftDeletedCells shiftDeleteCells)
        {
            var numberOfRows = RowCount();
            var numberOfColumns = ColumnCount();

            if (!RangeAddress.IsValid)
            {
                return;
            }

            Worksheet.SparklineGroups.Remove(this);

            IXLRange shiftedRangeFormula = Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber,
                RangeAddress.FirstAddress.ColumnNumber,
                RangeAddress.LastAddress.RowNumber,
                RangeAddress.LastAddress.ColumnNumber);

            // Shift formulas first
            foreach (var cell in Worksheet
                .Workbook
                .Worksheets
                .Cast<XLWorksheet>()
                .SelectMany(ws => ws
                    .Internals
                    .CellsCollection
                    .GetCells(c => c.HasFormula)))
            {
                if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
                {
                    cell.ShiftFormulaRows((XLRange)shiftedRangeFormula, numberOfRows * -1);
                }
                else
                {
                    cell.ShiftFormulaColumns((XLRange)shiftedRangeFormula, numberOfColumns * -1);
                }
            }

            // Range to shift...
            var cellsToInsert = new Dictionary<IXLAddress, XLCell>();
            var cellsToDelete = new List<IXLAddress>();

            var columnModifier = 0;
            var rowModifier = 0;
            IEnumerable<XLCell> cellsQuery;
            switch (shiftDeleteCells)
            {
                case XLShiftDeletedCells.ShiftCellsLeft:
                    cellsQuery = Worksheet.Internals.CellsCollection.GetCells(
                        RangeAddress.FirstAddress.RowNumber,
                        RangeAddress.FirstAddress.ColumnNumber,
                        RangeAddress.LastAddress.RowNumber,
                        Worksheet.Internals.CellsCollection.MaxColumnUsed);

                    columnModifier = ColumnCount();

                    break;

                case XLShiftDeletedCells.ShiftCellsUp:
                    cellsQuery = Worksheet.Internals.CellsCollection.GetCells(
                        RangeAddress.FirstAddress.RowNumber,
                        RangeAddress.FirstAddress.ColumnNumber,
                        Worksheet.Internals.CellsCollection.MaxRowUsed,
                        RangeAddress.LastAddress.ColumnNumber);

                    rowModifier = RowCount();

                    break;

                default:
                    cellsQuery = new XLCell[] { };
                    break;
            }

            foreach (var c in cellsQuery)
            {
                // Schedule for removal from CellsCollection
                cellsToDelete.Add(c.Address);

                // Generate new cell to insert into CellsCollection
                var newCellAddress = new XLAddress(Worksheet, c.Address.RowNumber - rowModifier,
                                           c.Address.ColumnNumber - columnModifier,
                                           fixedRow: false,
                                           fixedColumn: false);

                if (newCellAddress.IsValid)
                {
                    var canInsert = shiftDeleteCells == XLShiftDeletedCells.ShiftCellsLeft
                                         ? c.Address.ColumnNumber > RangeAddress.LastAddress.ColumnNumber
                                         : c.Address.RowNumber > RangeAddress.LastAddress.RowNumber;

                    c.Address = newCellAddress;

                    if (canInsert)
                    {
                        cellsToInsert.Add(newCellAddress, c);
                    }
                }
            }

            cellsToDelete.ForEach(c => Worksheet.Internals.CellsCollection.Remove(c.RowNumber, c.ColumnNumber));
            cellsToInsert.ForEach(
                c => Worksheet.Internals.CellsCollection.Add(c.Key.RowNumber, c.Key.ColumnNumber, c.Value));

            var mergesToRemove = Worksheet.Internals.MergedRanges.Where(Contains).ToList();
            mergesToRemove.ForEach(r => Worksheet.Internals.MergedRanges.Remove(r));

            var hyperlinksToRemove = Worksheet.Hyperlinks.Where(hl => Contains(hl.Cell.AsRange())).ToList();
            hyperlinksToRemove.ForEach(hl => Worksheet.Hyperlinks.Delete(hl));

            var shiftedRange = AsRange();
            if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
            {
                Worksheet.NotifyRangeShiftedRows(shiftedRange, rowModifier * -1);
            }
            else
            {
                Worksheet.NotifyRangeShiftedColumns(shiftedRange, columnModifier * -1);
            }

            Worksheet.DeleteRange(RangeAddress);
        }

        public override string ToString()
        {
            return string.Concat(
                Worksheet.Name.EscapeSheetName(),
                '!',
                RangeAddress.FirstAddress,
                ':',
                RangeAddress.LastAddress);
        }

        protected IXLRangeAddress ShiftColumns(IXLRangeAddress thisRangeAddress, XLRange shiftedRange, int columnsShifted)
        {
            if (!thisRangeAddress.IsValid || !shiftedRange.RangeAddress.IsValid)
            {
                return thisRangeAddress;
            }

            var allRowsAreCovered = thisRangeAddress.FirstAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber &&
                                     thisRangeAddress.LastAddress.RowNumber <= shiftedRange.RangeAddress.LastAddress.RowNumber;

            if (!allRowsAreCovered)
            {
                return thisRangeAddress;
            }

            var shiftLeftBoundary = (columnsShifted > 0 && thisRangeAddress.FirstAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber) ||
                                     (columnsShifted < 0 && thisRangeAddress.FirstAddress.ColumnNumber > shiftedRange.RangeAddress.FirstAddress.ColumnNumber);

            var shiftRightBoundary = thisRangeAddress.LastAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber;

            var newLeftBoundary = thisRangeAddress.FirstAddress.ColumnNumber;
            if (shiftLeftBoundary)
            {
                if (newLeftBoundary + columnsShifted > shiftedRange.RangeAddress.FirstAddress.ColumnNumber)
                {
                    newLeftBoundary = newLeftBoundary + columnsShifted;
                }
                else
                {
                    newLeftBoundary = shiftedRange.RangeAddress.FirstAddress.ColumnNumber;
                }
            }

            var newRightBoundary = thisRangeAddress.LastAddress.ColumnNumber;
            if (shiftRightBoundary)
            {
                newRightBoundary += columnsShifted;
            }

            var destroyedByShift = newRightBoundary < newLeftBoundary;

            var firstAddress = (XLAddress)thisRangeAddress.FirstAddress;
            var lastAddress = (XLAddress)thisRangeAddress.LastAddress;

            if (destroyedByShift)
            {
                firstAddress = Worksheet.InvalidAddress;
                lastAddress = Worksheet.InvalidAddress;
                Worksheet.DeleteRange(RangeAddress);
            }

            if (shiftLeftBoundary)
            {
                firstAddress = new XLAddress(Worksheet,
                                             thisRangeAddress.FirstAddress.RowNumber,
                                             newLeftBoundary,
                                             thisRangeAddress.FirstAddress.FixedRow,
                                             thisRangeAddress.FirstAddress.FixedColumn);
            }

            if (shiftRightBoundary)
            {
                lastAddress = new XLAddress(Worksheet,
                                            thisRangeAddress.LastAddress.RowNumber,
                                            newRightBoundary,
                                            thisRangeAddress.LastAddress.FixedRow,
                                            thisRangeAddress.LastAddress.FixedColumn);
            }

            return new XLRangeAddress(firstAddress, lastAddress);
        }

        protected IXLRangeAddress ShiftRows(IXLRangeAddress thisRangeAddress, XLRange shiftedRange, int rowsShifted)
        {
            if (!thisRangeAddress.IsValid || !shiftedRange.RangeAddress.IsValid)
            {
                return thisRangeAddress;
            }

            var allColumnsAreCovered = thisRangeAddress.FirstAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber &&
                                        thisRangeAddress.LastAddress.ColumnNumber <= shiftedRange.RangeAddress.LastAddress.ColumnNumber;

            if (!allColumnsAreCovered)
            {
                return thisRangeAddress;
            }

            var shiftTopBoundary = (rowsShifted > 0 && thisRangeAddress.FirstAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber) ||
                                    (rowsShifted < 0 && thisRangeAddress.FirstAddress.RowNumber > shiftedRange.RangeAddress.FirstAddress.RowNumber);

            var shiftBottomBoundary = thisRangeAddress.LastAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber;

            var newTopBoundary = thisRangeAddress.FirstAddress.RowNumber;
            if (shiftTopBoundary)
            {
                if (newTopBoundary + rowsShifted > shiftedRange.RangeAddress.FirstAddress.RowNumber)
                {
                    newTopBoundary = newTopBoundary + rowsShifted;
                }
                else
                {
                    newTopBoundary = shiftedRange.RangeAddress.FirstAddress.RowNumber;
                }
            }

            var newBottomBoundary = thisRangeAddress.LastAddress.RowNumber;
            if (shiftBottomBoundary)
            {
                newBottomBoundary += rowsShifted;
            }

            var destroyedByShift = newBottomBoundary < newTopBoundary;

            var firstAddress = (XLAddress)thisRangeAddress.FirstAddress;
            var lastAddress = (XLAddress)thisRangeAddress.LastAddress;

            if (destroyedByShift)
            {
                firstAddress = Worksheet.InvalidAddress;
                lastAddress = Worksheet.InvalidAddress;
                Worksheet.DeleteRange(RangeAddress);
            }

            if (shiftTopBoundary)
            {
                firstAddress = new XLAddress(Worksheet,
                                             newTopBoundary,
                                             thisRangeAddress.FirstAddress.ColumnNumber,
                                             thisRangeAddress.FirstAddress.FixedRow,
                                             thisRangeAddress.FirstAddress.FixedColumn);
            }

            if (shiftBottomBoundary)
            {
                lastAddress = new XLAddress(Worksheet,
                                            newBottomBoundary,
                                            thisRangeAddress.LastAddress.ColumnNumber,
                                            thisRangeAddress.LastAddress.FixedRow,
                                            thisRangeAddress.LastAddress.FixedColumn);
            }

            return new XLRangeAddress(firstAddress, lastAddress);
        }

        public IXLRange RangeUsed()
        {
            return RangeUsed(XLCellsUsedOptions.AllContents);
        }

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        public IXLRange RangeUsed(bool includeFormats)
        {
            return RangeUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents);
        }

        public IXLRange RangeUsed(XLCellsUsedOptions options)
        {
            var firstCell = (this as IXLRangeBase).FirstCellUsed(options);
            if (firstCell == null)
            {
                return null;
            }

            var lastCell = (this as IXLRangeBase).LastCellUsed(options);
            return Worksheet.Range(firstCell, lastCell);
        }

        public virtual void CopyTo(IXLRangeBase target)
        {
            CopyTo(target.FirstCell());
        }

        public virtual void CopyTo(IXLCell target)
        {
            target.Value = this;
        }

        //public IXLChart CreateChart(Int32 firstRow, Int32 firstColumn, Int32 lastRow, Int32 lastColumn)
        //{
        //    IXLChart chart = new XLChartWorksheet;
        //    chart.FirstRow = firstRow;
        //    chart.LastRow = lastRow;
        //    chart.LastColumn = lastColumn;
        //    chart.FirstColumn = firstColumn;
        //    Worksheet.Charts.Add(chart);
        //    return chart;
        //}

        IXLPivotTable IXLRangeBase.CreatePivotTable(IXLCell targetCell, string name)
        {
            return CreatePivotTable(targetCell, name);
        }

        public XLPivotTable CreatePivotTable(IXLCell targetCell, string name)
        {
            return (XLPivotTable)targetCell.Worksheet.PivotTables.Add(name, targetCell, AsRange());
        }

        public virtual IXLAutoFilter SetAutoFilter()
        {
            return SetAutoFilter(true);
        }

        public IXLAutoFilter SetAutoFilter(bool value)
        {
            if (value)
            {
                return Worksheet.AutoFilter.Set(this);
            }
            else
            {
                return Worksheet.AutoFilter.Clear();
            }
        }

        #region Sort

        public IXLSortElements SortRows => _sortRows ??= new XLSortElements();

        public IXLSortElements SortColumns => _sortColumns ??= new XLSortElements();

        private string DefaultSortString()
        {
            var sb = new StringBuilder();
            var maxColumn = ColumnCount();
            if (maxColumn == XLHelper.MaxColumnNumber)
            {
                maxColumn = (this as IXLRangeBase).LastCellUsed(XLCellsUsedOptions.All).Address.ColumnNumber;
            }

            for (var i = 1; i <= maxColumn; i++)
            {
                if (sb.Length > 0)
                {
                    sb.Append(',');
                }

                sb.Append(i);
            }

            return sb.ToString();
        }

        public IXLRangeBase Sort()
        {
            if (!SortColumns.Any())
            {
                return Sort(DefaultSortString());
            }

            SortRangeRows();
            return this;
        }

        public IXLRangeBase Sort(string columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true)
        {
            SortColumns.Clear();
            if (string.IsNullOrWhiteSpace(columnsToSortBy))
            {
                columnsToSortBy = DefaultSortString();
            }

            foreach (var coPairTrimmed in columnsToSortBy.Split(',').Select(coPair => coPair.Trim()))
            {
                string coString;
                string order;
                if (coPairTrimmed.Contains(' '))
                {
                    var pair = coPairTrimmed.Split(' ');
                    coString = pair[0];
                    order = pair[1];
                }
                else
                {
                    coString = coPairTrimmed;
                    order = sortOrder == XLSortOrder.Ascending ? "ASC" : "DESC";
                }

                if (!int.TryParse(coString, out var co))
                {
                    co = XLHelper.GetColumnNumberFromLetter(coString);
                }

                SortColumns.Add(co, string.Compare(order, "ASC", true, CultureInfo.CurrentCulture) == 0 ? XLSortOrder.Ascending : XLSortOrder.Descending, ignoreBlanks, matchCase);
            }

            SortRangeRows();
            return this;
        }

        public IXLRangeBase Sort(int columnToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true)
        {
            return Sort(columnToSortBy.ToString(), sortOrder, matchCase, ignoreBlanks);
        }

        public IXLRangeBase SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true)
        {
            SortRows.Clear();
            var maxColumn = ColumnCount();
            if (maxColumn == XLHelper.MaxColumnNumber)
            {
                maxColumn = (this as IXLRangeBase).LastCellUsed(XLCellsUsedOptions.All).Address.ColumnNumber;
            }

            for (var i = 1; i <= maxColumn; i++)
            {
                SortRows.Add(i, sortOrder, ignoreBlanks, matchCase);
            }

            SortRangeColumns();
            return this;
        }

        #region Sort Rows

        private void SortRangeRows()
        {
            var maxRow = RowCount();
            if (maxRow == XLHelper.MaxRowNumber)
            {
                maxRow = (this as IXLRangeBase).LastCellUsed(XLCellsUsedOptions.All).Address.RowNumber;
            }

            SortingRangeRows(1, maxRow);
        }

        private void SwapRows(int row1, int row2)
        {
            var row1InWs = RangeAddress.FirstAddress.RowNumber + row1 - 1;
            var row2InWs = RangeAddress.FirstAddress.RowNumber + row2 - 1;

            var firstColumn = RangeAddress.FirstAddress.ColumnNumber;
            var lastColumn = RangeAddress.LastAddress.ColumnNumber;

            var range1Sp1 = new XLSheetPoint(row1InWs, firstColumn);
            var range1Sp2 = new XLSheetPoint(row1InWs, lastColumn);
            var range2Sp1 = new XLSheetPoint(row2InWs, firstColumn);
            var range2Sp2 = new XLSheetPoint(row2InWs, lastColumn);

            Worksheet.Internals.CellsCollection.SwapRanges(new XLSheetRange(range1Sp1, range1Sp2),
                                                           new XLSheetRange(range2Sp1, range2Sp2), Worksheet);
        }

        private int SortRangeRows(int begPoint, int endPoint)
        {
            var pivot = begPoint;
            var m = begPoint + 1;
            var n = endPoint;
            while ((m < endPoint) && RowQuick(pivot).CompareTo(RowQuick(m), SortColumns) >= 0)
            {
                m++;
            }

            while (n > begPoint && RowQuick(pivot).CompareTo(RowQuick(n), SortColumns) <= 0)
            {
                n--;
            }

            while (m < n)
            {
                SwapRows(m, n);

                while (m < endPoint && RowQuick(pivot).CompareTo(RowQuick(m), SortColumns) >= 0)
                {
                    m++;
                }

                while (n > begPoint && RowQuick(pivot).CompareTo(RowQuick(n), SortColumns) <= 0)
                {
                    n--;
                }
            }

            if (pivot != n)
            {
                SwapRows(n, pivot);
            }

            return n;
        }

        private void SortingRangeRows(int beg, int end)
        {
            if (beg == end)
            {
                return;
            }

            var pivot = SortRangeRows(beg, end);
            if (pivot > beg)
            {
                SortingRangeRows(beg, pivot - 1);
            }

            if (pivot < end)
            {
                SortingRangeRows(pivot + 1, end);
            }
        }

        #endregion Sort Rows

        #region Sort Columns

        private void SortRangeColumns()
        {
            var maxColumn = ColumnCount();
            if (maxColumn == XLHelper.MaxColumnNumber)
            {
                maxColumn = (this as IXLRangeBase).LastCellUsed(XLCellsUsedOptions.All).Address.ColumnNumber;
            }

            SortingRangeColumns(1, maxColumn);
        }

        private void SwapColumns(int column1, int column2)
        {
            var col1InWs = RangeAddress.FirstAddress.ColumnNumber + column1 - 1;
            var col2InWs = RangeAddress.FirstAddress.ColumnNumber + column2 - 1;

            var firstRow = RangeAddress.FirstAddress.RowNumber;
            var lastRow = RangeAddress.LastAddress.RowNumber;

            var range1Sp1 = new XLSheetPoint(firstRow, col1InWs);
            var range1Sp2 = new XLSheetPoint(lastRow, col1InWs);
            var range2Sp1 = new XLSheetPoint(firstRow, col2InWs);
            var range2Sp2 = new XLSheetPoint(lastRow, col2InWs);

            Worksheet.Internals.CellsCollection.SwapRanges(new XLSheetRange(range1Sp1, range1Sp2),
                                                           new XLSheetRange(range2Sp1, range2Sp2), Worksheet);
        }

        private int SortRangeColumns(int begPoint, int endPoint)
        {
            var pivot = begPoint;
            var m = begPoint + 1;
            var n = endPoint;
            while ((m < endPoint) && ColumnQuick(pivot).CompareTo(ColumnQuick(m), SortRows) >= 0)
            {
                m++;
            }

            while ((n > begPoint) && (ColumnQuick(pivot).CompareTo(ColumnQuick(n), SortRows) <= 0))
            {
                n--;
            }

            while (m < n)
            {
                SwapColumns(m, n);

                while ((m < endPoint) && ColumnQuick(pivot).CompareTo(ColumnQuick(m), SortRows) >= 0)
                {
                    m++;
                }

                while ((n > begPoint) && ColumnQuick(pivot).CompareTo(ColumnQuick(n), SortRows) <= 0)
                {
                    n--;
                }
            }
            if (pivot != n)
            {
                SwapColumns(n, pivot);
            }

            return n;
        }

        private void SortingRangeColumns(int beg, int end)
        {
            if (end == beg)
            {
                return;
            }

            var pivot = SortRangeColumns(beg, end);
            if (pivot > beg)
            {
                SortingRangeColumns(beg, pivot - 1);
            }

            if (pivot < end)
            {
                SortingRangeColumns(pivot + 1, end);
            }
        }

        #endregion Sort Columns

        #endregion Sort

        public XLRangeColumn ColumnQuick(int column)
        {
            var firstCellAddress = new XLAddress(Worksheet,
                                                 RangeAddress.FirstAddress.RowNumber,
                                                 RangeAddress.FirstAddress.ColumnNumber + column - 1,
                                                 false,
                                                 false);
            var lastCellAddress = new XLAddress(Worksheet,
                                                RangeAddress.LastAddress.RowNumber,
                                                RangeAddress.FirstAddress.ColumnNumber + column - 1,
                                                false,
                                                false);
            return Worksheet.RangeColumn(new XLRangeAddress(firstCellAddress, lastCellAddress));
        }

        public XLRangeRow RowQuick(int row)
        {
            var firstCellAddress = new XLAddress(Worksheet,
                                                 RangeAddress.FirstAddress.RowNumber + row - 1,
                                                 RangeAddress.FirstAddress.ColumnNumber,
                                                 false,
                                                 false);
            var lastCellAddress = new XLAddress(Worksheet,
                                                RangeAddress.FirstAddress.RowNumber + row - 1,
                                                RangeAddress.LastAddress.ColumnNumber,
                                                false,
                                                false);

            return Worksheet.RangeRow(new XLRangeAddress(firstCellAddress, lastCellAddress));
        }

        [Obsolete("Use GetDataValidation() to access the existing rule, or CreateDataValidation() to create a new one.")]
        public IXLDataValidation SetDataValidation()
        {
            var existingValidation = GetDataValidation();
            if (existingValidation != null && existingValidation.Ranges.Any(r => r == this))
            {
                return existingValidation;
            }

            var dataValidationToCopy = Worksheet.DataValidations.GetAllInRange(RangeAddress)
                .FirstOrDefault();

            var newRange = AsRange();
            var dataValidation = new XLDataValidation(newRange);
            if (dataValidationToCopy != null)
            {
                dataValidation.CopyFrom(dataValidationToCopy);
            }

            Worksheet.DataValidations.Add(dataValidation);
            return dataValidation;
        }

        public IXLConditionalFormat AddConditionalFormat()
        {
            var cf = new XLConditionalFormat(AsRange());
            Worksheet.ConditionalFormats.Add(cf);
            return cf;
        }

        internal IXLConditionalFormat AddConditionalFormat(IXLConditionalFormat source)
        {
            var cf = new XLConditionalFormat(AsRange());
            cf.CopyFrom(source);
            Worksheet.ConditionalFormats.Add(cf);
            return cf;
        }

        public void Select()
        {
            Worksheet.SelectedRanges.Add(AsRange());
        }

        public IXLRangeBase Grow()
        {
            return Grow(1);
        }

        public IXLRangeBase Grow(int growCount)
        {
            var firstRow = Math.Max(1, RangeAddress.FirstAddress.RowNumber - growCount);
            var firstColumn = Math.Max(1, RangeAddress.FirstAddress.ColumnNumber - growCount);

            var lastRow = Math.Min(XLHelper.MaxRowNumber, RangeAddress.LastAddress.RowNumber + growCount);
            var lastColumn = Math.Min(XLHelper.MaxColumnNumber, RangeAddress.LastAddress.ColumnNumber + growCount);

            return Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
        }

        public IXLRangeBase Shrink()
        {
            return Shrink(1);
        }

        public IXLRangeBase Shrink(int shrinkCount)
        {
            var firstRow = RangeAddress.FirstAddress.RowNumber + shrinkCount;
            var firstColumn = RangeAddress.FirstAddress.ColumnNumber + shrinkCount;

            var lastRow = RangeAddress.LastAddress.RowNumber - shrinkCount;
            var lastColumn = RangeAddress.LastAddress.ColumnNumber - shrinkCount;

            if (firstRow > lastRow || firstColumn > lastColumn)
            {
                return null;
            }

            return Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
        }

        public IXLRangeAddress Intersection(IXLRangeBase otherRange, Func<IXLCell, bool> thisRangePredicate = null, Func<IXLCell, bool> otherRangePredicate = null)
        {
            if (otherRange == null)
            {
                return null;
            }

            if (!Worksheet.Equals(otherRange.Worksheet))
            {
                return null;
            }

            if (thisRangePredicate == null && otherRangePredicate == null)
            {
                // Special case, no predicates. We can optimise this a bit then.
                return RangeAddress.Intersection(otherRange.RangeAddress);
            }
            else
            {
                thisRangePredicate = thisRangePredicate ?? (c => true);
                otherRangePredicate = otherRangePredicate ?? (c => true);

                var intersectionCells = Cells(c => thisRangePredicate(c) && otherRange.Cells(otherRangePredicate).Contains(c));

                if (!intersectionCells.Any())
                {
                    return null;
                }

                var firstRow = intersectionCells.Min(c => c.Address.RowNumber);
                var firstColumn = intersectionCells.Min(c => c.Address.ColumnNumber);

                var lastRow = intersectionCells.Max(c => c.Address.RowNumber);
                var lastColumn = intersectionCells.Max(c => c.Address.ColumnNumber);

                return new XLRangeAddress
                (
                    new XLAddress(Worksheet, firstRow, firstColumn, fixedRow: false, fixedColumn: false),
                    new XLAddress(Worksheet, lastRow, lastColumn, fixedRow: false, fixedColumn: false)
                );
            }
        }

        public IXLCells SurroundingCells(Func<IXLCell, bool> predicate = null)
        {
            var cells = new XLCells(false, XLCellsUsedOptions.AllContents, predicate);
            Grow().Cells(c => !Contains(c)).ForEach(c => cells.Add(c as XLCell));
            return cells;
        }

        public IXLCells Union(IXLRangeBase otherRange, Func<IXLCell, bool> thisRangePredicate = null, Func<IXLCell, bool> otherRangePredicate = null)
        {
            if (otherRange == null)
            {
                return Cells(thisRangePredicate);
            }

            var cells = new XLCells(false, XLCellsUsedOptions.AllContents);
            if (!Worksheet.Equals(otherRange.Worksheet))
            {
                return cells;
            }

            if (thisRangePredicate == null)
            {
                thisRangePredicate = c => true;
            }

            if (otherRangePredicate == null)
            {
                otherRangePredicate = c => true;
            }

            Cells(thisRangePredicate).Concat(otherRange.Cells(otherRangePredicate)).Distinct().ForEach(c => cells.Add(c as XLCell));
            return cells;
        }

        public IXLCells Difference(IXLRangeBase otherRange, Func<IXLCell, bool> thisRangePredicate = null, Func<IXLCell, bool> otherRangePredicate = null)
        {
            if (otherRange == null)
            {
                return Cells(thisRangePredicate);
            }

            var cells = new XLCells(false, XLCellsUsedOptions.AllContents);
            if (!Worksheet.Equals(otherRange.Worksheet))
            {
                return cells;
            }

            if (thisRangePredicate == null)
            {
                thisRangePredicate = c => true;
            }

            if (otherRangePredicate == null)
            {
                otherRangePredicate = c => true;
            }

            Cells(c => thisRangePredicate(c) && !otherRange.Cells(otherRangePredicate).Contains(c)).ForEach(c => cells.Add(c as XLCell));
            return cells;
        }

        private IEnumerable<IXLCell> CellsUsedInternal(XLCellsUsedOptions options, Func<IXLRange, IXLCell> selector, Func<IXLCell, bool> predicate)
        {
            predicate ??= (t => true);

            //To avoid unnecessary initialization of thousands cells
            var opt = options
                      & ~XLCellsUsedOptions.ConditionalFormats
                      & ~XLCellsUsedOptions.DataValidation
                      & ~XLCellsUsedOptions.MergedRanges;

            // If opt == 0 then we're basically back at unconstrained, so just set back the original options
            if (opt == XLCellsUsedOptions.NoConstraints)
            {
                opt = options;
            }

            IEnumerable<IXLCell> cellsUsed = CellsUsed(opt, predicate);

            if (options.HasFlag(XLCellsUsedOptions.ConditionalFormats))
            {
                cellsUsed = cellsUsed.Union(
                    Worksheet.ConditionalFormats
                        .SelectMany(cf => cf.Ranges.GetIntersectedRanges(RangeAddress))
                        .Select(selector)
                        .Where(predicate)
                );
            }
            if (options.HasFlag(XLCellsUsedOptions.DataValidation))
            {
                cellsUsed = cellsUsed.Union(
                    Worksheet.DataValidations
                        .GetAllInRange(RangeAddress)
                        .SelectMany(dv => dv.Ranges)
                        .Select(selector)
                        .Where(predicate)
                );
            }
            if (options.HasFlag(XLCellsUsedOptions.MergedRanges))
            {
                cellsUsed = cellsUsed.Union(
                    Worksheet.MergedRanges.GetIntersectedRanges(RangeAddress)
                        .Select(selector)
                        .Where(predicate)
                );
            }

            return cellsUsed;
        }
    }
}