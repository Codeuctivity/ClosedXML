using System;
using System.Collections.Generic;
using System.Linq;
using System.Collections;
using ClosedXML.Excel.Style;

namespace ClosedXML.Excel
{
    internal class XLRows : XLStylizedBase, IXLRows, IXLStylized
    {
        private readonly List<XLRow> _rowsCollection = new List<XLRow>();
        private readonly XLWorksheet _worksheet;

        private bool IsMaterialized => _lazyEnumerable == null;

        private IEnumerable<XLRow> _lazyEnumerable;
        private IEnumerable<XLRow> Rows => _lazyEnumerable ?? _rowsCollection.AsEnumerable();


        /// <summary>
        /// Create a new instance of <see cref="XLRows"/>.
        /// </summary>
        /// <param name="worksheet">If worksheet is specified it means that the created instance represents
        /// all rows on a worksheet so changing its height will affect all rows.</param>
        /// <param name="defaultStyle">Default style to use when initializing child entries.</param>
        /// <param name="lazyEnumerable">A predefined enumerator of <see cref="XLRow"/> to support lazy initialization.</param>
        public XLRows(XLWorksheet worksheet, XLStyleValue defaultStyle = null, IEnumerable<XLRow> lazyEnumerable = null)
            : base(defaultStyle)
        {
            _worksheet = worksheet;
            _lazyEnumerable = lazyEnumerable;
        }

        #region IXLRows Members

        public IEnumerator<IXLRow> GetEnumerator()
        {
            return Rows.Cast<IXLRow>().OrderBy(r => r.RowNumber()).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public double Height
        {
            set
            {
                Rows.ForEach(c => c.Height = value);
                if (_worksheet == null)
                {
                    return;
                }

                _worksheet.RowHeight = value;
                _worksheet.Internals.RowsCollection.ForEach(r => r.Value.Height = value);
            }
        }

        public void Delete()
        {
            if (_worksheet != null)
            {
                _worksheet.Internals.RowsCollection.Clear();
                _worksheet.Internals.CellsCollection.Clear();
            }
            else
            {
                var toDelete = new Dictionary<IXLWorksheet, List<int>>();
                foreach (var r in Rows)
                {
                    if (!toDelete.TryGetValue(r.Worksheet, out var list))
                    {
                        list = new List<int>();
                        toDelete.Add(r.Worksheet, list);
                    }

                    list.Add(r.RowNumber());
                }

                foreach (var kp in toDelete)
                {
                    foreach (var r in kp.Value.OrderByDescending(r => r))
                    {
                        kp.Key.Row(r).Delete();
                    }
                }
            }
        }

        public IXLRows AdjustToContents()
        {
            Rows.ForEach(r => r.AdjustToContents());
            return this;
        }

        public IXLRows AdjustToContents(int startColumn)
        {
            Rows.ForEach(r => r.AdjustToContents(startColumn));
            return this;
        }

        public IXLRows AdjustToContents(int startColumn, int endColumn)
        {
            Rows.ForEach(r => r.AdjustToContents(startColumn, endColumn));
            return this;
        }

        public IXLRows AdjustToContents(double minHeight, double maxHeight)
        {
            Rows.ForEach(r => r.AdjustToContents(minHeight, maxHeight));
            return this;
        }

        public IXLRows AdjustToContents(int startColumn, double minHeight, double maxHeight)
        {
            Rows.ForEach(r => r.AdjustToContents(startColumn, minHeight, maxHeight));
            return this;
        }

        public IXLRows AdjustToContents(int startColumn, int endColumn, double minHeight, double maxHeight)
        {
            Rows.ForEach(r => r.AdjustToContents(startColumn, endColumn, minHeight, maxHeight));
            return this;
        }

        public void Hide()
        {
            Rows.ForEach(r => r.Hide());
        }

        public void Unhide()
        {
            Rows.ForEach(r => r.Unhide());
        }

        public void Group()
        {
            Group(false);
        }

        public void Group(int outlineLevel)
        {
            Group(outlineLevel, false);
        }

        public void Ungroup()
        {
            Ungroup(false);
        }

        public void Group(bool collapse)
        {
            Rows.ForEach(r => r.Group(collapse));
        }

        public void Group(int outlineLevel, bool collapse)
        {
            Rows.ForEach(r => r.Group(outlineLevel, collapse));
        }

        public void Ungroup(bool ungroupFromAll)
        {
            Rows.ForEach(r => r.Ungroup(ungroupFromAll));
        }

        public void Collapse()
        {
            Rows.ForEach(r => r.Collapse());
        }

        public void Expand()
        {
            Rows.ForEach(r => r.Expand());
        }

        public IXLCells Cells()
        {
            var cells = new XLCells(false, XLCellsUsedOptions.AllContents);
            foreach (var container in Rows)
            {
                cells.Add(container.RangeAddress);
            }

            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(true, XLCellsUsedOptions.AllContents);
            foreach (var container in Rows)
            {
                cells.Add(container.RangeAddress);
            }

            return cells;
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
            var cells = new XLCells(true, options);
            foreach (var container in Rows)
            {
                cells.Add(container.RangeAddress);
            }

            return cells;
        }

        public IXLRows AddHorizontalPageBreaks()
        {
            foreach (var row in Rows)
            {
                row.Worksheet.PageSetup.AddHorizontalPageBreak(row.RowNumber());
            }

            return this;
        }

        public IXLRows SetDataType(XLDataType dataType)
        {
            Rows.ForEach(c => c.DataType = dataType);
            return this;
        }

        #endregion IXLRows Members

        #region IXLStylized Members
        protected override IEnumerable<XLStylizedBase> Children
        {
            get
            {
                if (_worksheet != null)
                {
                    yield return _worksheet;
                }
                else
                {
                    foreach (var row in Rows)
                    {
                        yield return row;
                    }
                }
            }
        }

        public override IEnumerable<IXLStyle> Styles
        {
            get
            {
                yield return Style;
                if (_worksheet != null)
                {
                    yield return _worksheet.Style;
                }
                else
                {
                    foreach (var s in Rows.SelectMany(row => row.Styles))
                    {
                        yield return s;
                    }
                }
            }
        }
        
        public override IXLRanges RangesUsed
        {
            get
            {
                var retVal = new XLRanges();
                this.ForEach(c => retVal.Add(c.AsRange()));
                return retVal;
            }
        }

        #endregion IXLStylized Members

        public void Add(XLRow row)
        {
            Materialize();
            _rowsCollection.Add(row);
        }

        public IXLRows Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            Rows.ForEach(c => c.Clear(clearOptions));
            return this;
        }

        public void Select()
        {
            foreach (var range in this)
            {
                range.Select();
            }
        }

        private void Materialize()
        {
            if (IsMaterialized)
            {
                return;
            }

            _rowsCollection.AddRange(Rows);
            _lazyEnumerable = null;
        }
    }
}
