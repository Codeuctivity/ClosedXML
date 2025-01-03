
using System;
using System.Linq;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal class XLAutoFilter : IXLAutoFilter
    {
        private readonly Dictionary<int, XLFilterColumn> _columns = new Dictionary<int, XLFilterColumn>();

        public XLAutoFilter()
        {
            Filters = new Dictionary<int, List<XLFilter>>();
        }

        public Dictionary<int, List<XLFilter>> Filters { get; private set; }

        #region IXLAutoFilter Members

        [Obsolete("Use IsEnabled")]
        public bool Enabled { get => IsEnabled; set => IsEnabled = value; }
        public IEnumerable<IXLRangeRow> HiddenRows => Range.Rows(r => r.WorksheetRow().IsHidden);
        public bool IsEnabled { get; set; }
        public IXLRange Range { get; set; }
        public int SortColumn { get; set; }
        public bool Sorted { get; set; }
        public XLSortOrder SortOrder { get; set; }
        public IEnumerable<IXLRangeRow> VisibleRows => Range.Rows(r => !r.WorksheetRow().IsHidden);

        IXLAutoFilter IXLAutoFilter.Clear()
        {
            return Clear();
        }

        public IXLFilterColumn Column(string column)
        {
            var columnNumber = XLHelper.GetColumnNumberFromLetter(column);
            if (columnNumber < 1 || columnNumber > XLHelper.MaxColumnNumber)
            {
                throw new ArgumentOutOfRangeException(nameof(column), "Column '" + column + "' is outside the allowed column range.");
            }

            return Column(columnNumber);
        }

        public IXLFilterColumn Column(int column)
        {
            if (column < 1 || column > XLHelper.MaxColumnNumber)
            {
                throw new ArgumentOutOfRangeException(nameof(column), "Column " + column + " is outside the allowed column range.");
            }

            if (!_columns.TryGetValue(column, out var filterColumn))
            {
                filterColumn = new XLFilterColumn(this, column);
                _columns.Add(column, filterColumn);
            }

            return filterColumn;
        }

        public IXLAutoFilter Reapply()
        {
            var ws = Range.Worksheet as XLWorksheet;
            ws.SuspendEvents();

            // Recalculate shown / hidden rows
            var rows = Range.Rows(2, Range.RowCount());
            rows.ForEach(row =>
                row.WorksheetRow().Unhide()
            );

            foreach (var row in rows)
            {
                var rowMatch = true;

                foreach (var columnIndex in Filters.Keys)
                {
                    var columnFilters = Filters[columnIndex];

                    var columnFilterMatch = true;

                    // If the first filter is an 'Or', we need to fudge the initial condition
                    if (columnFilters.Count > 0 && columnFilters.First().Connector == XLConnector.Or)
                    {
                        columnFilterMatch = false;
                    }

                    foreach (var filter in columnFilters)
                    {
                        var condition = filter.Condition;
                        var isText = filter.Value is string;
                        var isDateTime = filter.Value is DateTime;

                        bool filterMatch;

                        if (isText)
                        {
                            filterMatch = condition(row.Cell(columnIndex).GetFormattedString());
                        }
                        else if (isDateTime)
                        {
                            filterMatch = row.Cell(columnIndex).DataType == XLDataType.DateTime &&
                                    condition(row.Cell(columnIndex).GetDateTime());
                        }
                        else
                        {
                            filterMatch = row.Cell(columnIndex).DataType == XLDataType.Number &&
                                    condition(row.Cell(columnIndex).GetDouble());
                        }

                        if (filter.Connector == XLConnector.And)
                        {
                            columnFilterMatch &= filterMatch;
                            if (!columnFilterMatch)
                            {
                                break;
                            }
                        }
                        else
                        {
                            columnFilterMatch |= filterMatch;
                            if (columnFilterMatch)
                            {
                                break;
                            }
                        }
                    }

                    rowMatch &= columnFilterMatch;

                    if (!rowMatch)
                    {
                        break;
                    }
                }

                if (!rowMatch)
                {
                    row.WorksheetRow().Hide();
                }
            }

            ws.ResumeEvents();
            return this;
        }

        IXLAutoFilter IXLAutoFilter.Sort(int columnToSortBy, XLSortOrder sortOrder, bool matchCase,
                                                                                                         bool ignoreBlanks)
        {
            return Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);
        }

        #endregion IXLAutoFilter Members

        public XLAutoFilter Clear()
        {
            if (!IsEnabled)
            {
                return this;
            }

            IsEnabled = false;
            Filters.Clear();
            foreach (var row in Range.Rows().Where(r => r.RowNumber() > 1))
            {
                row.WorksheetRow().Unhide();
            }

            return this;
        }

        public XLAutoFilter Set(IXLRangeBase range)
        {
            var firstOverlappingTable = range.Worksheet.Tables.FirstOrDefault(t => t.RangeUsed().Intersects(range));
            if (firstOverlappingTable != null)
            {
                throw new InvalidOperationException($"The range {range.RangeAddress.ToStringRelative(includeSheet: true)} is already part of table '{firstOverlappingTable.Name}'");
            }

            Range = range.AsRange();
            IsEnabled = true;
            return this;
        }

        public XLAutoFilter Sort(int columnToSortBy, XLSortOrder sortOrder, bool matchCase, bool ignoreBlanks)
        {
            if (!IsEnabled)
            {
                throw new InvalidOperationException("Filter has not been enabled.");
            }

            var ws = Range.Worksheet as XLWorksheet;
            ws.SuspendEvents();
            Range.Range(Range.FirstCell().CellBelow(), Range.LastCell()).Sort(columnToSortBy, sortOrder, matchCase,
                                                                              ignoreBlanks);

            Sorted = true;
            SortOrder = sortOrder;
            SortColumn = columnToSortBy;

            ws.ResumeEvents();

            Reapply();

            return this;
        }
    }
}
