
using System;

namespace ClosedXML.Excel
{
    internal class XLSheetView : IXLSheetView
    {
        private XLAddress _topLeftCellAddress;
        private int _zoomScale;

        public XLSheetView(XLWorksheet worksheet)
        {
            Worksheet = worksheet;
            View = XLSheetViewOptions.Normal;

            ZoomScale = 100;
            ZoomScaleNormal = 100;
            ZoomScalePageLayoutView = 100;
            ZoomScaleSheetLayoutView = 100;
        }

        public XLSheetView(XLWorksheet worksheet, XLSheetView sheetView)
            : this(worksheet)
        {
            SplitRow = sheetView.SplitRow;
            SplitColumn = sheetView.SplitColumn;
            FreezePanes = sheetView.FreezePanes;
            TopLeftCellAddress = new XLAddress(Worksheet, sheetView.TopLeftCellAddress.RowNumber, sheetView.TopLeftCellAddress.ColumnNumber, sheetView.TopLeftCellAddress.FixedRow, sheetView.TopLeftCellAddress.FixedColumn);
        }

        public bool FreezePanes { get; set; }
        public int SplitColumn { get; set; }
        public int SplitRow { get; set; }

        IXLAddress IXLSheetView.TopLeftCellAddress { get => TopLeftCellAddress; set => TopLeftCellAddress = (XLAddress)value; }

        public XLAddress TopLeftCellAddress
        {
            get => _topLeftCellAddress;
            set
            {
                if (value.HasWorksheet && !value.Worksheet.Equals(Worksheet))
                {
                    throw new ArgumentException($"The value should be on the same worksheet as the sheet view.");
                }

                _topLeftCellAddress = value;
            }
        }

        public XLSheetViewOptions View { get; set; }

        IXLWorksheet IXLSheetView.Worksheet => Worksheet;
        public XLWorksheet Worksheet { get; internal set; }

        public int ZoomScale
        {
            get { return _zoomScale; }
            set
            {
                _zoomScale = value;
                switch (View)
                {
                    case XLSheetViewOptions.Normal:
                        ZoomScaleNormal = value;
                        break;

                    case XLSheetViewOptions.PageBreakPreview:
                        ZoomScalePageLayoutView = value;
                        break;

                    case XLSheetViewOptions.PageLayout:
                        ZoomScaleSheetLayoutView = value;
                        break;
                }
            }
        }

        public int ZoomScaleNormal { get; set; }

        public int ZoomScalePageLayoutView { get; set; }

        public int ZoomScaleSheetLayoutView { get; set; }

        public void Freeze(int rows, int columns)
        {
            SplitRow = rows;
            SplitColumn = columns;
            FreezePanes = true;
        }

        public void FreezeColumns(int columns)
        {
            SplitColumn = columns;
            FreezePanes = true;
        }

        public void FreezeRows(int rows)
        {
            SplitRow = rows;
            FreezePanes = true;
        }

        public IXLSheetView SetView(XLSheetViewOptions value)
        {
            View = value;
            return this;
        }
    }
}
