﻿
namespace ClosedXML.Excel
{
    public interface IXLPivotFieldStyleFormats
    {
        IXLPivotValueStyleFormat DataValuesFormat { get; }
        IXLPivotStyleFormat Header { get; }
        IXLPivotStyleFormat Label { get; }
        IXLPivotStyleFormat Subtotal { get; }
    }
}
