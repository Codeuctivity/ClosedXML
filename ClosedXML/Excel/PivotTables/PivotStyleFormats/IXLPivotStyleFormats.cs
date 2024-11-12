
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLPivotStyleFormats : IEnumerable<IXLPivotStyleFormat>
    {
        IXLPivotStyleFormat ForElement(XLPivotStyleFormatElement element);
    }
}
