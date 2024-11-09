
namespace ClosedXML.Excel
{
    public interface IXLPivotStyleFormat
    {
        XLPivotStyleFormatElement AppliesTo { get; }
        IXLPivotField PivotField { get; }
        IXLStyle Style { get; set; }
    }
}
