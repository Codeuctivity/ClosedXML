
namespace ClosedXML.Excel
{
    internal static class IntegerExtensions
    {
        public static bool Between(this int val, int from, int to)
        {
            return val >= from && val <= to;
        }
    }
}
