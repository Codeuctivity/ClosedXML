﻿

namespace ClosedXML.Excel
{
    public class LoadOptions
    {
        public XLEventTracking EventTracking { get; set; } = XLEventTracking.Enabled;
        public bool RecalculateAllFormulas { get; set; }
    }
}
