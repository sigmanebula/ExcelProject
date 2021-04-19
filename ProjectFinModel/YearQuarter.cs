using OfficeOpenXml;
using System.Collections.Generic;

namespace ProjectFinModel
{
    public class YearQuarter
    {
        public ExcelRangeBase Cell { get; set; }
        public Dictionary<string, ExcelRangeBase> Quarter = new Dictionary<string, ExcelRangeBase>();
    }
}
