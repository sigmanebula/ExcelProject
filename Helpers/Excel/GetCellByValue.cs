using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Data;

namespace Helpers
{
    public static partial class Excel
    {
        public static ExcelRangeBase GetCellByValue(ExcelWorksheet worksheet, string value)
        {
            return GetCellByValue(worksheet, value, 0, 0);
        }

        public static ExcelRangeBase GetCellByValue(ExcelWorksheet worksheet, string value, int rowStart, int rowEnd)
        {
            foreach (var cell in worksheet.Cells)
                if ((cell.Value ?? "").ToString() == value)
                    if (rowStart == 0 || cell.Start.Row >= rowStart)
                        if (rowEnd == 0 || cell.Start.Row <= rowEnd)
                            return cell;
            return null;
        }
    }
}
