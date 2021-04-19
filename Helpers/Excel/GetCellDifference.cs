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
        public static int[] GetCellDifference(ExcelWorksheet worksheet, string cellFirst, string cellSecond)
        {
            return new int[] {
                  worksheet.Cells[cellSecond].Start.Row     - worksheet.Cells[cellFirst].Start.Row
                , worksheet.Cells[cellSecond].Start.Column  - worksheet.Cells[cellFirst].Start.Column
            };  //row, col
        }
    }
}
