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
        public static void FillWorksheetEmptyValues(ExcelWorksheet worksheet, int rowStart, int columnStart, int rowEnd, int columnEnd)
        {
            for (int i = rowStart; i <= rowEnd; i++)
                for (int j = columnStart; j <= columnEnd; j++)
                    if (worksheet.Cells[i, j].Value == null)
                        worksheet.Cells[i, j].Value = "-";
        }
    }
}
