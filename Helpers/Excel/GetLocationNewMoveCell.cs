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

        public static string GetLocationNewMoveCell(ExcelWorksheet worksheet, string cell, int[] rowsColumns)
        {
            return GetLocationNewMoveCell(worksheet, cell, rowsColumns[0], rowsColumns[1]);
        }

        public static string GetLocationNewMoveCell(ExcelWorksheet worksheet, string cell, int rows, int columns)
        {
            return ExcelAddress.GetAddress(worksheet.Cells[cell].Start.Row + rows, worksheet.Cells[cell].Start.Column + columns);
        }
    }
}
