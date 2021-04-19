using System;
using OfficeOpenXml;

namespace Helpers
{
    public static partial class Excel
    {
        public static void CheckExistedCellLocation(ExcelWorksheet worksheet, string location)
        {
            if (location != ExcelAddress.GetAddress(worksheet.Cells[location].Start.Row, worksheet.Cells[location].Start.Column))
                throw new Exception("Некорректный адрес ячейки: " + location);
        }
    }
}
