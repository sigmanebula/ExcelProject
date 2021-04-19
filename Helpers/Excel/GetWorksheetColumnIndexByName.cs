using OfficeOpenXml;

namespace Helpers
{
    public static partial class Excel
    {
        public static int GetWorksheetColumnIndexByName(ExcelWorksheet worksheet, string columnName)
        {
            return worksheet.Cells[columnName + "1"].Start.Column;
        }
    }
}
