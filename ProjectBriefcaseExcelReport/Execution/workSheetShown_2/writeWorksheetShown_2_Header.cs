using System;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void writeWorksheetShown_2_Header(ExcelPackage excelPackage, ExcelWorksheet worksheet, string dateStart, string dateEnd)
        {
            try
            {
                //////////////////Главный заголовок
                worksheet.Cells[Settings.SQLVariables.WorksheetShown_2_Header_LabelStartCell].Value
                    = Convert.ToString(worksheet.Cells[Settings.SQLVariables.WorksheetShown_2_Header_LabelStartCell].Value ?? "")
                    + " "
                    + dateStart
                    + " - "
                    + dateEnd;
            }
            catch (Exception ex)
            {
                throw new Exception("Главный заголовок: " + ex.Message);
            }
        }

    }
}
