using System;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void writeWorksheetShown_2_3(ExcelPackage excelPackage, ExcelWorksheet worksheet, string dateEnd)
        {
            try
            {
                //////////////////Заголовок круговой диаграммы
                worksheet.Cells[Settings.SQLVariables.WorksheetShown_2_3_Header_LabelStartCell].Value
                    = Convert.ToString(worksheet.Cells[Settings.SQLVariables.WorksheetShown_2_3_Header_LabelStartCell].Value ?? "")
                    + " "
                    + dateEnd;
            }
            catch (Exception ex)
            {
                throw new Exception("Заголовок круговой диаграммы: " + ex.Message);
            }
        }

    }
}
