using System;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void workSheetShown_1_Write(ExcelPackage excelPackage, string dateEnd)
        {
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetShown_1_Name);
            try
            {
                string dateEndFormated = Helpers.Sugar.ConvertDateToFormat(dateEnd, '-', '.', "YYYY.MM.DD", "DD.MM.YYYY");

                worksheet.Cells[Settings.SQLVariables.WorksheetShown_1_HeaderEndDateLocation].Value =
                    Convert.ToString(worksheet.Cells[Settings.SQLVariables.WorksheetShown_1_HeaderEndDateLocation].Value ?? "") + " " + dateEndFormated;

                worksheet.Cells[Settings.SQLVariables.WorksheetShown_1_SmallLabelEndDateLocation].Value =
                    Convert.ToString(worksheet.Cells[Settings.SQLVariables.WorksheetShown_1_SmallLabelEndDateLocation].Value ?? "") + " " + dateEndFormated;
                
                worksheet.Cells[Settings.SQLVariables.WorksheetShown_1_PeriodLocation].Value = getPeriodName(false);
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetShown_1_Name));
            }
        }
    }
}
