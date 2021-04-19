using System;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void workSheetShown_3_Write(ExcelPackage excelPackage, string dateStart, string dateEnd)
        {
            var workSheetShown_3 = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetShown_3_Name);
            try
            {
                dateStart = Helpers.Sugar.ConvertDateToFormat(dateStart, '-', '.', "YYYY.MM.DD", "DD.MM.YYYY");
                dateEnd = Helpers.Sugar.ConvertDateToFormat(dateEnd, '-', '.', "YYYY.MM.DD", "DD.MM.YYYY");

                //////////////////Главный заголовок
                writeWorksheetShown_3_Header(excelPackage, workSheetShown_3, dateStart, dateEnd);

                //////////////////На КРБИ утверждены изменения 
                writeWorksheetShown_3_Label(excelPackage, workSheetShown_3);
                
                //////////////////таблица и нижние строки
                writeWorksheetShown_3_1(excelPackage, workSheetShown_3, dateStart, dateEnd);

            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetShown_1_Name));
            }
        }
    }
}
