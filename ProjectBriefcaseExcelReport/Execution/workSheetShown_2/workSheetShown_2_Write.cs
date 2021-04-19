using System;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void workSheetShown_2_Write(ExcelPackage excelPackage, string dateStart, string dateEnd)
        {
            var worksheetShown_2 = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetShown_2_Name);
            try
            {
                dateStart   = Helpers.Sugar.ConvertDateToFormat(dateStart,  '-', '.', "YYYY.MM.DD", "DD.MM.YYYY");
                dateEnd     = Helpers.Sugar.ConvertDateToFormat(dateEnd,    '-', '.', "YYYY.MM.DD", "DD.MM.YYYY");

                //////////////////Главный заголовок
                writeWorksheetShown_2_Header(excelPackage, worksheetShown_2, dateStart, dateEnd);

                //////////////////Заголовок круговой диаграммы
                writeWorksheetShown_2_3(excelPackage, worksheetShown_2, dateEnd);

                //////////////////Заголовок столбчатой диаграммы
                writeWorksheetShown_2_2_Header(excelPackage, worksheetShown_2, dateStart, dateEnd);

                //////////////////Здоровье динамических и водопадных проектов
                writeWorksheetShown_2_3_4(worksheetShown_2);

                //////////////////Структура портфеля технологических задач
                writeWorksheetShown_2_1(excelPackage, worksheetShown_2, dateEnd);
                
                //////////////////Исключены из портфеля
                writeWorksheetShown_2_5(excelPackage, worksheetShown_2);

                //////////////////Включены в портфель
                writeWorksheetShown_2_6(excelPackage, worksheetShown_2);

                //////////////////Статичная надпись
                writeWorksheetShown_2_7(excelPackage, worksheetShown_2, dateEnd);
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetShown_2_Name));
            }
        }

    }
}
