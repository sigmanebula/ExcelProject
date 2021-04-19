using System;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void writeWorksheetShown_3_Header(ExcelPackage excelPackage, ExcelWorksheet worksheet, string dateStart, string dateEnd)
        {
            try
            {
                //////////////////Главный заголовок
                worksheet.Cells[Settings.SQLVariables.WorksheetShown_3_Header_LabelStartCell].Value =
                      Convert.ToString(worksheet.Cells[Settings.SQLVariables.WorksheetShown_3_Header_LabelStartCell].Value ?? "")
                    + dateStart
                    + " - "
                    + dateEnd;

                //объединяем ячейки
                worksheet.Cells[
                      worksheet.Cells[Settings.SQLVariables.WorksheetShown_3_Header_LabelStartCell].Start.Row
                    , worksheet.Cells[Settings.SQLVariables.WorksheetShown_3_Header_LabelStartCell].Start.Column
                    , worksheet.Cells[Settings.SQLVariables.WorksheetShown_3_Header_LabelStartCell].Start.Row
                    , worksheet.Cells[Settings.SQLVariables.WorksheetShown_3_Header_LabelStartCell].Start.Column + 2
                    ].Merge = true;

                //применяем стили к Главному заголовку
                Settings.Variables.DefaultDataTableHeaderStyle.FillRange(worksheet, Settings.SQLVariables.WorksheetShown_3_Header_LabelStartCell);
            }
            catch (Exception ex)
            {
                throw new Exception("Главный заголовок: " + ex.Message);
            }
        }
    }
}
