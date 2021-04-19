using System;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void writeWorksheetShown_3_Label(ExcelPackage excelPackage, ExcelWorksheet worksheet)
        {
            try
            {
                //////////////////На КРБИ утверждены изменения 
                worksheet.Cells[Settings.SQLVariables.WorksheetShown_3_LabelStartCell].Value
                    = Convert.ToString(worksheet.Cells[Settings.SQLVariables.WorksheetShown_3_LabelStartCell].Value ?? "")
                    + Settings.Variables.WorksheetHidden_3_1_ProjectCount.ToString()
                    + " проектов в части следующих ключевых параметров:";
                //объединяем ячейки
                worksheet.Cells[
                      worksheet.Cells[Settings.SQLVariables.WorksheetShown_3_LabelStartCell].Start.Row
                    , worksheet.Cells[Settings.SQLVariables.WorksheetShown_3_LabelStartCell].Start.Column
                    , worksheet.Cells[Settings.SQLVariables.WorksheetShown_3_LabelStartCell].Start.Row
                    , worksheet.Cells[Settings.SQLVariables.WorksheetShown_3_LabelStartCell].Start.Column + 2
                    ].Merge = true;
                //применяем стили к заголовоку КРБИ
                Settings.Variables.DefaultDataTableHeaderStyle.FillRange(worksheet, Settings.SQLVariables.WorksheetShown_3_LabelStartCell);
            }
            catch (Exception ex)
            {
                throw new Exception("Заголовок круговой диаграммы: " + ex.Message);
            }
        }
    }
}
