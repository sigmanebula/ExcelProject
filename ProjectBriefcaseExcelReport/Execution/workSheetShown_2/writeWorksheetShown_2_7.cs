using System;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void writeWorksheetShown_2_7(ExcelPackage excelPackage, ExcelWorksheet worksheet, string dateEnd)
        {
            try
            {
                //////////////////Статичная надпись
                Settings.Variables.WorksheetShown_2_7_StartCell = ExcelAddress.GetAddress(
                      worksheet.Cells[Settings.Variables.WorksheetShown_2_6_EndCell].Start.Row + 2
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_6_StartCell].Start.Column
                    );

                Settings.Variables.WorksheetShown_2_7_EndCell = ExcelAddress.GetAddress(
                      worksheet.Cells[Settings.Variables.WorksheetShown_2_7_StartCell].Start.Row + 1
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_7_StartCell].Start.Column + 2
                    );

                //мёржим строки 1
                worksheet.Cells[
                      worksheet.Cells[Settings.Variables.WorksheetShown_2_7_StartCell].Start.Row
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_7_StartCell].Start.Column
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_7_StartCell].Start.Row
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_7_EndCell].Start.Column
                    ].Merge = true;

                //мёржим строки 2
                worksheet.Cells[
                      worksheet.Cells[Settings.Variables.WorksheetShown_2_7_EndCell].Start.Row
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_7_StartCell].Start.Column
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_7_EndCell].Start.Row
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_7_EndCell].Start.Column
                    ].Merge = true;

                string LabelFirst
                    = "* В статистике отражены изменения без учёта вопросов, вынесенных на рассмотрение КРБИ "
                    + dateEnd;

                string nextQuarter = Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeEndNext, "Quarter");
                string nextYear = Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeEndNext, "Year");
                if (nextQuarter == "")
                {
                    nextQuarter = Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeStartNext, "Quarter");
                    nextYear = Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeStartNext, "Year");
                }

                worksheet.Row(worksheet.Cells[Settings.Variables.WorksheetShown_2_7_StartCell].Start.Row).Height = worksheet.Row(worksheet.Cells[Settings.Variables.WorksheetShown_2_7_StartCell].Start.Row).Height * 2.2;
                worksheet.Row(worksheet.Cells[Settings.Variables.WorksheetShown_2_7_EndCell].Start.Row).Height = worksheet.Row(worksheet.Cells[Settings.Variables.WorksheetShown_2_7_EndCell].Start.Row).Height * 2.2;

                //пишем строку 1
                worksheet.Cells[Settings.Variables.WorksheetShown_2_7_StartCell].Value
                    = "* В статистике отражены изменения без учёта вопросов, вынесенных на рассмотрение КРБИ "
                    + dateEnd;

                //пишем строку 2
                worksheet.Cells[ExcelAddress.GetAddress(
                      worksheet.Cells[Settings.Variables.WorksheetShown_2_7_EndCell].Start.Row
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_7_StartCell].Start.Column)].Value
                        = "* Результаты рассмотрения данных вопросов будут включены в отчет по портфелю за "
                        + nextQuarter + " квартал " + nextYear + " года";

                //применяем стили к заголовкам
                Settings.Variables.DefaultDataTableHeaderStyle.FillRange(worksheet, Settings.Variables.WorksheetShown_2_7_StartCell + ":" + Settings.Variables.WorksheetShown_2_7_EndCell);

            }
            catch (Exception ex)
            {
                throw new Exception("Нижняя надпись: " + ex.Message);
            }
        }
    }
}

