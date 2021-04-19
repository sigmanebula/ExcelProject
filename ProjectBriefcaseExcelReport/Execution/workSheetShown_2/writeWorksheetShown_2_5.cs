using System;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void writeWorksheetShown_2_5(ExcelPackage excelPackage, ExcelWorksheet worksheet)
        {
            try
            {
                //////////////////Исключены из портфеля
                Settings.Variables.WorksheetShown_2_5_StartCell = Settings.SQLVariables.WorksheetShown_2_5_LabelStartCell;

                //пишем заголовок
                worksheet.Cells[Settings.Variables.WorksheetShown_2_5_StartCell].Value =
                    "Исключены из портфеля проектов " + Settings.Variables.WorksheetHidden_2_5_ProjectCount.ToString() + " проектных инициатив";

                //мёржим строку заголовка
                worksheet.Cells[
                      worksheet.Cells[Settings.Variables.WorksheetShown_2_5_StartCell].Start.Row
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_5_StartCell].Start.Column
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_5_StartCell].Start.Row
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_5_StartCell].Start.Column + 2
                    ].Merge = true;

                //применяем стили к заголовку
                Settings.Variables.DefaultDataTableHeaderStyle.FillRange(worksheet, Settings.Variables.WorksheetShown_2_5_StartCell);

                //пишем колонки таблицы
                int HeaderTableRowStart = worksheet.Cells[Settings.Variables.WorksheetShown_2_5_StartCell].Start.Row + 1;
                int HeaderTableColumnStart = worksheet.Cells[Settings.Variables.WorksheetShown_2_5_StartCell].Start.Column;

                worksheet.Cells[HeaderTableRowStart, HeaderTableColumnStart].Value = "№";
                worksheet.Cells[HeaderTableRowStart, HeaderTableColumnStart + 1].Value = "Формат";
                worksheet.Cells[HeaderTableRowStart, HeaderTableColumnStart + 2].Value = "Название";

                //применяем стили к заголовкам колонок
                Settings.Variables.DefaultDataTableColumnHeaderStyle.FillRange(worksheet
                    , ExcelAddress.GetAddress(HeaderTableRowStart, HeaderTableColumnStart) + ":" + ExcelAddress.GetAddress(HeaderTableRowStart, HeaderTableColumnStart + 2));

                if (Settings.Variables.WorksheetHidden_2_5_ProjectCount > 0)
                {
                    //вставляем нумерацию
                    for (int i = HeaderTableRowStart + 1, number = 1; i <= HeaderTableRowStart + Settings.Variables.WorksheetHidden_2_5_ProjectCount; i++, number++)
                        worksheet.Cells[i, HeaderTableColumnStart].Value = number;

                    //вставляем колонку формат с другого листа
                    int column_format = Helpers.Excel.GetCellByValue(Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_5_Name), "ProjectTypeName").Start.Column;
                    Helpers.Excel.CopyDataFromAnotherWorkSheet(
                          2
                        , column_format
                        , 1 + Settings.Variables.WorksheetHidden_2_5_ProjectCount
                        , column_format
                        , HeaderTableRowStart + 1
                        , HeaderTableColumnStart + 1
                        , Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_5_Name)
                        , worksheet
                        , true
                        );

                    //вставляем колонку формат с другого листа
                    int column_projectName = Helpers.Excel.GetCellByValue(Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_5_Name), "ProjectName").Start.Column;
                    Helpers.Excel.CopyDataFromAnotherWorkSheet(
                          2
                        , column_projectName
                        , 1 + Settings.Variables.WorksheetHidden_2_5_ProjectCount
                        , column_projectName
                        , HeaderTableRowStart + 1
                        , HeaderTableColumnStart + 2
                        , Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_5_Name)
                        , worksheet
                        , true
                        );

                    //применяем стили к таблице
                    Settings.Variables.DefaultDataTableCellStyle.FillRange(
                          worksheet
                        , HeaderTableRowStart + 1
                        , HeaderTableColumnStart
                        , HeaderTableRowStart + Settings.Variables.WorksheetHidden_2_5_ProjectCount
                        , HeaderTableColumnStart + 2
                        );
                }

                Settings.Variables.WorksheetShown_2_5_EndCell = ExcelAddress.GetAddress(
                      worksheet.Cells[Settings.Variables.WorksheetShown_2_5_StartCell].Start.Row + 1 + Settings.Variables.WorksheetHidden_2_5_ProjectCount
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_5_StartCell].Start.Column + 2
                    );

            }
            catch (Exception ex)
            {
                throw new Exception("Исключены из портфеля: " + ex.Message);
            }
        }

    }
}
