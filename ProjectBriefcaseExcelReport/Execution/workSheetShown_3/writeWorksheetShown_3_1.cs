using System;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {

        static void writeWorksheetShown_3_1(ExcelPackage excelPackage, ExcelWorksheet worksheet, string dateStart, string dateEnd)
        {
            try
            {
                //////////////////Исключены из портфеля
                Settings.Variables.WorksheetShown_3_1_StartCell = Settings.SQLVariables.WorksheetShown_3_1_TableStartCell;

                //пишем колонки таблицы
                int HeaderTableRowStart = worksheet.Cells[Settings.Variables.WorksheetShown_3_1_StartCell].Start.Row;
                int HeaderTableColumnStart = worksheet.Cells[Settings.Variables.WorksheetShown_3_1_StartCell].Start.Column;

                worksheet.Cells[HeaderTableRowStart, HeaderTableColumnStart].Value = "Параметр";
                worksheet.Cells[HeaderTableRowStart, HeaderTableColumnStart + 1].Value = "Количество проектов, изменивших параметр";

                //применяем стили к заголовкам колонок
                Settings.Variables.DefaultDataTableColumnHeaderStyle.FillRange(worksheet
                    , ExcelAddress.GetAddress(HeaderTableRowStart, HeaderTableColumnStart) + ":" + ExcelAddress.GetAddress(HeaderTableRowStart, HeaderTableColumnStart + 1));

                if (Settings.Variables.WorksheetHidden_3_1_CountLine > 0)
                {
                    //вставляем с другого листа
                    int column_format = Helpers.Excel.GetCellByValue(Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_3_1_Name), "Parametrs").Start.Column;
                    Helpers.Excel.CopyDataFromAnotherWorkSheet(
                          2
                        , column_format
                        , 1 + Settings.Variables.WorksheetHidden_3_1_CountLine
                        , column_format + 1
                        , HeaderTableRowStart + 1
                        , HeaderTableColumnStart
                        , Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_3_1_Name)
                        , worksheet
                        , true
                        );

                    //применяем стили к таблице
                    Settings.Variables.DefaultDataTableCellStyle.FillRange(
                          worksheet
                        , HeaderTableRowStart + 1
                        , HeaderTableColumnStart
                        , HeaderTableRowStart + Settings.Variables.WorksheetHidden_3_1_CountLine
                        , HeaderTableColumnStart + 1
                        );
                }

                //пишем колонки для текста после таблицы
                int LabelRowStart = HeaderTableRowStart + 1 + Settings.Variables.WorksheetHidden_3_1_CountLine;
                int LabelColumnStart = HeaderTableColumnStart;

                Settings.Variables.WorksheetShown_3_1_EndCell = ExcelAddress.GetAddress(LabelRowStart, HeaderTableColumnStart + 1);

                LabelRowStart++;
                //пишем заголовок №2
                worksheet.Row(LabelRowStart).Height = worksheet.Row(LabelRowStart).Height * 2;

                worksheet.Cells[LabelRowStart, LabelColumnStart].Value =
                    "В статистике отражены изменения без учёта запросов, утверждаемых в дату рассмотрения отчета по портфелю проектов за "
                    + dateStart
                    + " - "
                    + dateEnd;

                //мёржим строку заголовок №2
                worksheet.Cells[
                      LabelRowStart
                    , LabelColumnStart
                    , LabelRowStart
                    , LabelColumnStart + 2
                    ].Merge = true;

                //применяем стили к заголовоку №2
                Settings.Variables.DefaultDataTableHeaderStyle.FillRange(worksheet, ExcelAddress.GetAddress(LabelRowStart, LabelColumnStart));
                LabelRowStart++;

                //пишем заголовок №3
                worksheet.Cells[LabelRowStart, LabelColumnStart].Value = "Данные запросы будут включены в следующий отчет.";

                //мёржим строку заголовок №3
                worksheet.Cells[
                      LabelRowStart
                    , LabelColumnStart
                    , LabelRowStart
                    , LabelColumnStart + 2
                    ].Merge = true;

                //применяем стили к заголовоку №3
                Settings.Variables.DefaultDataTableHeaderStyle.FillRange(worksheet, ExcelAddress.GetAddress(LabelRowStart, LabelColumnStart));
                LabelRowStart += 2;

                //пишем заголовок №3
                worksheet.Cells[LabelRowStart, LabelColumnStart].Value =
                     Settings.Variables.WorksheetHidden_3_1_WaitCount
                     + " - количество запросов на изменение бюджета утверждены к включению в Лист ожидания решения в части финансирования (см. Приложение 1).";

                //мёржим строку заголовок №3
                worksheet.Cells[
                      LabelRowStart
                    , LabelColumnStart
                    , LabelRowStart
                    , LabelColumnStart + 2
                    ].Merge = true;

                //применяем стили к заголовоку №3
                Settings.Variables.DefaultDataTableHeaderStyle.FillRange(worksheet, ExcelAddress.GetAddress(LabelRowStart, LabelColumnStart));
                LabelRowStart += 2;

                //пишем заголовок №4
                worksheet.Cells[LabelRowStart, LabelColumnStart].Value =
                    "Детальная информация о параметрах всех "
                    + Settings.Variables.WorksheetHidden_3_1_ProjectCount.ToString()
                    + " запросов на изменение приведена в Приложении 2.";

                //мёржим строку заголовок №4
                worksheet.Cells[
                      LabelRowStart
                    , LabelColumnStart
                    , LabelRowStart
                    , LabelColumnStart + 2
                    ].Merge = true;

                //применяем стили к заголовоку №4
                Settings.Variables.DefaultDataTableHeaderStyle.FillRange(worksheet, ExcelAddress.GetAddress(LabelRowStart, LabelColumnStart));

                /*
                В статистике отражены изменения без учёта запросов, утверждаемых в дату рассмотрения отчета по портфелю проектов за 
                Данные запросы будут включены в следующий отчет.

                Х запроса на изменение бюджета утверждены к включению в Лист ожидания решения в части финансирования (см. Приложение 1). (Х - оставить символ "Х" для ручной корректировки, т.к. в системе не реализован учет подобных случаев)

                Детальная информация о параметрах всех Х запросов на изменение приведена в Приложении 2. (Х – указывается количество проектов, по которым были созданы новые версии карточек за указанный период и место принятия решения - "КРБИ")
                */

            }
            catch (Exception ex)
            {
                throw new Exception("Таблица изменения проектов: " + ex.Message);
            }
        }
    }
}
