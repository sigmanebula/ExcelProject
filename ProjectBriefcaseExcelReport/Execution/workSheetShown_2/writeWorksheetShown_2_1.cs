using System;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void writeWorksheetShown_2_1(ExcelPackage excelPackage, ExcelWorksheet worksheet, string dateEnd)
        {
            string errorPrefix = "Структура портфеля технологических задач. ";
            int HeaderTableRowStart = 0;
            int HeaderTableColumnStart = 0;
            string worksheetShown_2_1_DataStartCell = "";
            
            try
            {
                //////////////////Структура портфеля технологических задач, получаем стартовую ячейку
                Settings.Variables.WorksheetShown_2_1_StartCell = ExcelAddress.GetAddress(
                      worksheet.Cells[Settings.Variables.WorksheetShown_2_3_4_EndCell].Start.Row + 2
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_3_4_StartCell].Start.Column
                    );
            }
            catch (Exception exception)
            {
                throw new Exception(errorPrefix
                    + "Не удалось получить стартовую ячейку"
                    + ", WorksheetShown_2_3_4_EndCell: " + Settings.Variables.WorksheetShown_2_3_4_EndCell
                    + ", WorksheetShown_2_3_4_StartCell: " + Settings.Variables.WorksheetShown_2_3_4_StartCell
                    + ", exception.Message: " + exception.Message
                    );
            }

            try
            {
                //пишем заголовок
                worksheet.Cells[Settings.Variables.WorksheetShown_2_1_StartCell].Value =
                    "Cтруктура портфеля технологических задач на "
                    + dateEnd;

            }
            catch (Exception exception)
            {
                throw new Exception(errorPrefix
                    + "Не удалось записать заголовок"
                    + ", WorksheetShown_2_1_StartCell: " + Settings.Variables.WorksheetShown_2_1_StartCell
                    + ", exception.Message: " + exception.Message
                    );
            }

            try
            {
                //мёржим строку заголовка
                worksheet.Cells[
                      worksheet.Cells[Settings.Variables.WorksheetShown_2_1_StartCell].Start.Row
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_1_StartCell].Start.Column
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_1_StartCell].Start.Row
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_1_StartCell].Start.Column + 3
                    ].Merge = true;
            }
            catch (Exception exception)
            {
                throw new Exception(errorPrefix
                    + "Не объединить ячейки заголовка"
                    + ", WorksheetShown_2_1_StartCell: " + Settings.Variables.WorksheetShown_2_1_StartCell
                    + ", exception.Message: " + exception.Message
                    );
            }

            try
            {
                //применяем стили к заголовку
                Settings.Variables.DefaultDataTableHeaderStyle.FillRange(worksheet, Settings.Variables.WorksheetShown_2_1_StartCell);
            }
            catch (Exception exception)
            {
                throw new Exception(errorPrefix
                    + "Не удалось применить стили к заголовку"
                    + ", WorksheetShown_2_1_StartCell: " + Settings.Variables.WorksheetShown_2_1_StartCell
                    + ", exception.Message: " + exception.Message
                    );
            }
            
            try
            {
                //применяем стили к заголовку
                Settings.Variables.DefaultDataTableHeaderStyle.FillRange(worksheet, Settings.Variables.WorksheetShown_2_1_StartCell);
            }
            catch (Exception exception)
            {
                throw new Exception(errorPrefix
                    + "Не удалось применить стили к заголовку"
                    + ", WorksheetShown_2_1_StartCell: " + Settings.Variables.WorksheetShown_2_1_StartCell
                    + ", exception.Message: " + exception.Message
                    );
            }

            try
            {
                //пишем колонки таблицы
                HeaderTableRowStart = worksheet.Cells[Settings.Variables.WorksheetShown_2_1_StartCell].Start.Row + 1;
                HeaderTableColumnStart = worksheet.Cells[Settings.Variables.WorksheetShown_2_1_StartCell].Start.Column;

                worksheet.Row(HeaderTableRowStart).Height = worksheet.Row(HeaderTableRowStart).Height * 3;
                worksheet.Cells[HeaderTableRowStart, HeaderTableColumnStart].Value = "Направление";
                worksheet.Cells[HeaderTableRowStart, HeaderTableColumnStart, HeaderTableRowStart, HeaderTableColumnStart + 1].Merge = true;
                worksheet.Cells[HeaderTableRowStart, HeaderTableColumnStart + 2].Value = "Количественные характеристики, проекты";
                worksheet.Cells[HeaderTableRowStart, HeaderTableColumnStart + 3].Value = "Ссылка на слайд с детальной информацией";
            }
            catch (Exception exception)
            {
                throw new Exception(errorPrefix
                    + "Не удалось записать колонки таблицы"
                    + ", WorksheetShown_2_1_StartCell: " + Settings.Variables.WorksheetShown_2_1_StartCell
                    + ", HeaderTableRowStart: " + HeaderTableRowStart.ToString()
                    + ", HeaderTableColumnStart: " + HeaderTableColumnStart.ToString()
                    + ", exception.Message: " + exception.Message
                    );
            }

            try
            {
                //применяем стили к заголовкам колонок
                Settings.Variables.DefaultDataTableColumnHeaderStyle.FillRange(
                      worksheet
                    , ExcelAddress.GetAddress(HeaderTableRowStart, HeaderTableColumnStart) + ":" + ExcelAddress.GetAddress(HeaderTableRowStart, HeaderTableColumnStart + 3)
                    );
            }
            catch (Exception exception)
            {
                throw new Exception(errorPrefix
                    + "Не удалось применить стили к заголовкам колонок"
                    + ", HeaderTableRowStart: " + HeaderTableRowStart.ToString()
                    + ", HeaderTableColumnStart: " + HeaderTableColumnStart.ToString()
                    + ", exception.Message: " + exception.Message
                    );
            }

            try
            {
                //вносим данные из другого листа
                worksheetShown_2_1_DataStartCell = ExcelAddress.GetAddress(HeaderTableRowStart + 1, HeaderTableColumnStart);

                Helpers.Excel.CopyDataFromAnotherWorkSheet(
                      Settings.Variables.WorksheetHidden_2_1_DataStartCell
                    , Settings.Variables.WorksheetHidden_2_1_DataEndCell
                    , ExcelAddress.GetAddress(HeaderTableRowStart + 1, HeaderTableColumnStart)
                    , Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_1_Name)
                    , worksheet
                    , true
                    );
            }
            catch (Exception exception)
            {
                throw new Exception(errorPrefix
                    + "Не удалось скопировать данные с другого листа"
                    + ", worksheetShown_2_1_DataStartCell: " + worksheetShown_2_1_DataStartCell
                    + ", HeaderTableRowStart: " + HeaderTableRowStart.ToString()
                    + ", HeaderTableColumnStart: " + HeaderTableColumnStart.ToString()
                    + ", WorksheetHidden_2_1_DataStartCell: " + Settings.Variables.WorksheetHidden_2_1_DataStartCell
                    + ", WorksheetHidden_2_1_DataEndCell: " + Settings.Variables.WorksheetHidden_2_1_DataEndCell
                    + ", exception.Message: " + exception.Message
                    );
            }
            
            try
            {
                //ищем край таблицы
                int[] worksheetShown_2_1_move = Helpers.Excel.GetCellDifference(worksheet, Settings.Variables.WorksheetHidden_2_1_DataStartCell, worksheetShown_2_1_DataStartCell);

                Settings.Variables.WorksheetShown_2_1_EndCell = Helpers.Excel.GetLocationNewMoveCell(
                      worksheet
                    , Settings.Variables.WorksheetHidden_2_1_DataEndCell
                    , worksheetShown_2_1_move[0]
                    , worksheetShown_2_1_move[1] + 1    //+колонка под ссылки
                    );
            }
            catch (Exception exception)
            {
                throw new Exception(errorPrefix
                    + "Не удалось найти край таблицы"
                    + ", WorksheetHidden_2_1_DataStartCell: " + Settings.Variables.WorksheetHidden_2_1_DataStartCell
                    + ", worksheetShown_2_1_DataStartCell: " + worksheetShown_2_1_DataStartCell
                    + ", exception.Message: " + exception.Message
                    );
            }
            
            try
            {
                //применяем стили к таблице
                Settings.Variables.DefaultDataTableCellStyle.FillRange(
                      worksheet
                    , worksheetShown_2_1_DataStartCell + ":" + Settings.Variables.WorksheetShown_2_1_EndCell
                    );
            }
            catch (Exception exception)
            {
                throw new Exception(errorPrefix
                    + "Не удалось применить стили"
                    + ", worksheetShown_2_1_DataStartCell: " + worksheetShown_2_1_DataStartCell
                    + ", WorksheetShown_2_1_EndCell: " + Settings.Variables.WorksheetShown_2_1_EndCell
                    + ", exception.Message: " + exception.Message
                    );
            }
        }
    }
}
