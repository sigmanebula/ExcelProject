using OfficeOpenXml;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace ProjectBriefcaseExcelReport
{
    public static class Execution
    {
        static string getPeriodName(bool isLowerQuarterWord, ref string errorText)
        {
            string result = "";

            if (errorText == "")
                try
                {
                    string quarterWord = (isLowerQuarterWord) ? " квартал " : " КВАРТАЛ ";

                    string startWord =
                          Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeStart, "Quarter")
                        + quarterWord
                        + Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeStart, "Year");

                    string endWord =
                          Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeEnd, "Quarter")
                        + quarterWord
                        + Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeEnd, "Year");

                    if (startWord == endWord)
                        result += startWord;
                    else
                        result += startWord + " - " + endWord;
                }
                catch (Exception ex)
                {
                    errorText += "Ошибка получения периода. " + ex.Message;
                }

            return result;
        }
        
        static string getPeriodName(bool isLowerQuarterWord)
        {
            string errorText = "";
            string result = getPeriodName(isLowerQuarterWord, ref errorText);

            if (errorText != "")
                throw new System.Exception(errorText);

            return result;
        }


        static void getProductionCalendar(string dateStart, string dateEnd, SqlConnection connection, ref string errorText)
        {
            if (errorText == "")
                try
                {
                    using (var cmd = new SqlCommand()) //записываем
                    {
                        cmd.Connection = connection;
                        cmd.CommandText = "[ITProject].[spGetExcelReportProductionCalendar]";
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                        cmd.Parameters.AddWithValue("@StartDate", dateStart);
                        cmd.Parameters.AddWithValue("@EndDate", dateEnd);
                        cmd.ExecuteNonQuery();
                        var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                        var dataSet = new DataSet();
                        dataAdapter.Fill(dataSet);
                        Settings.Variables.ProductionCalendar = dataSet.Tables[0];

                        foreach (DataRow row in Settings.Variables.ProductionCalendar.Rows)
                            row["Date"] = Convert.ToString(row["Date"] ?? "").Replace("Z", "");
                    }
                }
                catch (Exception ex)
                {
                    errorText += "Не удалось получить данные календаря, причина: " + ex.Message;
                }
        }

        static void workSheetShown_2_Write(ExcelPackage excelPackage, string dateStart, string dateEnd)
        {
            var worksheetShown_2 = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetShown_2_Name);
            try
            {
                dateStart = Helpers.Sugar.ConvertDateToFormat(dateStart, '-', '.', "YYYY.MM.DD", "DD.MM.YYYY");
                dateEnd = Helpers.Sugar.ConvertDateToFormat(dateEnd, '-', '.', "YYYY.MM.DD", "DD.MM.YYYY");

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

        static void writeWorksheetShown_2_2_Header(ExcelPackage excelPackage, ExcelWorksheet worksheet, string dateStart, string dateEnd)
        {
            try
            {
                //////////////////Заголовок столбчатой диаграммы
                worksheet.Cells[Settings.SQLVariables.WorksheetShown_2_2_Header_LabelStartCell].Value
                    = Convert.ToString(worksheet.Cells[Settings.SQLVariables.WorksheetShown_2_2_Header_LabelStartCell].Value ?? "")
                    + " "
                    + dateStart
                    + " - "
                    + dateEnd;
            }
            catch (Exception ex)
            {
                throw new Exception("Заголовок столбчатой диаграммы: " + ex.Message);
            }
        }


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

        static void writeWorksheetShown_2_3_4(ExcelWorksheet worksheet)
        {
            try
            {
                //определяем начало
                Settings.Variables.WorksheetShown_2_3_4_StartCell = Settings.SQLVariables.WorksheetShown_2_3_4_LabelStartCell;

                //////////////////Здоровье динамических и водопадных проектов
                //пишем заголовок
                worksheet.Cells[Settings.Variables.WorksheetShown_2_3_4_StartCell].Value =
                    "В оценке участвовало " + Settings.Variables.WorksheetHidden_2_3_ProjectCount.ToString() + " проектов, в том числе закрытые на КРБИ:";

                //мёржим строку заголовка
                worksheet.Cells[
                      worksheet.Cells[Settings.Variables.WorksheetShown_2_3_4_StartCell].Start.Row
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_3_4_StartCell].Start.Column
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_3_4_StartCell].Start.Row
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_3_4_StartCell].Start.Column + int.Parse(Settings.SQLVariables.WorksheetShown_2_3_4_MergeCellCount) - 1
                    ].Merge = true;

                //применяем стили к заголовку
                Settings.Variables.DefaultDataTableHeaderStyle.FillRange(worksheet, Settings.Variables.WorksheetShown_2_3_4_StartCell);

                if (Settings.Variables.WorksheetHidden_2_4_ProjectCount > 0)
                {
                    int startRowTo = worksheet.Cells[Settings.Variables.WorksheetShown_2_3_4_StartCell].Start.Row + 1;
                    int startColumnTo = worksheet.Cells[Settings.Variables.WorksheetShown_2_3_4_StartCell].Start.Column;

                    int startRowFrom = 2;
                    int startColumnFrom = 2;
                    int differenceRowTo = startRowTo - startRowFrom;

                    string formula = "=CONCATENATE(\" * \",{0}!{1},\" - оценка \",IF(OR(ISBLANK({0}!{2}),{0}!{2}=\"\"),\"не определена\",{0}!{2}))";

                    for (int i = startRowFrom; i <= Settings.Variables.WorksheetHidden_2_4_ProjectCount + 1; i++)
                    {
                        worksheet.Cells[i + differenceRowTo, startColumnTo].Formula = String.Format(
                            formula
                            , Settings.SQLVariables.WorksheetHidden_2_4_Name
                            , ExcelAddress.GetAddress(i, startColumnFrom)
                            , ExcelAddress.GetAddress(i, startColumnFrom + 1)
                            );

                        worksheet.Cells[
                              i + differenceRowTo
                            , startColumnTo
                            , i + differenceRowTo
                            , startColumnTo + int.Parse(Settings.SQLVariables.WorksheetShown_2_3_4_MergeCellCount) - 1
                            ].Merge = true;
                    }

                    //применяем стили
                    Settings.Variables.DefaultLabelStyle.FillRange(
                          worksheet
                        , ExcelAddress.GetAddress(startRowTo, startColumnTo) + ":" + ExcelAddress.GetAddress(startRowTo + Settings.Variables.WorksheetHidden_2_4_ProjectCount, startColumnTo)
                        );
                }

                //находим конец
                Settings.Variables.WorksheetShown_2_3_4_EndCell = ExcelAddress.GetAddress(
                      worksheet.Cells[Settings.Variables.WorksheetShown_2_3_4_StartCell].Start.Row + Settings.Variables.WorksheetHidden_2_4_ProjectCount
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_3_4_StartCell].Start.Column + int.Parse(Settings.SQLVariables.WorksheetShown_2_3_4_MergeCellCount) - 1
                    );
            }
            catch (Exception ex)
            {
                throw new Exception("Здоровье динамических и водопадных проектов: " + ex.Message);
            }
        }

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


        static void writeWorksheetShown_2_6(ExcelPackage excelPackage, ExcelWorksheet worksheet)
        {
            try
            {
                //////////////////Включены в портфель
                Settings.Variables.WorksheetShown_2_6_StartCell = ExcelAddress.GetAddress(
                      worksheet.Cells[Settings.Variables.WorksheetShown_2_5_EndCell].Start.Row + 2
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_5_StartCell].Start.Column
                    );

                //пишем заголовок
                worksheet.Cells[Settings.Variables.WorksheetShown_2_6_StartCell].Value =
                    "Включены в портфель " + Settings.Variables.WorksheetHidden_2_6_ProjectCount.ToString() + " проектных инициатив";

                //мёржим строку заголовка
                worksheet.Cells[
                      worksheet.Cells[Settings.Variables.WorksheetShown_2_6_StartCell].Start.Row
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_6_StartCell].Start.Column
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_6_StartCell].Start.Row
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_6_StartCell].Start.Column + 2
                    ].Merge = true;

                //применяем стили к заголовку
                Settings.Variables.DefaultDataTableHeaderStyle.FillRange(worksheet, Settings.Variables.WorksheetShown_2_6_StartCell);

                //пишем колонки таблицы
                int HeaderTableRowStart = worksheet.Cells[Settings.Variables.WorksheetShown_2_6_StartCell].Start.Row + 1;
                int HeaderTableColumnStart = worksheet.Cells[Settings.Variables.WorksheetShown_2_6_StartCell].Start.Column;

                //worksheet.Column(HeaderTableColumnStart + 1).Width = worksheet.Column(HeaderTableColumnStart + 1).Width * 2;
                //worksheet.Column(HeaderTableColumnStart + 2).Width = worksheet.Column(HeaderTableColumnStart + 2).Width * 1.5;  //считается смёрженная ячейка целиком, а не первая

                worksheet.Cells[HeaderTableRowStart, HeaderTableColumnStart].Value = "№";
                worksheet.Cells[HeaderTableRowStart, HeaderTableColumnStart + 1].Value = "Формат";
                worksheet.Cells[HeaderTableRowStart, HeaderTableColumnStart + 2].Value = "Название";

                //применяем стили к заголовкам колонок
                Settings.Variables.DefaultDataTableColumnHeaderStyle.FillRange(worksheet
                    , ExcelAddress.GetAddress(HeaderTableRowStart, HeaderTableColumnStart) + ":" + ExcelAddress.GetAddress(HeaderTableRowStart, HeaderTableColumnStart + 2));

                if (Settings.Variables.WorksheetHidden_2_6_ProjectCount > 0)
                {
                    //вставляем нумерацию
                    for (int i = HeaderTableRowStart + 1, number = 1; i <= HeaderTableRowStart + Settings.Variables.WorksheetHidden_2_6_ProjectCount; i++, number++)
                        worksheet.Cells[i, HeaderTableColumnStart].Value = number;

                    //вставляем колонку формат с другого листа
                    int column_format = Helpers.Excel.GetCellByValue(Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_6_Name), "ProjectTypeName").Start.Column;
                    Helpers.Excel.CopyDataFromAnotherWorkSheet(
                          2
                        , column_format
                        , 1 + Settings.Variables.WorksheetHidden_2_6_ProjectCount
                        , column_format
                        , HeaderTableRowStart + 1
                        , HeaderTableColumnStart + 1
                        , Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_6_Name)
                        , worksheet
                        , true
                        );

                    //вставляем колонку формат с другого листа
                    int column_projectName = Helpers.Excel.GetCellByValue(Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_6_Name), "ProjectName").Start.Column;
                    Helpers.Excel.CopyDataFromAnotherWorkSheet(
                          2
                        , column_projectName
                        , 1 + Settings.Variables.WorksheetHidden_2_6_ProjectCount
                        , column_projectName
                        , HeaderTableRowStart + 1
                        , HeaderTableColumnStart + 2
                        , Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_6_Name)
                        , worksheet
                        , true
                        );

                    //применяем стили к таблице
                    Settings.Variables.DefaultDataTableCellStyle.FillRange(
                          worksheet
                        , HeaderTableRowStart + 1
                        , HeaderTableColumnStart
                        , HeaderTableRowStart + Settings.Variables.WorksheetHidden_2_6_ProjectCount
                        , HeaderTableColumnStart + 2
                        );
                }

                Settings.Variables.WorksheetShown_2_6_EndCell = ExcelAddress.GetAddress(
                      worksheet.Cells[Settings.Variables.WorksheetShown_2_6_StartCell].Start.Row + 1 + Settings.Variables.WorksheetHidden_2_6_ProjectCount
                    , worksheet.Cells[Settings.Variables.WorksheetShown_2_6_StartCell].Start.Column + 2
                    );
            }
            catch (Exception ex)
            {
                throw new Exception("Включены в портфель: " + ex.Message);
            }
        }


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


        static void writeWorksheetShown_2_Header(ExcelPackage excelPackage, ExcelWorksheet worksheet, string dateStart, string dateEnd)
        {
            try
            {
                //////////////////Главный заголовок
                worksheet.Cells[Settings.SQLVariables.WorksheetShown_2_Header_LabelStartCell].Value
                    = Convert.ToString(worksheet.Cells[Settings.SQLVariables.WorksheetShown_2_Header_LabelStartCell].Value ?? "")
                    + " "
                    + dateStart
                    + " - "
                    + dateEnd;
            }
            catch (Exception ex)
            {
                throw new Exception("Главный заголовок: " + ex.Message);
            }
        }


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


        public static Helpers.ReturnClass GetFromSQLToFile(string dateStart, string dateEnd, string projectTypeIdList, string stateIdList)
        {
            projectTypeIdList = (projectTypeIdList ?? "").ToString(); //удалить после реализации на UI
            stateIdList = (stateIdList ?? "").ToString(); //удалить после реализации на UI

            dateStart = dateStart.Split(' ')[0];
            dateEnd = dateEnd.Split(' ')[0];

            Settings.SQLVariables = new SQLVariablesClass();
            Settings.Variables = new VariablesClass();
            Settings.Variables.Refresh();
            string errorText = "";

            using (var connection = Helpers.SugarSQLConnection.GetSQLConnection())
            {
                Helpers.SugarSQLConnection.OpenInUsing(connection);

                Settings.SQLVariables.GetSettings(connection, Settings.SettingsTypeCodeList, ref errorText);    //получаем настройки

                getProductionCalendar(dateStart, dateEnd, connection, ref errorText);  //получаем данные календаря

                string fileShortNameNew = Settings.SQLVariables.NewFileNamePrefix + getPeriodName(true, ref errorText) + "." + Settings.FileExtention; //тут получаем название нового файла

                string fileFullNameNew = Settings.SQLVariables.FolderPath + fileShortNameNew;

                string fileData = "";

                Helpers.SugarFile.Copy(Settings.SQLVariables.FolderPath + Settings.SQLVariables.TemplateFileShortName, fileFullNameNew, ref errorText); //копируем файл

                if (errorText == "")  //основное действие
                {
                    string productionCalendarIDStart = Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeStart, "ProductionCalendarID");
                    string productionCalendarIDEnd = Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeEnd, "ProductionCalendarID");
                    int startPeriodNumber = int.Parse(Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeStart, "PeriodNumber"));
                    int endPeriodNumber = int.Parse(Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeEnd, "PeriodNumber"));

                    ExcelPackage excelPackage = null;
                    try
                    {
                        excelPackage = new ExcelPackage(new FileInfo(fileFullNameNew));
                        if (excelPackage == null)
                            throw new Exception("\nПустой excelPackage");

                        worksheetHidden_Debug_Write(excelPackage, connection, dateStart, dateEnd, stateIdList, projectTypeIdList);
                        workSheetShown_1_Write(excelPackage, dateEnd);
                        worksheetHidden_2_1_Write(excelPackage, connection, dateEnd, stateIdList, projectTypeIdList);
                        worksheetHidden_2_2_Write(excelPackage, connection, dateStart, dateEnd, stateIdList, projectTypeIdList);
                        worksheetHidden_2_3_Write(excelPackage, connection, dateStart, dateEnd, productionCalendarIDStart, productionCalendarIDEnd);
                        worksheetHidden_2_4_Write(excelPackage, connection, dateStart, dateEnd, productionCalendarIDStart, productionCalendarIDEnd);
                        worksheetHidden_2_5_Write(excelPackage, connection, dateStart, dateEnd, stateIdList, projectTypeIdList);
                        worksheetHidden_2_6_Write(excelPackage, connection, dateStart, dateEnd, stateIdList, projectTypeIdList);
                        workSheetShown_2_Write(excelPackage, dateStart, dateEnd);
                        worksheetHidden_3_0_Write(excelPackage, connection, dateStart, dateEnd, stateIdList, projectTypeIdList);
                        worksheetHidden_3_1_Write(excelPackage, connection, dateStart, dateEnd, stateIdList, projectTypeIdList);
                        workSheetShown_3_Write(excelPackage, dateStart, dateEnd);
                        worksheetHidden_4_0_Write(excelPackage, connection, dateStart, dateEnd, startPeriodNumber, endPeriodNumber, stateIdList, projectTypeIdList);

                        excelPackage.Save();
                    }
                    catch (Exception ex)
                    {
                        errorText += "\nОшибка в файле. " + ex.Message;
                    }
                    finally
                    {
                        excelPackage.Dispose();
                    }

                    fileData = Helpers.SugarFile.GetK2Xml(fileFullNameNew, fileShortNameNew, "", ref errorText);
                }

                string userMessage = Settings.Variables.UserMessage;
                bool isGetErrorMessage = Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsGetErrorMessage);

                Settings.SQLVariables = new SQLVariablesClass();
                Settings.Variables = new VariablesClass();
                connection.Close();

                Helpers.SugarFile.DeleteIfExists(fileFullNameNew, "\nОшибка при удалении временного файла: ", ref errorText);

                userMessage = Helpers.Sugar.GetUserMessageAndErrorText(userMessage, errorText, isGetErrorMessage);

                GC.Collect();
                return new Helpers.ReturnClass() { FileData = fileData, UserMessage = userMessage };
            }
        }


        static void worksheetHidden_2_1_Write(ExcelPackage excelPackage, SqlConnection connection, string dateEnd, string stateIdList, string projectTypeIdList)
        {
            //////////////////Структура портфеля технологических задач
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_1_Name);

            try
            {
                DataTable dataTable = new DataTable();
                using (var cmd = new SqlCommand())
                {
                    cmd.Connection = connection;
                    cmd.CommandText = "[ITProject].[spGetExcelReportBriefcaseStructure]";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                    cmd.Parameters.AddWithValue("@EndDateString", dateEnd);
                    cmd.Parameters.AddWithValue("@ProjectStateIDList", stateIdList);
                    cmd.Parameters.AddWithValue("@ProjectTypeIDList", projectTypeIdList);
                    cmd.ExecuteNonQuery();
                    var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                    var dataSet = new DataSet();
                    dataAdapter.Fill(dataSet);
                    dataTable = dataSet.Tables[0];
                }

                Helpers.Excel.WriteDataTableToWorkSheet(dataTable, worksheet);

                Settings.Variables.WorksheetHidden_2_1_DataStartCell = ExcelAddress.GetAddress(2, 1);
                Settings.Variables.WorksheetHidden_2_1_DataEndCell = ExcelAddress.GetAddress(dataTable.Rows.Count + 1, dataTable.Columns.Count);
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetHidden_2_1_Name));
            }
        }

        static void worksheetHidden_2_2_Write(ExcelPackage excelPackage, SqlConnection connection, string dateStart, string dateEnd, string stateIdList, string projectTypeIdList)
        {
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_2_Name);

            try
            {
                DataTable dataTable = new DataTable();
                using (var cmd = new SqlCommand())
                {
                    cmd.Connection = connection;
                    cmd.CommandText = "[ITProject].[spGetExcelReportBriefcaseDynamic]";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                    cmd.Parameters.AddWithValue("@StartDateString", dateStart);
                    cmd.Parameters.AddWithValue("@EndDateString", dateEnd);
                    cmd.Parameters.AddWithValue("@ProjectStateIDList", stateIdList);
                    cmd.Parameters.AddWithValue("@ProjectTypeIDList", projectTypeIdList);
                    cmd.ExecuteNonQuery();
                    var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                    var dataSet = new DataSet();
                    dataAdapter.Fill(dataSet);
                    dataTable = dataSet.Tables[0];
                }

                Helpers.Excel.WriteDataTableToWorkSheet(dataTable, worksheet);
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetHidden_2_2_Name));
            }
        }

        static void worksheetHidden_2_3_Write(ExcelPackage excelPackage, SqlConnection connection, string dateStart, string dateEnd, string productionCalendarIDStart, string productionCalendarIDEnd)
        {
            //////////////////Здоровье динамических и водопадных проектов
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_3_Name);

            try
            {
                DataTable dataTable = new DataTable();
                using (var cmd = new SqlCommand())
                {
                    cmd.Connection = connection;
                    cmd.CommandText = "[ITProject].[spGetExcelReportProjectScore]";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                    cmd.Parameters.AddWithValue("@StartDateString", dateStart);
                    cmd.Parameters.AddWithValue("@EndDateString", dateEnd);
                    cmd.Parameters.AddWithValue("@ProductionCalendarIDStart", productionCalendarIDStart);
                    cmd.Parameters.AddWithValue("@ProductionCalendarIDEnd", productionCalendarIDEnd);
                    cmd.Parameters.AddWithValue("@IsFinishedOnly", false);
                    cmd.ExecuteNonQuery();
                    var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                    var dataSet = new DataSet();
                    dataAdapter.Fill(dataSet);
                    dataTable = dataSet.Tables[0];
                }

                Helpers.Excel.WriteDataTableToWorkSheet(dataTable, worksheet);

                //считаем число проектов, пригодится
                for (int i = 2; i <= worksheet.Cells.Rows; i++)
                    if (Convert.ToString(worksheet.Cells[i, 1].Value ?? "") != "")
                        Settings.Variables.WorksheetHidden_2_3_ProjectCount++;

            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetHidden_2_3_Name));
            }
        }

        static void worksheetHidden_2_4_Write(ExcelPackage excelPackage, SqlConnection connection, string dateStart, string dateEnd, string productionCalendarIDStart, string productionCalendarIDEnd)
        {
            //////////////////Здоровье динамических и водопадных проектов
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_4_Name);

            try
            {
                DataTable dataTable = new DataTable();
                using (var cmd = new SqlCommand())
                {
                    cmd.Connection = connection;
                    cmd.CommandText = "[ITProject].[spGetExcelReportProjectScore]";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                    cmd.Parameters.AddWithValue("@StartDateString", dateStart);
                    cmd.Parameters.AddWithValue("@EndDateString", dateEnd);
                    cmd.Parameters.AddWithValue("@ProductionCalendarIDStart", productionCalendarIDStart);
                    cmd.Parameters.AddWithValue("@ProductionCalendarIDEnd", productionCalendarIDEnd);
                    cmd.Parameters.AddWithValue("@IsFinishedOnly", true);
                    cmd.ExecuteNonQuery();
                    var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                    var dataSet = new DataSet();
                    dataAdapter.Fill(dataSet);
                    dataTable = dataSet.Tables[0];
                }

                Helpers.Excel.WriteDataTableToWorkSheet(dataTable, worksheet);

                //считаем число проектов, пригодится
                for (int i = 2; i <= worksheet.Cells.Rows; i++)
                    if (Convert.ToString(worksheet.Cells[i, 1].Value ?? "") != "")
                        Settings.Variables.WorksheetHidden_2_4_ProjectCount++;
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetHidden_2_4_Name));
            }
        }

        static void worksheetHidden_2_5_Write(ExcelPackage excelPackage, SqlConnection connection, string dateStart, string dateEnd, string stateIdList, string projectTypeIdList)
        {
            //////////////////Исключены из портфеля
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_5_Name);

            try
            {
                DataTable dataTable = new DataTable();
                using (var cmd = new SqlCommand())
                {
                    cmd.Connection = connection;
                    cmd.CommandText = "[ITProject].[spGetExcelReportBriefcaseDynamicPortfolio]";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                    cmd.Parameters.AddWithValue("@StartDateString", dateStart);
                    cmd.Parameters.AddWithValue("@EndDateString", dateEnd);
                    cmd.Parameters.AddWithValue("@ProjectStateIDList", stateIdList);
                    cmd.Parameters.AddWithValue("@ProjectTypeIDList", projectTypeIdList);
                    cmd.Parameters.AddWithValue("@Mode", "ChangedDate_finished");

                    cmd.ExecuteNonQuery();
                    var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                    var dataSet = new DataSet();
                    dataAdapter.Fill(dataSet);
                    dataTable = dataSet.Tables[0];
                }

                Helpers.Excel.WriteDataTableToWorkSheet(dataTable, worksheet);

                //считаем число проектов, пригодится
                for (int i = 2; i <= worksheet.Cells.Rows; i++)
                    if (Convert.ToString(worksheet.Cells[i, 1].Value ?? "") != "")
                        Settings.Variables.WorksheetHidden_2_5_ProjectCount++;
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetHidden_2_5_Name));
            }
        }

        static void worksheetHidden_2_6_Write(ExcelPackage excelPackage, SqlConnection connection, string dateStart, string dateEnd, string stateIdList, string projectTypeIdList)
        {
            //////////////////Включены в портфель
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_6_Name);

            try
            {
                DataTable dataTable = new DataTable();
                using (var cmd = new SqlCommand())
                {
                    cmd.Connection = connection;
                    cmd.CommandText = "[ITProject].[spGetExcelReportBriefcaseDynamicPortfolio]";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                    cmd.Parameters.AddWithValue("@StartDateString", dateStart);
                    cmd.Parameters.AddWithValue("@EndDateString", dateEnd);
                    cmd.Parameters.AddWithValue("@ProjectStateIDList", stateIdList);
                    cmd.Parameters.AddWithValue("@ProjectTypeIDList", projectTypeIdList);
                    cmd.Parameters.AddWithValue("@Mode", "ChangedDate_portfolio");

                    cmd.ExecuteNonQuery();
                    var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                    var dataSet = new DataSet();
                    dataAdapter.Fill(dataSet);
                    dataTable = dataSet.Tables[0];
                }

                Helpers.Excel.WriteDataTableToWorkSheet(dataTable, worksheet);

                //считаем число проектов, пригодится
                for (int i = 2; i <= worksheet.Cells.Rows; i++)
                    if (Convert.ToString(worksheet.Cells[i, 1].Value ?? "") != "")
                        Settings.Variables.WorksheetHidden_2_6_ProjectCount++;
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetHidden_2_6_Name));
            }
        }

        static void worksheetHidden_3_0_Write(ExcelPackage excelPackage, SqlConnection connection, string dateStart, string dateEnd, string stateIdList, string projectTypeIdList)
        {
            //////////////////Статистика запросов на изменение портфеля
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_3_0_Name);

            try
            {
                DataTable dataTable = new DataTable();
                using (var cmd = new SqlCommand())
                {
                    cmd.Connection = connection;
                    cmd.CommandText = "[ITProject].[spGetExcelReportPortfolioChange]";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                    cmd.Parameters.AddWithValue("@StartDateString", dateStart);
                    cmd.Parameters.AddWithValue("@EndDateString", dateEnd);
                    cmd.Parameters.AddWithValue("@ProjectStateIDList", stateIdList);
                    cmd.Parameters.AddWithValue("@ProjectTypeIDList", projectTypeIdList);
                    cmd.ExecuteNonQuery();
                    var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                    var dataSet = new DataSet();
                    dataAdapter.Fill(dataSet);
                    dataTable = dataSet.Tables[0];
                }

                Helpers.Excel.WriteDataTableToWorkSheet(dataTable, worksheet);
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetHidden_3_0_Name));
            }
        }

        static void worksheetHidden_3_1_Write(ExcelPackage excelPackage, SqlConnection connection, string dateStart, string dateEnd, string stateIdList, string projectTypeIdList)
        {
            //////////////////Статистика запросов на изменение портфеля
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_3_1_Name);

            try
            {
                DataTable dataTable = new DataTable();
                using (var cmd = new SqlCommand())
                {
                    cmd.Connection = connection;
                    cmd.CommandText = "[ITProject].[spGetExcelReportPortfolioChangeStatistic]";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                    cmd.Parameters.AddWithValue("@StartDateString", dateStart);
                    cmd.Parameters.AddWithValue("@EndDateString", dateEnd);
                    cmd.Parameters.AddWithValue("@ProjectStateIDList", stateIdList);
                    cmd.Parameters.AddWithValue("@ProjectTypeIDList", projectTypeIdList);
                    cmd.ExecuteNonQuery();
                    var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                    var dataSet = new DataSet();
                    dataAdapter.Fill(dataSet);
                    dataTable = dataSet.Tables[0];
                }

                foreach (DataRow row in dataTable.Rows)
                    row[0] = Convert.ToString(row[0]).Replace(Environment.NewLine, null);

                int projectCountRowIndex = Helpers.SugarDataTable.GetRowIndex(dataTable, 0, "Temp_ProjectCount");
                int waitingCountRowIndex = Helpers.SugarDataTable.GetRowIndex(dataTable, 0, "Temp_WaitingCount");

                //записываем число проектов на изменения
                Settings.Variables.WorksheetHidden_3_1_ProjectCount = Convert.ToInt16(Convert.ToString(dataTable.Rows[projectCountRowIndex][1]));
                Settings.Variables.WorksheetHidden_3_1_WaitCount = Convert.ToInt16(Convert.ToString(dataTable.Rows[waitingCountRowIndex][1]));

                foreach (DataRow row in dataTable.Rows)
                    if (Convert.ToString(row[0]).IndexOf("Temp_") > -1)
                        row.Delete();
                dataTable.AcceptChanges();

                Helpers.Excel.WriteDataTableToWorkSheet(dataTable, worksheet);

                //считаем число параметров
                /*
                for (int i = 2; i <= worksheet.Cells.Rows; i++)
                    if (Convert.ToString(worksheet.Cells[i, 1].Value ?? "") != "")
                        Settings.Variables.WorksheetHidden_3_1_CountLine++;
                */

                Settings.Variables.WorksheetHidden_3_1_CountLine = dataTable.Rows.Count;
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetHidden_3_1_Name));
            }
        }

        static void worksheetHidden_4_0_Write(ExcelPackage excelPackage, SqlConnection connection, string dateStart, string dateEnd, int startPeriodNumber, int endPeriodNumber, string stateIdList, string projectTypeIdList)
        {
            //////////////////Статистика запросов на изменение портфеля
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_4_0_Name);

            try
            {
                DataTable dataTable = new DataTable();
                using (var cmd = new SqlCommand())
                {
                    cmd.Connection = connection;
                    cmd.CommandText = "[ITProject].[spGetExcelReportProjectIntegralRatings]";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                    cmd.Parameters.AddWithValue("@StartDateString", dateStart);
                    cmd.Parameters.AddWithValue("@EndDateString", dateEnd);
                    cmd.Parameters.AddWithValue("@StartPeriodNumber", startPeriodNumber);
                    cmd.Parameters.AddWithValue("@EndPeriodNumber", endPeriodNumber);
                    cmd.Parameters.AddWithValue("@ProjectStateIDList", stateIdList);
                    cmd.Parameters.AddWithValue("@ProjectTypeIDList", projectTypeIdList);
                    cmd.ExecuteNonQuery();
                    var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                    var dataSet = new DataSet();
                    dataAdapter.Fill(dataSet);
                    dataTable = dataSet.Tables[0];
                }

                Helpers.Excel.WriteDataTableToWorkSheet(dataTable, worksheet);
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetHidden_4_0_Name));
            }
        }

        static void worksheetHidden_Debug_Write(ExcelPackage excelPackage, SqlConnection connection, string dateStart, string dateEnd, string stateIdList, string projectTypeIdList)
        {
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_Debug_Name);

            try
            {
                Settings.Variables.DefaultLabelStyle.SetPropertiesFromCell(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_StyleCellLabel);
                Settings.Variables.DefaultDataTableHeaderStyle.SetPropertiesFromCell(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_StyleCellTableHeader);
                Settings.Variables.DefaultDataTableColumnHeaderStyle.SetPropertiesFromCell(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_StyleCellColumnHeader);
                Settings.Variables.DefaultDataTableCellStyle.SetPropertiesFromCell(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_StyleCellInner);

                string rgb_splitter = ", ";
                Settings.Variables.DefaultLabelStyle.SetCellBorderColorFromCellValuesRGB(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_CellForBorderLabel, rgb_splitter);
                Settings.Variables.DefaultDataTableHeaderStyle.SetCellBorderColorFromCellValuesRGB(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_CellForBorderTableHeader, rgb_splitter);
                Settings.Variables.DefaultDataTableColumnHeaderStyle.SetCellBorderColorFromCellValuesRGB(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_CellForBorderColumnHeader, rgb_splitter);
                Settings.Variables.DefaultDataTableCellStyle.SetCellBorderColorFromCellValuesRGB(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_CellForBorderInner, rgb_splitter);

                int debugRow = int.Parse(Settings.SQLVariables.WorksheetHidden_Debug_Row);

                worksheet.Cells[debugRow, 1].Value = "dateStart";
                worksheet.Cells[debugRow, 2].Value = dateStart;
                debugRow++;

                worksheet.Cells[debugRow, 1].Value = "dateEnd";
                worksheet.Cells[debugRow, 2].Value = dateEnd;
                debugRow++;

                worksheet.Cells[debugRow, 1].Value = "stateIdList";
                worksheet.Cells[debugRow, 2].Value = stateIdList;
                debugRow++;

                worksheet.Cells[debugRow, 1].Value = "projectTypeIdList";
                worksheet.Cells[debugRow, 2].Value = projectTypeIdList;
                debugRow++;
                debugRow++;

                Helpers.Excel.WriteDataTableToWorkSheet(debugRow, 1, true, Settings.Variables.ProductionCalendar, worksheet);
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetHidden_Debug_Name));
            }
        }

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
