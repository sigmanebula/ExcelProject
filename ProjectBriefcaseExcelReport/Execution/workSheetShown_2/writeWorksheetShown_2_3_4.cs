using System;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
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
    }
}
