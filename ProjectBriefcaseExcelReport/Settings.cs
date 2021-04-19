using System;
using System.Data;

namespace ProjectBriefcaseExcelReport
{
    public static class Settings
    {
        public static SQLVariablesClass SQLVariables = new SQLVariablesClass();
        public static VariablesClass Variables = new VariablesClass();

        public static string FileExtention                   = "xlsx";
        public static string ProductionCalendarCodeStart     = "start";
        public static string ProductionCalendarCodeEnd       = "end";
        public static string ProductionCalendarCodeStartNext = "start_next";
        public static string ProductionCalendarCodeEndNext   = "end_next";
        public static string SettingsTypeCodeList            = "'ProjectBriefcaseExcelReport'";


        public class SQLVariablesClass : Helpers.SQLVariablesClass
        {
            public string TemplateFileShortName { get; set; }
            public string NewFileNamePrefix { get; set; }
            public string IsGetErrorMessage { get; set; }

            public string WorksheetHidden_Debug_Name { get; set; }

            public string WorksheetHidden_Debug_StyleCellLabel { get; set; }
            public string WorksheetHidden_Debug_StyleCellTableHeader { get; set; }
            public string WorksheetHidden_Debug_StyleCellColumnHeader { get; set; }
            public string WorksheetHidden_Debug_StyleCellInner { get; set; }

            public string WorksheetHidden_Debug_CellForBorderLabel { get; set; }
            public string WorksheetHidden_Debug_CellForBorderTableHeader { get; set; }
            public string WorksheetHidden_Debug_CellForBorderColumnHeader { get; set; }
            public string WorksheetHidden_Debug_CellForBorderInner { get; set; }

            public string WorksheetHidden_Debug_Row { get; set; }

            public string WorksheetShown_1_Name { get; set; }
            public string WorksheetShown_1_HeaderEndDateLocation { get; set; }
            public string WorksheetShown_1_PeriodLocation { get; set; }
            public string WorksheetShown_1_SmallLabelEndDateLocation { get; set; }
            public string WorksheetShown_2_Name { get; set; }
            public string WorksheetHidden_2_1_Name { get; set; }
            public string WorksheetHidden_2_2_Name { get; set; }
            public string WorksheetHidden_2_3_Name { get; set; }
            public string WorksheetHidden_2_4_Name { get; set; }
            public string WorksheetHidden_2_5_Name { get; set; }
            public string WorksheetHidden_2_6_Name { get; set; }

            public string WorksheetShown_2_Header_LabelStartCell { get; set; }
            public string WorksheetShown_2_3_Header_LabelStartCell { get; set; }
            public string WorksheetShown_2_2_Header_LabelStartCell { get; set; }

            public string WorksheetShown_2_3_4_LabelStartCell { get; set; }
            public string WorksheetShown_2_3_4_MergeCellCount { get; set; }

            public string WorksheetShown_2_5_LabelStartCell { get; set; }
            public string WorksheetShown_3_Name { get; set; }
            public string WorksheetHidden_3_0_Name { get; set; }
            public string WorksheetHidden_3_1_Name { get; set; }
            public string WorksheetShown_3_Header_LabelStartCell { get; set; }
            public string WorksheetShown_3_LabelStartCell { get; set; }
            public string WorksheetShown_3_1_TableStartCell { get; set; }

            public string WorksheetHidden_4_0_Name { get; set; }
            public string WorksheetShown_4_Name { get; set; }
        }

        public class VariablesClass : Helpers.VariablesClass
        {
            public DataTable ProductionCalendar { get; set; }

            public Helpers.Excel.CellStyleClass DefaultLabelStyle = new Helpers.Excel.CellStyleClass();

            public Helpers.Excel.CellStyleClass DefaultDataTableHeaderStyle = new Helpers.Excel.CellStyleClass();

            public Helpers.Excel.CellStyleClass DefaultDataTableColumnHeaderStyle = new Helpers.Excel.CellStyleClass();

            public Helpers.Excel.CellStyleClass DefaultDataTableCellStyle = new Helpers.Excel.CellStyleClass();

            public int WorksheetHidden_2_3_ProjectCount { get; set; }
            public int WorksheetHidden_2_4_ProjectCount { get; set; }
            public int WorksheetHidden_2_5_ProjectCount { get; set; }
            public int WorksheetHidden_2_6_ProjectCount { get; set; }

            public string WorksheetShown_2_1_StartCell { get; set; }
            public string WorksheetShown_2_3_4_StartCell { get; set; }
            public string WorksheetShown_2_5_StartCell { get; set; }
            public string WorksheetShown_2_6_StartCell { get; set; }
            public string WorksheetShown_2_7_StartCell { get; set; }

            public string WorksheetShown_2_1_EndCell { get; set; }
            public string WorksheetShown_2_3_4_EndCell { get; set; }
            public string WorksheetShown_2_5_EndCell { get; set; }
            public string WorksheetShown_2_6_EndCell { get; set; }
            public string WorksheetShown_2_7_EndCell { get; set; }


            public string WorksheetHidden_2_1_DataStartCell { get; set; }
            public string WorksheetHidden_2_1_DataEndCell { get; set; }
            public int WorksheetHidden_3_1_ProjectCount { get; set; }
            public int WorksheetHidden_3_1_WaitCount { get; set; }
            public int WorksheetHidden_3_1_CountLine { get; set; }

            public string WorksheetShown_3_1_StartCell { get; set; }
            public string WorksheetShown_3_1_EndCell { get; set; }

            public string GetProductionCalendar(string code, string columnName)
            {
                string result = "";
                for (int i = 0; i < ProductionCalendar.Rows.Count; i++)
                    if (Convert.ToString(ProductionCalendar.Rows[i]["Code"] ?? "") == code)
                        return Convert.ToString(ProductionCalendar.Rows[i][columnName] ?? "");
                return result;
            }

            /*
            public void Initialize_DefaultDataTableCellStyle(bool isNew)
            {
                if (isNew)
                    DefaultDataTableCellStyle = new Helpers.Excel.CellStyleClass();

                System.Drawing.Color color                      = System.Drawing.Color.Gray;
                
                //DefaultDataTableCellStyle.HorizontalAlignment   = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                //DefaultDataTableCellStyle.VerticalAlignment     = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                //DefaultDataTableCellStyle.FontColor             = System.Drawing.Color.DarkBlue;
                //DefaultDataTableCellStyle.FontBold              = false;
                //DefaultDataTableCellStyle.FillBackgroundColor   = System.Drawing.Color.Transparent;

                DefaultDataTableCellStyle.BorderLeftStyle       = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                DefaultDataTableCellStyle.BorderRightStyle      = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                DefaultDataTableCellStyle.BorderTopStyle        = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                DefaultDataTableCellStyle.BorderBottomStyle     = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                DefaultDataTableCellStyle.BorderTopColor        = color;
                DefaultDataTableCellStyle.BorderBottomColor     = color;
                DefaultDataTableCellStyle.BorderLeftColor       = color;
                DefaultDataTableCellStyle.BorderRightColor      = color;
            }
            */
        }





    }
}
