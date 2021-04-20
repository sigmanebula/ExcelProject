using System;

namespace ProjectBriefcaseExcelReport
{
    public class VariablesClass : Helpers.VariablesClass
        {
            public DataTable ProductionCalendar { get; set; }

            public Helpers.CellStyleClass DefaultLabelStyle = new Helpers.CellStyleClass();

            public Helpers.CellStyleClass DefaultDataTableHeaderStyle = new Helpers.CellStyleClass();

            public Helpers.CellStyleClass DefaultDataTableColumnHeaderStyle = new Helpers.CellStyleClass();

            public Helpers.CellStyleClass DefaultDataTableCellStyle = new Helpers.CellStyleClass();

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
                    DefaultDataTableCellStyle = new Helpers.CellStyleClass();

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

