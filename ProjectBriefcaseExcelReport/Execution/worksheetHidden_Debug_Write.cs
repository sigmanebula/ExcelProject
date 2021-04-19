using System;
using System.Data.SqlClient;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void worksheetHidden_Debug_Write(ExcelPackage excelPackage, SqlConnection connection, string dateStart, string dateEnd, string stateIdList, string projectTypeIdList)
        {
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_Debug_Name);
            
            try
            {
                Settings.Variables.DefaultLabelStyle                 .SetPropertiesFromCell(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_StyleCellLabel);
                Settings.Variables.DefaultDataTableHeaderStyle       .SetPropertiesFromCell(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_StyleCellTableHeader);
                Settings.Variables.DefaultDataTableColumnHeaderStyle .SetPropertiesFromCell(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_StyleCellColumnHeader);
                Settings.Variables.DefaultDataTableCellStyle         .SetPropertiesFromCell(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_StyleCellInner);

                string rgb_splitter = ", ";
                Settings.Variables.DefaultLabelStyle                 .SetCellBorderColorFromCellValuesRGB(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_CellForBorderLabel, rgb_splitter);
                Settings.Variables.DefaultDataTableHeaderStyle       .SetCellBorderColorFromCellValuesRGB(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_CellForBorderTableHeader, rgb_splitter);
                Settings.Variables.DefaultDataTableColumnHeaderStyle .SetCellBorderColorFromCellValuesRGB(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_CellForBorderColumnHeader, rgb_splitter);
                Settings.Variables.DefaultDataTableCellStyle         .SetCellBorderColorFromCellValuesRGB(worksheet, Settings.SQLVariables.WorksheetHidden_Debug_CellForBorderInner, rgb_splitter);
                
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
    }
}