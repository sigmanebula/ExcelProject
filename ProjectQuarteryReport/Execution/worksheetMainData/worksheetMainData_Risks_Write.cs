using System;
using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;

namespace ProjectQuarteryReport
{
    public static partial class Execution
    {
        static void worksheetMainData_Risks_Write(ExcelWorksheet worksheet, SqlConnection connection)
        {
            DataTable dataTable = Helpers.SugarSQLConnection.ExecuteSQLCommand(
                  connection
                , String.Format(
                      Settings.SQLCommandGetProjectQuarterlyRisks
                    , Settings.Variables.ProjectID
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    )
                , ""
                );

            Helpers.Excel.WriteDataTableToWorkSheet(
                  Settings.SQLVariables.Data_Risks_Column + Settings.SQLVariables.Data_Risks_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsDataTableHasHeaders)
                , dataTable
                , worksheet
                );
        }
    }
}

