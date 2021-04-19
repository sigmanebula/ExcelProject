using System;
using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;

namespace ProjectQuarteryReport
{
    public static partial class Execution
    {
        static void worksheetMainData_BudgetGeneral_Write(ExcelWorksheet worksheet, SqlConnection connection, string budgetName, string startColumn)
        {
            DataTable dataTable = Helpers.SugarSQLConnection.ExecuteSQLCommand(
                  connection
                , String.Format(Settings.SQLCommandGetProjectBudgetGeneral, Settings.Variables.ProjectID, budgetName)
                , ""
                );
                
            DataTable dataTableSorted = Helpers.SugarDataTable.CopyDataTableByColumnList(dataTable, Settings.SQLVariables.Data_BudgetGeneral_FieldList);

            Helpers.Excel.WriteDataTableToWorkSheet(
                  startColumn + Settings.SQLVariables.Data_BudgetGeneral_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsDataTableHasHeaders)
                , dataTableSorted
                , worksheet
                );
        }
    }
}