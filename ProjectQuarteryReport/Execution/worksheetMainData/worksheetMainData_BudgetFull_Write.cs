using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;

namespace ProjectQuarteryReport
{
    public static partial class Execution
    {
        static void worksheetMainData_BudgetFull_Write(ExcelWorksheet worksheet, SqlConnection connection)
        {
            DataTable dataTable = new DataTable();
            using (var cmd = new SqlCommand())
            {
                cmd.Connection = connection;
                cmd.CommandText = "[ITProject].[spGetProjectBudgetFullQuartersAll]";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                cmd.Parameters.AddWithValue("@ProjectID", Settings.Variables.ProjectID);
                cmd.ExecuteNonQuery();
                var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                var dataSet = new DataSet();
                dataAdapter.Fill(dataSet);
                dataTable = dataSet.Tables[0];
            }

            DataTable dataTableSorted = Helpers.SugarDataTable.CopyDataTableByColumnList(dataTable, Settings.SQLVariables.Data_BudgetFull_FieldList);
                
            Helpers.Excel.WriteDataTableToWorkSheet(
                  Settings.SQLVariables.Data_BudgetFull_Column + Settings.SQLVariables.Data_BudgetFull_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.Data_BudgetFull_IsDataTableHasHeaders)
                , dataTableSorted
                , worksheet
                );
        }
    }
}
