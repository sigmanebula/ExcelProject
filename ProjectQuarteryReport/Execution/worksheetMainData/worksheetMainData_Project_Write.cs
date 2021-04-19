using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;

namespace ProjectQuarteryReport
{
    public static partial class Execution
    {
        static void worksheetMainData_Project_Write(ExcelWorksheet worksheet, SqlConnection connection)
        {
            DataTable dataTable = new DataTable();
            using (var cmd = new SqlCommand())
            {
                cmd.Connection = connection;
                cmd.CommandText = "[ITProject].[spGetExcelProjectQuarteryReportMainPart]";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                cmd.Parameters.AddWithValue("@ProjectID", Settings.Variables.ProjectID);
                cmd.Parameters.AddWithValue("@ProductionCalendarID", Settings.Variables.ProductionCalendarID);
                cmd.Parameters.AddWithValue("@StuffPrefix", Settings.SQLVariables.StuffPrefix);
                cmd.ExecuteNonQuery();
                var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                var dataSet = new DataSet();
                dataAdapter.Fill(dataSet);
                dataTable = dataSet.Tables[0];
            }

            Helpers.Excel.WriteDataTableToWorkSheet(
                  Settings.SQLVariables.Data_Project_Column + Settings.SQLVariables.MainData_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsDataTableHasHeaders)
                , dataTable
                , worksheet
                );
        }
    }
}
