using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;

namespace ProjectQuarteryReport
{
    public static partial class Execution
    {
        static void worksheetMainData_KPIDynamic_Comment_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string year
            , string quarter
            , string methodType
            , string startColumn
            , string columnList
            )
        {
            DataTable dataTable = new DataTable();
            using (var cmd = new SqlCommand())
            {
                cmd.Connection = connection;
                cmd.CommandText = "[ITProject].[spGetKPIDynamicComment]";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                cmd.Parameters.AddWithValue("@ProjectID", Settings.Variables.ProjectID);
                cmd.Parameters.AddWithValue("@Year", year);
                cmd.Parameters.AddWithValue("@Quarter", quarter);
                cmd.Parameters.AddWithValue("@MethodType", methodType);
                cmd.ExecuteNonQuery();
                var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                var dataSet = new DataSet();
                dataAdapter.Fill(dataSet);
                dataTable = dataSet.Tables[0];
            }

            DataTable dataTableSorted = Helpers.SugarDataTable.CopyDataTableByColumnList(dataTable, columnList);

            Helpers.Excel.WriteDataTableToWorkSheet(
                  startColumn + Settings.SQLVariables.KPIData_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsDataTableHasHeaders)
                , dataTableSorted
                , worksheet
                );
        }

        static void worksheetMainData_KPIDynamic_Comment_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string year
            , string quarter
            , string methodType
            , string startColumn
            )
        {
            worksheetMainData_KPIDynamic_Comment_Write(
                  worksheet
                , connection
                , year
                , quarter
                , methodType
                , startColumn
                , ""
                );
        }

    }
}