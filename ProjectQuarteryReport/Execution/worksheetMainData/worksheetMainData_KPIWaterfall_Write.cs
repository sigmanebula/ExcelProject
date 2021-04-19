using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;

namespace ProjectQuarteryReport
{
    public static partial class Execution
    {
        static void worksheetMainData_KPIWaterfall_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string productionCalendarID
            , string methodType
            , string startColumn
            , string columnList
            )
        {
            DataTable dataTable = new DataTable();
            using (var cmd = new SqlCommand())
            {
                cmd.Connection = connection;
                cmd.CommandText = "[ITProject].[spGetExcelProjectQuarteryReportKPIWaterfallPart]";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                cmd.Parameters.AddWithValue("@ProjectID", Settings.Variables.ProjectID);
                cmd.Parameters.AddWithValue("@ProductionCalendarID", productionCalendarID);
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

        static void worksheetMainData_KPIWaterfall_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string productionCalendarID
            , string methodType
            , string startColumn
            )
        {
            worksheetMainData_KPIWaterfall_Write(
                  worksheet
                , connection
                , productionCalendarID
                , methodType
                , startColumn
                , ""
                );
        }
    }
}
