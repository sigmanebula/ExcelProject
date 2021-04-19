using System;
using System.Data;
using System.Data.SqlClient;

namespace ProjectQuarteryReport
{
    public static partial class Execution
    {
        static void getProductionCalendarData(SqlConnection connection, ref string errorText)
        {
            if (errorText == "")
                try
                {
                    using (var cmd = new SqlCommand())
                    {
                        cmd.Connection = connection;
                        cmd.CommandText = "[ITProject].[spGetExcelProductionCalendar]";
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                        cmd.Parameters.AddWithValue("@ProductionCalendarID", Settings.Variables.ProductionCalendarID);
                        cmd.ExecuteNonQuery();
                        var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                        var dataSet = new DataSet();
                        dataAdapter.Fill(dataSet);
                        Settings.Variables.ProductionCalendar = dataSet.Tables[0];
                    }
                }
                catch (Exception ex)
                {
                    errorText += "Ошибка: не удалось получить данные по календарю, причина: " + ex.Message;
                }
        }
    }
}