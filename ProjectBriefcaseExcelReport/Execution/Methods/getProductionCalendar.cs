using System;
using System.Data;
using System.Data.SqlClient;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void getProductionCalendar(string dateStart, string dateEnd, SqlConnection connection, ref string errorText)
        {
            if (errorText == "")
                try
                {
                    using (var cmd = new SqlCommand()) //записываем
                    {
                        cmd.Connection = connection;
                        cmd.CommandText = "[ITProject].[spGetExcelReportProductionCalendar]";
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                        cmd.Parameters.AddWithValue("@StartDate", dateStart);
                        cmd.Parameters.AddWithValue("@EndDate", dateEnd);
                        cmd.ExecuteNonQuery();
                        var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                        var dataSet = new DataSet();
                        dataAdapter.Fill(dataSet);
                        Settings.Variables.ProductionCalendar = dataSet.Tables[0];

                        foreach (DataRow row in Settings.Variables.ProductionCalendar.Rows)
                            row["Date"] = Convert.ToString(row["Date"] ?? "").Replace("Z", "");
                    }
                }
                catch (Exception ex)
                {
                    errorText += "Не удалось получить данные календаря, причина: " + ex.Message;
                }
        }
    }
}
