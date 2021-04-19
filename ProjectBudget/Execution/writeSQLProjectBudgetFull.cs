using System;
using System.Data;
using System.Data.SqlClient;

namespace ProjectBudget
{
    public static partial class Execution
    {
        static void writeSQLProjectBudgetFull(SqlConnection connection, int projectNumber, string xml, ref string errorText)
        {
            if (errorText == "")
                try
                {
                    if (xml == "")
                        throw new System.Exception("xml пуст!");

                    using (var command = new SqlCommand()) //записываем
                    {
                        command.Connection = connection;
                        command.CommandText = "[ITProject].[spCreateProjectBudgetFull]";
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                        command.Parameters.AddWithValue("@XML", xml);
                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception exception)
                {
                    errorText += "\nОшибка записи полной информации по бюджету проекта " + projectNumber.ToString() + ": " + exception.Message;
                }
        }
    }
}