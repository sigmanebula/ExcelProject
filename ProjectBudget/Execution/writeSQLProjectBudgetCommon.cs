using System;
using System.Data;
using System.Data.SqlClient;

namespace ProjectBudget
{
    public static partial class Execution
    {
        static void writeSQLProjectBudgetCommon(SqlConnection connection, int projectID, int projectNumber, string xml, ref string errorText)
        {
            if (errorText == "")
                try
                {
                    if (xml == "")
                        throw new System.Exception("xml пуст!");

                    using (var command = new SqlCommand())
                    {
                        command.Connection = connection;
                        command.CommandText = "[ITProject].[spCreateProjectBudget]";
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                        command.Parameters.AddWithValue("@XML", xml);
                        command.Parameters.AddWithValue("@ProjectID", projectID);
                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception exception)
                {
                    errorText += "\nОшибка записи общей информации по бюджету проекта " + projectNumber.ToString() + ": " + exception.Message;
                }
        }
    }
}
