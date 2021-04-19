using System;
using System.Data;
using System.Data.SqlClient;

namespace ProjectQuarteryReport
{
    public static partial class Execution
    {
        static string getFileShortNameNew(SqlConnection connection, ref string errorText)
        {
            if (errorText == "")
            {
                try
                {
                    DataTable dataTable = Helpers.SugarSQLConnection.ExecuteSQLCommand(
                        connection
                        , String.Format(Settings.SQLCommandGetProjectNumberShortName, Settings.Variables.ProjectID)
                        , ""
                        );
                    
                    Settings.Variables.ProjectNumberShortName = dataTable.Rows[0][0].ToString();

                    return Settings.SQLVariables.NewFileNamePrefix
                        + Settings.Variables.ProjectNumberShortName
                        + " за "
                        + Settings.Variables.GetProductionCalendar("Current", "Year")
                        + " г. "
                        + Settings.Variables.GetProductionCalendar("Current", "Quarter")
                        + "кв.."
                        + Settings.FileExtention;
                }
                catch (Exception ex)
                {
                    errorText += "Ошибка: не удалось получить данные по проекту и календарю, причина: " + ex.Message;
                    return "";
                }
            }
            else
                return "";
        }
    }
}
