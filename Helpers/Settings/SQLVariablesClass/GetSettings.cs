namespace Helpers
{
    public partial class SQLVariablesClass
    {
        public void GetSettings(System.Data.SqlClient.SqlConnection connection, string settingsTypeCodeList, ref string errorText)
        {
            if (errorText == "")
                try
                {
                    System.Data.DataTable dataTable = Helpers.SugarSQLConnection.ExecuteSQLCommand(
                          connection
                        , System.String.Format(SettingsGlobal.SQLCommandGetSettings, settingsTypeCodeList)
                        , "Не удалось получить настройки, причина: "
                        );

                    foreach (System.Data.DataRow row in dataTable.Rows)
                        SetValueByName(row["Code"].ToString(), row["Value"].ToString(), ref errorText);

                    CheckValues(ref errorText);
                }
                catch (System.Exception exception)
                {
                    errorText += "\nОшибка: " + exception.Message;
                }
        }
    }
}