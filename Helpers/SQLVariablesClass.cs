namespace Helpers
{
    public class SQLVariablesClass
    {
        public string FolderPath { get; set; }

        public void SetValueByName(string propertyName, string value)
        {

                foreach (var property in this.GetType().GetProperties())
                    if (property.Name == propertyName)
                    {

                            property.SetValue(this, value);
 
                    }
        }

        public string GetStringListSettings(string prefix)
        {
            string result = prefix + "Settings:";

            foreach (var property in this.GetType().GetProperties())
                result += prefix + property.Name + " " + (property.GetValue(this) ?? "").ToString();

            return result;
        }

        public string GetStringListSettings()
        {
            return GetStringListSettings("\n");
        }


        public void GetSettings(System.Data.SqlClient.SqlConnection connection, string settingsTypeCodeList)
        {
                try
                {
                    System.Data.DataTable dataTable = Helpers.SugarSQLConnection.ExecuteSQLCommand(
                          connection
                        , System.String.Format(SettingsGlobal.SQLCommandGetSettings, settingsTypeCodeList)
                        , "Не удалось получить настройки, причина: "
                        );

                    foreach (System.Data.DataRow row in dataTable.Rows)
                        SetValueByName(row["Code"].ToString(), row["Value"].ToString());

                }
                catch (System.Exception exception)
                {
                    exception.Message = "\nОшибка: " + exception.Message;
                }
        }
    }
}
