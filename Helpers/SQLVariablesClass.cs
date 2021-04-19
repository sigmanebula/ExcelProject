namespace Helpers
{
    public class SQLVariablesClass
    {
        public string FolderPath { get; set; }

        public void SetValueByName(string propertyName, string value, ref string errorText)
        {
            if (errorText == "")
                foreach (var property in this.GetType().GetProperties())
                    if (property.Name == propertyName)
                    {
                        try
                        {
                            property.SetValue(this, value);
                            break;
                        }
                        catch
                        {
                            errorText += "\nОшибка получения настроек: " + propertyName;
                        }
                    }
        }

        public void SetValueByName(string propertyName, string value)
        {
            string errorText = "";

            SetValueByName(propertyName, value, ref errorText);

            if (errorText != "")
                throw new System.Exception(errorText);
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


        public void CheckValues(ref string errorText, System.Collections.Generic.List<string> excludeValueNameList)
        {
            foreach (var property in this.GetType().GetProperties())
                if (property.GetValue(this) == null)
                {
                    bool founded = false;
                    foreach (string excludeValueName in excludeValueNameList)
                        if (excludeValueName == property.Name)
                        {
                            founded = true;
                            break;
                        }

                    if (!founded)
                    {
                        errorText += "\nПараметр настроек пуст: " + property.Name;
                        break;
                    }
                }

            if (errorText == "")
                FolderPath = Helpers.Sugar.NormalizeFolderPath(FolderPath, ref errorText);
        }

        public void CheckValues(System.Collections.Generic.List<string> excludeValueNameList)
        {
            string errorText = "";

            CheckValues(ref errorText, excludeValueNameList);

            if (errorText != "")
                throw new System.Exception(errorText);
        }

        public void CheckValues()
        {
            CheckValues(new System.Collections.Generic.List<string>());
        }

        public void CheckValues(ref string errorText)
        {
            CheckValues(ref errorText, new System.Collections.Generic.List<string>());
        }


    }
}
