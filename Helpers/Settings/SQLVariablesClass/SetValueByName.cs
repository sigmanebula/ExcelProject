namespace Helpers
{
    public partial class SQLVariablesClass
    {
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
    }
}
