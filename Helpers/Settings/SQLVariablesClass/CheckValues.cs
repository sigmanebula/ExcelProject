namespace Helpers
{
    public partial class SQLVariablesClass
    {
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
