namespace Helpers
{
    public static partial class SugarFile
    {
        public static void DeleteIfExists(string fileFullName, string errorPrefix, ref string errorText)
        {
            //!не делаем проверку на errorText == ""
            if (fileFullName != "" && System.IO.File.Exists(fileFullName))
            {
                try
                {
                    System.IO.File.Delete(fileFullName);
                }
                catch (System.Exception ex)
                {
                    errorText += errorPrefix + "\nНе удалось удалить " + fileFullName + ", причина: " + ex.Message;
                }
            }
        }
        
        public static void DeleteIfExists(string fileFullName, string errorPrefix)
        {
            string errorText = "";
            DeleteIfExists(fileFullName, errorPrefix, ref errorText);

            if (errorText != "")
                throw new System.Exception(errorText);
        }
    }
}
