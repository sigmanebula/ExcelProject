namespace Helpers
{
    public static partial class SugarFile
    {
        public static string GetK2Xml(string fileFullName, string fileShortName, string errorPrefix, ref string errorText)
        {
            if (errorText == "")
            {
                try
                {
                    return
                          "<file><name>"
                        + fileShortName
                        + "</name><content>"
                        + System.Convert.ToBase64String(System.IO.File.ReadAllBytes(fileFullName))
                        + "</content></file>";
                }
                catch (System.Exception ex)
                {
                    errorText += errorPrefix + "\nНе удалось создать K2 XML на основе файла: " + fileFullName + ", причина: " + ex.Message;
                    return "";
                }
            }
            else
                return "";
        }
    }
}
