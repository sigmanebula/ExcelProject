namespace Helpers
{
    public static partial class Sugar
    {
        public static string RemoveStringLastChars(string str, string removeChars, ref string errorText)
        {
            if (errorText == "")
                try
                {
                    if (removeChars.Length > 0)
                        for (int i = 0; i < removeChars.Length && str.Length > 0; i++)
                            if (str[str.Length - 1] == removeChars[i])
                            {
                                str = str.Substring(0, str.Length - 1);
                                i = -1;
                            }
                }
                catch(System.Exception exception)
                {
                    errorText += "\nОшибка форматирования строки: " + str + ", причина: " + exception.Message;
                }

            return str;
        }
        
        public static string RemoveStringLastChars(string str, string removeChars)
        {
            string errorText = "";

            str = RemoveStringLastChars(str, removeChars, ref errorText);

            if (errorText != "")
                throw new System.Exception(errorText);

            return str;
        }

    }
}
