namespace Helpers
{
    public static partial class Sugar
    {
        public static string NormalizeFolderPath(string path, ref string errorText)
        {
            if (errorText == "")
            {
                try
                {
                    if (path[path.Length - 1] != '\\')
                        path += "\\";

                    if (!System.IO.Directory.Exists(path))
                        throw new System.Exception("\nДиректория не существует: " + path);
                }
                catch (System.Exception exception)
                {
                    errorText += exception.Message;
                }
            }
            return path;
        }
        
        public static string NormalizeFolderPath(string path)
        {
            string errorText = "";
            path = NormalizeFolderPath(path, ref errorText);

            if (errorText != "")
                throw new System.Exception(errorText);

            return path;
        }
    }
}
