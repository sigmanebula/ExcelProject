namespace Helpers
{
    public static partial class SugarFile
    {
        public static void Copy(string fileFullNameOld, string fileFullNameNew, ref string errorText)
        {
            //процедура нужна для перехвата и расшифровки ошибок, получение подробного текста ошибок актуально.
            if (errorText == "")
            {
                try
                {
                    SugarFile.DeleteIfExists(fileFullNameNew, "\nОшибка при удалении файла: ");

                    if (!System.IO.File.Exists(fileFullNameOld))
                        throw new System.Exception("\nфайл не найден: " + fileFullNameOld);

                    System.IO.File.Copy(fileFullNameOld, fileFullNameNew);
                }
                catch (System.Exception ex)
                {
                    errorText +=
                        "\nОшибка при копировании файла: "
                        + ex.Message
                        + "\nfileFullNameOld: "
                        + fileFullNameOld
                        + "\nfileFullNameNew"
                        + fileFullNameNew
                        ;
                }
            }
        }
    }
}
