namespace Helpers
{
    public static partial class SugarFile
    {
        public static string CreateFromK2Xml(string fileK2XML, string folderPath, string fileExtention, ref string errorText)
        {
            string fileShortName = "";
            string fileFullname = "";
            string fileContent = "";

            fileExtention = fileExtention.Replace(".", "");
            folderPath = Sugar.NormalizeFolderPath(folderPath);

            var dataSet = Sugar.GetDataSetFromXML(fileK2XML, ref errorText);

            if (errorText == "")
            {
                fileShortName = (dataSet.Tables[0].Rows[0]["name"] ?? "").ToString();
                fileContent = (dataSet.Tables[0].Rows[0]["content"] ?? "").ToString();

                if (GetExtention(fileShortName) != fileExtention)
                    errorText += "Создание файла из К2 XML. Расширение файла должно быть ." + fileExtention + System.Environment.NewLine;
            }

            if (errorText == "")
                fileFullname = folderPath + fileShortName;

            if (errorText == "")
                try
                {
                    DeleteIfExists(fileFullname, "Создание файла из К2 XML. Удаление имеющегося файла. ", ref errorText);

                    if (errorText == "")
                        System.IO.File.WriteAllBytes(fileFullname, System.Convert.FromBase64String(fileContent));
                }
                catch (System.Exception ex)
                {
                    errorText += "Создание файла из К2 XML. Не удалось сохранить файл: " + fileFullname + ", причина: " + ex.Message + System.Environment.NewLine;
                }

            if (errorText == "")
                if (!System.IO.File.Exists(fileFullname))
                    errorText += "Создание файла из К2 XML. Файл не найден: " + fileFullname + System.Environment.NewLine;

            return fileShortName;
        }
    }
}