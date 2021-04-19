namespace Helpers
{
    public static class SugarFile
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


        public static string FindByFirstNamePart(string firstNamePart, string fileExtention, string folderPath)
        {
            folderPath = Sugar.NormalizeFolderPath(folderPath);

            if (fileExtention != "" && System.IO.Directory.Exists(folderPath))
            {
                string[] fileNames = System.IO.Directory.GetFiles(
                      folderPath
                    , "*." + fileExtention.Replace(".", "")
                    , System.IO.SearchOption.AllDirectories
                    );

                foreach (string fileName in fileNames)
                    if (fileName.Replace(folderPath, "").Substring(0, firstNamePart.Length) == firstNamePart)
                        return fileName;
            }

            return "";
        }


        public static string GetExtention(string fileName)  //jpg, txt, etc
        {
            string fileExtention = "";
            for (int i = fileName.Length - 1; i >= 0; i--)
            {
                if (fileName[i] != '.')
                    fileExtention = fileName[i].ToString() + fileExtention;
                else
                    break;
            }
            return fileExtention;
        }

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
