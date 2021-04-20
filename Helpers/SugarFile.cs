namespace Helpers
{
  public static class SugarFile
  {
    public static void Copy(string fileFullNameOld, string fileFullNameNew)
    {
      //процедура нужна для перехвата и расшифровки ошибок, получение подробного текста ошибок актуально.
      try
      {
        SugarFile.DeleteIfExists(fileFullNameNew);

        if (!System.IO.File.Exists(fileFullNameOld))
          throw new System.Exception("\nфайл не найден: " + fileFullNameOld);

        System.IO.File.Copy(fileFullNameOld, fileFullNameNew);
      }
      catch (System.Exception ex)
      {
        ex.Message =
            "\nОшибка при копировании файла: "
            + ex.Message
            + "\nfileFullNameOld: "
            + fileFullNameOld
            + "\nfileFullNameNew"
            + fileFullNameNew
            ;
      }
    }

    public static string CreateFromK2Xml(string fileK2XML, string folderPath, string fileExtention)
    {
      string fileShortName = "";
      string fileFullname = "";
      string fileContent = "";

      fileExtention = fileExtention.Replace(".", "");
      folderPath = Sugar.NormalizeFolderPath(folderPath);

      var dataSet = Sugar.GetDataSetFromXML(fileK2XML);

      fileShortName = (dataSet.Tables[0].Rows[0]["name"] ?? "").ToString();
      fileContent = (dataSet.Tables[0].Rows[0]["content"] ?? "").ToString();

      if (GetExtention(fileShortName) != fileExtention)
      {
        throw new System.Exception("Создание файла из К2 XML. Расширение файла должно быть ." + fileExtention + System.Environment.NewLine);
      }

      fileFullname = folderPath + fileShortName;

      try
      {
        DeleteIfExists(fileFullname);


        System.IO.File.WriteAllBytes(fileFullname, System.Convert.FromBase64String(fileContent));
      }
      catch (System.Exception ex)
      {
        ex.Message = "Создание файла из К2 XML. Не удалось сохранить файл: " + fileFullname + ", причина: " + ex.Message + System.Environment.NewLine;
      }


      if (!System.IO.File.Exists(fileFullname))
      {
        throw new System.Exception("Создание файла из К2 XML. Файл не найден: " + fileFullname + System.Environment.NewLine);
      }


      return fileShortName;
    }

    public static void DeleteIfExists(string fileFullName)
    {
      if (fileFullName != "" && System.IO.File.Exists(fileFullName))
      {
        try
        {
          System.IO.File.Delete(fileFullName);
        }
        catch (System.Exception ex)
        {
          ex.Message = "\nНе удалось удалить " + fileFullName + ", причина: " + ex.Message;
        }
      }
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

    public static string GetK2Xml(string fileFullName, string fileShortName)
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
          ex.Message = "\nНе удалось создать K2 XML на основе файла: " + fileFullName + ", причина: " + ex.Message;
          return "";
        }
    }

  }
}
