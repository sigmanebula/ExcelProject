namespace Helpers
{
    public static partial class SugarFile
    {
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
    }
}