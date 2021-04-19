namespace Helpers
{
    public static partial class SugarFile
    {
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
    }
}
