namespace Helpers
{
    public static class Sugar
    {
       
        public static string NormalizeFolderPath(string path)
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
                    
                }

            return path;
        }


        public static string GetUserMessageAndErrorText(string userMessage)
        {
            if (!string.IsNullOrEmpty(userMessage))
                userMessage = "";

            return userMessage;
        }

        public static System.Data.DataSet GetDataSetFromXML(string xml)
        {
            System.Data.DataSet dataSet = new System.Data.DataSet();

                try
                {
                    var stream = new System.IO.MemoryStream();
                    var writer = new System.IO.StreamWriter(stream);

                    writer.Write(xml);
                    writer.Flush();
                    stream.Position = 0;

                    dataSet.ReadXml(stream);
                }
                catch (System.Exception exception)
                {
                    exception.Message = "Создание датасета из К2 XML. Ошибка: " + exception.Message + System.Environment.NewLine;
                }

            return dataSet;
        }


        public static bool ConvertStringToBool(string word, bool isExceptionIfNotInCase)
        {
            word = (word ?? "").ToLower();

            bool result = false;

            switch (word)
            {
                case "1":
                case "yy":
                case "y":
                case "yes":
                case "+":
                case "да":
                case "д":
                case "true":
                case "t":
                    result = true; break;

                case "0":
                case "nn":
                case "n":
                case "no":
                case "-":
                case "нет":
                case "н":
                case "false":
                case "f":
                    result = false; break;

                default:
                    if (isExceptionIfNotInCase)
                        throw new System.Exception("Ошибка конверсии в булево значение: " + word);
                    else
                        break; //false
            }

            return result;
        }

        public static bool ConvertStringToBool(string word)
        {
            return ConvertStringToBool(word, false);
        }

        public static string ConvertDateToFormat(string date, char spitFrom, char splitTo, string codeFrom, string codeTo)  //date: 2020-01-01
        {
            string[] dateParts = new string[3]; //YYYY, MM, DD

            string[] datePartsFromTemp = date.Split(spitFrom); //unknown

            string result = "";

            if (datePartsFromTemp.Length == 3)
            {
                switch (codeFrom)
                {
                    case "YYYY.MM.DD": dateParts = new string[3] { datePartsFromTemp[0], datePartsFromTemp[1], datePartsFromTemp[2] }; break;
                    case "DD.MM.YYYY": dateParts = new string[3] { datePartsFromTemp[2], datePartsFromTemp[1], datePartsFromTemp[0] }; break;
                    default: throw new System.Exception("Ошибка конверсии даты, неверный формат исходной: " + codeFrom);
                }

                switch (codeTo)
                {
                    case "YYYY.MM.DD": result = datePartsFromTemp[0] + splitTo + datePartsFromTemp[1] + splitTo + datePartsFromTemp[2]; break;
                    case "DD.MM.YYYY": result = datePartsFromTemp[2] + splitTo + datePartsFromTemp[1] + splitTo + datePartsFromTemp[0]; break;
                    default: throw new System.Exception("Ошибка конверсии даты, неверный формат целевой: " + codeTo);
                }
            }
            else
                throw new System.Exception("Ошибка конверсии даты: " + date);

            return result;
        }


    }
}
