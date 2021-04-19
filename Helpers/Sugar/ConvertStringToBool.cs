namespace Helpers
{
    public static partial class Sugar
    {
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
    }
}

