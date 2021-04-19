namespace Helpers
{
    public static partial class Sugar
    {
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
