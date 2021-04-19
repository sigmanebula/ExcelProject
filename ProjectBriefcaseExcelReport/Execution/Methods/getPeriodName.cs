using System;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static string getPeriodName(bool isLowerQuarterWord, ref string errorText)
        {
            string result = "";

            if (errorText == "")
                try
                {
                    string quarterWord = (isLowerQuarterWord) ? " квартал " : " КВАРТАЛ ";

                    string startWord =
                          Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeStart, "Quarter")
                        + quarterWord
                        + Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeStart, "Year");

                    string endWord =
                          Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeEnd, "Quarter")
                        + quarterWord
                        + Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeEnd, "Year");

                    if (startWord == endWord)
                        result += startWord;
                    else
                        result += startWord + " - " + endWord;
                }
                catch (Exception ex)
                {
                    errorText += "Ошибка получения периода. " + ex.Message;
                }

            return result;
        }
        
        static string getPeriodName(bool isLowerQuarterWord)
        {
            string errorText = "";
            string result = getPeriodName(isLowerQuarterWord, ref errorText);

            if (errorText != "")
                throw new System.Exception(errorText);

            return result;
        }
    }
}
