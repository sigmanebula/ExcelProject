using System;
using System.Data;

namespace ProjectQuarteryReport
{
    public static partial class Settings
    {
        public class VariablesClass : Helpers.VariablesClass
        {
            public string ProjectNumberShortName { get; set; }
            
            public string ProjectID { get; set; }
            public string ProductionCalendarID { get; set; }
            public DataTable ProductionCalendar { get; set; }

            public string GetProductionCalendar(string code, string columnName)
            {
                string result = "";
                for (int i = 0; i < ProductionCalendar.Rows.Count; i++)
                    if (Convert.ToString(ProductionCalendar.Rows[i]["Code"] ?? "") == code)
                        return Convert.ToString(ProductionCalendar.Rows[i][columnName] ?? "");
                return result;
            }
        }
    }
}
