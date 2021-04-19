using System;
using OfficeOpenXml;

namespace Helpers
{
    public static partial class Excel
    {
        public static void CheckExcelPackage(ExcelPackage excelPackage, ref string errorText)
        {
            if (errorText == "")
            {
                try
                {
                    var worksheet = excelPackage.Workbook.Worksheets[1];
                }
                catch(Exception exception)
                {
                    errorText += "Ошибка в Excel: " + exception.Message;
                }
            }
        }

        public static void CheckExcelPackage(ExcelPackage excelPackage)
        {
            string errorText = "";

            CheckExcelPackage(excelPackage, ref errorText);

            if (errorText != "")
                throw new Exception(errorText);
        }
    }
}
