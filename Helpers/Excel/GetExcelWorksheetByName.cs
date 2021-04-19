using System;
using System.Linq;
using OfficeOpenXml;

namespace Helpers
{
    public static partial class Excel
    {
        public static ExcelWorksheet GetExcelWorksheetByName(ExcelPackage excelPackage, string worksheetName, ref string errorText)
        {
            ExcelWorksheet worksheet = null;

            if (errorText == "")
                try
                {
                    worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault(w => w.Name == worksheetName);

                    if (worksheet == null)
                        throw new Exception("worksheet is null ");
                }
                catch(Exception exception)
                {
                    errorText += "Вкладка excel не найдена: " + worksheetName + ", " + exception.Message;
                }
            
            return worksheet;
        }
        
        public static ExcelWorksheet GetExcelWorksheetByName(ExcelPackage excelPackage, string worksheetName)
        {
            string errorText = "";
            ExcelWorksheet worksheet = GetExcelWorksheetByName(excelPackage, worksheetName, ref errorText);

            if (errorText != "")
                throw new Exception(errorText);

            return worksheet;
        }
    }
}
