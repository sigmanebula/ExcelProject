namespace Helpers
{
    public static partial class Excel
    {
        public static System.Drawing.Color GetExcelColor(OfficeOpenXml.Style.ExcelColor excelColor)
        {
            return GetExcelColor(excelColor, System.Drawing.Color.Transparent);
        }
        
        public static System.Drawing.Color GetExcelColor(OfficeOpenXml.Style.ExcelColor excelColor, System.Drawing.Color defaultColor)
        {
            try
            {
                return System.Drawing.ColorTranslator.FromHtml("#" + excelColor.Rgb.ToString());
                //return System.Drawing.ColorTranslator.FromHtml(excelColor.LookupColor());
            }
            catch
            {
                return defaultColor;
            }
        }
    }
}