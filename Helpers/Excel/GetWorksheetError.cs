namespace Helpers
{
    public static partial class Excel
    {
        public static string GetWorksheetError(string exceptionMessage, string workSheetName)
        {
            return "\nОшибка в листе " + workSheetName + ", текст ошибки: " + exceptionMessage;
        }
    }
}
