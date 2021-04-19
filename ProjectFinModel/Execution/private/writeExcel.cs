using OfficeOpenXml;
using System.IO;
using System;
using System.Data;

namespace ProjectFinModel
{
    public static partial class Execution
    {
        static void writeExcel(int projectNumber, DataTable dataTable, string fileName, ref string errorText)
        {
            if (errorText == "")
            {
                bool hasPassword = true;
                ExcelPackage excelPackage = new ExcelPackage();
                
                try
                {
                    try
                    {
                        excelPackage = new ExcelPackage(new FileInfo(fileName), Settings.SQLVariables.ExcelPassword);
                    }
                    catch
                    {
                        hasPassword = false;
                        excelPackage = new ExcelPackage(new FileInfo(fileName));
                    }
                }
                catch(Exception exception)
                {
                    errorText += "\nОшибка при открытии файла excel: " + exception.Message;
                }
                
                if (excelPackage == null && errorText == "")
                    errorText += "\nОшибка в файле, пустой excelPackage";
                
                Helpers.Excel.CheckExcelPackage(excelPackage, ref errorText);

                ExcelWorksheet worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetName, ref errorText);

                writeWorksheet(worksheet, projectNumber, dataTable, ref errorText);

                GC.Collect(); //без него не работает последующее сохранение!

                if (errorText == "")
                {
                    try
                    {
                        if (hasPassword)
                            excelPackage.Save(Settings.SQLVariables.ExcelPassword);
                        else
                            excelPackage.Save();
                    }
                    catch(Exception exception)
                    {
                        errorText +=
                            "\nНе удалось сохранить файл Excel, причина: " + exception.Message
                            + "; есть пароль? - " + ((hasPassword) ? "да" : "нет")
                            ;
                    }
                }

                excelPackage.Dispose();

                GC.Collect(); //без него не работает освобождение файла от процесса!
            }
        }
    }
}
