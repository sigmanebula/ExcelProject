using System;
using System.IO;
using OfficeOpenXml;

namespace ProjectQuarteryReport
{
    public static partial class Execution
    {
        public static Helpers.ReturnClass GetFromSQLToFile(string projectID, string productionCalendarID)
        {
            Settings.SQLVariables = new Settings.SQLVariablesClass();
            Settings.Variables    = new Settings.VariablesClass();
            Settings.Variables.Refresh();
            string errorText = "";

            using (var connection = Helpers.SugarSQLConnection.GetSQLConnection())
            {
                string fileData = "";

                try
                {
                    int projectIDInt                        = int.Parse(projectID);
                    int productionCalendarIDInt             = int.Parse(productionCalendarID);
                    Settings.Variables.ProjectID            = projectID;
                    Settings.Variables.ProductionCalendarID = productionCalendarID;
                }
                catch (Exception exception)
                {
                    throw new Exception("\nОшибка во входных данных: projectID = " + projectID + ", productionCalendarID = " + productionCalendarID + ", текст ошибки: " + exception.Message);
                }

                Helpers.SugarSQLConnection.OpenInUsing(connection);

                Settings.SQLVariables.GetSettings(connection, Settings.SettingsTypeCodeList, ref errorText);
                
                getProductionCalendarData(connection, ref errorText);

                string fileShortNameNew = getFileShortNameNew(connection, ref errorText);

                string fileFullNameNew = Settings.SQLVariables.FolderPath + fileShortNameNew;
                
                Helpers.SugarFile.Copy(Settings.SQLVariables.FolderPath + Settings.SQLVariables.TemplateFileShortName, fileFullNameNew, ref errorText);
                
                if (errorText == "")  //основное действие
                {
                    ExcelPackage excelPackage = null;
                    try
                    {
                        excelPackage = new ExcelPackage(new FileInfo(fileFullNameNew));
                        if (excelPackage == null)
                            throw new Exception("\nПустой excelPackage");

                        var worksheetMainData = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetMainData_Name);

                        writeWorksheetMainData(connection, worksheetMainData, ref errorText);

                        excelPackage.Save();
                    }
                    catch (Exception ex)
                    {
                        errorText += "\nОшибка в файле. " + ex.Message;
                    }
                    finally
                    {
                        excelPackage.Dispose();
                    }

                    fileData = Helpers.SugarFile.GetK2Xml(fileFullNameNew, fileShortNameNew, "", ref errorText);
                }
                
                Helpers.SugarFile.DeleteIfExists(fileFullNameNew, "\nОшибка при удалении временного файла: ", ref errorText);
                
                string userMessage = Settings.Variables.UserMessage;
                bool isGetErrorMessage = Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsGetErrorMessage);

                Settings.SQLVariables = new Settings.SQLVariablesClass();
                Settings.Variables = new Settings.VariablesClass();
                connection.Close();

                GC.Collect();

                userMessage = Helpers.Sugar.GetUserMessageAndErrorText(userMessage, errorText, isGetErrorMessage);
                
                return new Helpers.ReturnClass() { FileData = fileData, UserMessage = userMessage };
            }
        }
    }
}
