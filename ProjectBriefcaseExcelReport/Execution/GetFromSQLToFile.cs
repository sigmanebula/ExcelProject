using System;
using System.IO;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        public static Helpers.ReturnClass GetFromSQLToFile(string dateStart, string dateEnd, string projectTypeIdList, string stateIdList)
        {
            projectTypeIdList = (projectTypeIdList ?? "").ToString(); //удалить после реализации на UI
            stateIdList = (stateIdList ?? "").ToString(); //удалить после реализации на UI

            dateStart = dateStart.Split(' ')[0];
            dateEnd = dateEnd.Split(' ')[0];
            
            Settings.SQLVariables = new Settings.SQLVariablesClass();
            Settings.Variables = new Settings.VariablesClass();
            Settings.Variables.Refresh();
            string errorText = "";

            using (var connection = Helpers.SugarSQLConnection.GetSQLConnection())
            {
                Helpers.SugarSQLConnection.OpenInUsing(connection);

                Settings.SQLVariables.GetSettings(connection, Settings.SettingsTypeCodeList, ref errorText);    //получаем настройки

                getProductionCalendar(dateStart, dateEnd, connection, ref errorText);  //получаем данные календаря

                string fileShortNameNew = Settings.SQLVariables.NewFileNamePrefix + getPeriodName(true, ref errorText) + "." + Settings.FileExtention; //тут получаем название нового файла

                string fileFullNameNew = Settings.SQLVariables.FolderPath + fileShortNameNew;

                string fileData = "";
                
                Helpers.SugarFile.Copy(Settings.SQLVariables.FolderPath + Settings.SQLVariables.TemplateFileShortName, fileFullNameNew, ref errorText); //копируем файл
                
                if (errorText == "")  //основное действие
                {
                    string productionCalendarIDStart = Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeStart, "ProductionCalendarID");
                    string productionCalendarIDEnd = Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeEnd, "ProductionCalendarID");
                    int startPeriodNumber = int.Parse(Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeStart, "PeriodNumber"));
                    int endPeriodNumber = int.Parse(Settings.Variables.GetProductionCalendar(Settings.ProductionCalendarCodeEnd, "PeriodNumber"));

                    ExcelPackage excelPackage = null;
                    try
                    {
                        excelPackage = new ExcelPackage(new FileInfo(fileFullNameNew));
                        if (excelPackage == null)
                            throw new Exception("\nПустой excelPackage");

                        worksheetHidden_Debug_Write(excelPackage, connection, dateStart, dateEnd, stateIdList, projectTypeIdList);
                        workSheetShown_1_Write(excelPackage, dateEnd);
                        worksheetHidden_2_1_Write(excelPackage, connection, dateEnd, stateIdList, projectTypeIdList);
                        worksheetHidden_2_2_Write(excelPackage, connection, dateStart, dateEnd, stateIdList, projectTypeIdList);
                        worksheetHidden_2_3_Write(excelPackage, connection, dateStart, dateEnd, productionCalendarIDStart, productionCalendarIDEnd);
                        worksheetHidden_2_4_Write(excelPackage, connection, dateStart, dateEnd, productionCalendarIDStart, productionCalendarIDEnd);
                        worksheetHidden_2_5_Write(excelPackage, connection, dateStart, dateEnd, stateIdList, projectTypeIdList);
                        worksheetHidden_2_6_Write(excelPackage, connection, dateStart, dateEnd, stateIdList, projectTypeIdList);
                        workSheetShown_2_Write(excelPackage, dateStart, dateEnd);
                        worksheetHidden_3_0_Write(excelPackage, connection, dateStart, dateEnd, stateIdList, projectTypeIdList);
                        worksheetHidden_3_1_Write(excelPackage, connection, dateStart, dateEnd, stateIdList, projectTypeIdList);
                        workSheetShown_3_Write(excelPackage, dateStart, dateEnd);
                        worksheetHidden_4_0_Write(excelPackage, connection, dateStart, dateEnd, startPeriodNumber, endPeriodNumber, stateIdList, projectTypeIdList);

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
                
                string userMessage = Settings.Variables.UserMessage;
                bool isGetErrorMessage = Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsGetErrorMessage);
                
                Settings.SQLVariables = new Settings.SQLVariablesClass();
                Settings.Variables = new Settings.VariablesClass();
                connection.Close();

                Helpers.SugarFile.DeleteIfExists(fileFullNameNew, "\nОшибка при удалении временного файла: ", ref errorText);
                
                userMessage = Helpers.Sugar.GetUserMessageAndErrorText(userMessage, errorText, isGetErrorMessage);

                GC.Collect();
                return new Helpers.ReturnClass() { FileData = fileData, UserMessage = userMessage };
            }
        }
    }
}
