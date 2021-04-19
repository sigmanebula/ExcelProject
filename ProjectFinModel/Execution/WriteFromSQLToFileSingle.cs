using System;
using System.Data;
using System.IO;

namespace ProjectFinModel
{
    public static partial class Execution
    {
        public static Helpers.ReturnClass WriteFromSQLToFileSingle(string projectIDList, string fileData)
        {
            projectIDList = Helpers.Sugar.RemoveStringLastChars(projectIDList, ", ");
            Settings.SQLVariables = new Settings.SQLVariablesClass();
            Settings.Variables = new Settings.VariablesClass();
            Settings.Variables.Refresh();
            string errorText = "";

            DataTable dataTableReportRaw = new DataTable();

            try
            {
                using (var connection = Helpers.SugarSQLConnection.GetSQLConnection())
                {
                    Helpers.SugarSQLConnection.OpenInUsing(connection);

                    Settings.ProjectIDNumber.GetData(connection, projectIDList, ref errorText);  //получаем все ProjectID\Number проектов из входной строки

                    if (errorText == "" && Settings.ProjectIDNumber.List.Count != 1)
                        errorText += "\nОшибка входных данных: ID проекта не найден: " + projectIDList;

                    Settings.SQLVariables.GetSettings(connection, Settings.SettingsTypeCodeList, ref errorText);    //получаем настройки

                    dataTableReportRaw = getReportProjectResourceIntensityDataTable(Settings.ProjectIDNumber.List[0], connection, ref errorText);

                    try
                    {
                        connection.Close();
                    }
                    catch (Exception exception)
                    {
                        throw new Exception("\nОшибка закрытия коннекции: " + exception.Message);
                    }
                }
            }
            catch(Exception exception)
            {
                errorText += "\nОшибка SQLConnection: " + exception.Message;
            }

            DataTable dataTableReport = deleteJunkDataTableReport(dataTableReportRaw, ref errorText);
            dataTableReportRaw = null;

            GC.Collect(); //для connection

            if (Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsDebugSQL))
            {
                errorText +=
                      Settings.ProjectIDNumber.GetStringListData()
                    + "\ndataTableReport rows count " + dataTableReport.Rows.Count.ToString()
                    + Settings.SQLVariables.GetStringListSettings()
                    ;
            }
            
            if (errorText == "" && dataTableReport.Rows.Count == 0)
                errorText += "В базе нет данных для заполнения проекта " + Settings.ProjectIDNumber.List[0].ProjectNumber.ToString();
            
            string fileShortName = Helpers.SugarFile.CreateFromK2Xml(fileData, Settings.SQLVariables.FolderPath, Settings.FileExtention, ref errorText);
            
            string fileName = Helpers.Sugar.NormalizeFolderPath(Settings.SQLVariables.FolderPath, ref errorText) + fileShortName;

            if (errorText == "" && !File.Exists(fileName))
                errorText += "\nВременный файл для работы не найден: " + fileName;

            fileData = "";
            
            if (errorText == "")
                writeExcel(
                      Settings.ProjectIDNumber.List[0].ProjectNumber
                    , dataTableReport
                    , fileName
                    , ref errorText
                    );
            
            fileData = Helpers.SugarFile.GetK2Xml(fileName, fileShortName, "", ref errorText);
                
            Helpers.SugarFile.DeleteIfExists(fileName, "\nВыходные данные: ", ref errorText);
            
            string userMessage = Settings.Variables.UserMessage;
            bool isGetErrorMessage = Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsGetErrorMessage);

            Settings.SQLVariables = new Settings.SQLVariablesClass();
            Settings.Variables = new Settings.VariablesClass();
            
            GC.Collect();
            
            if (errorText != "")
            {
                if (isGetErrorMessage)
                    userMessage += errorText;
                else
                    throw new Exception(errorText);
            }
            
            return new Helpers.ReturnClass() { FileData = fileData, UserMessage = userMessage };
        }
    }
}
