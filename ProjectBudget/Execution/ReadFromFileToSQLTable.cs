using System;

namespace ProjectBudget
{
    public static partial class Execution
    {
        public static void ReadFromFileToSQLTable(string projectIDList)
        {
            Settings.SQLVariables = new Settings.SQLVariablesClass();
            string errorText = "";

            projectIDList = Helpers.Sugar.RemoveStringLastChars(projectIDList, ", ", ref errorText);

            errorText = readFromFileToSQLTableUseConnection(projectIDList, errorText);

            GC.Collect();

            if (!string.IsNullOrEmpty(errorText))
            {
                //System.IO.File.WriteAllText(@"D:\VSProject\errorText.txt", errorText);
                throw new System.Exception(errorText);
            }


        }

        static string readFromFileToSQLTableUseConnection(string projectIDList, string errorText)
        {
            if (errorText == "")
                using (var connection = Helpers.SugarSQLConnection.GetSQLConnection())
                {
                    Helpers.SugarSQLConnection.OpenInUsing(connection, ref errorText);

                    Settings.SQLVariables.GetSettings(connection, Settings.SettingsTypeCodeList, ref errorText); //получаем настройки

                    Settings.ProjectIDNumber.GetData(connection, projectIDList, ref errorText);  //получаем все ProjectID\Number проектов из входной строки

                    if (errorText == "")
                        foreach (var Project in Settings.ProjectIDNumber.List) //получаем данные и записываем
                        {
                            string filePath = Helpers.SugarFile.FindByFirstNamePart(
                                 Project.ProjectNumber.ToString() + Settings.SQLVariables.ProjectNumberDelimeter
                                , Settings.FileExtention
                                , Settings.SQLVariables.FolderPath
                                );

                            if (filePath != "") //если файл не пустой, то он найден, можно брать инфу
                            {
                                string xmlBudgetCommon = getXMLDataFromFile(filePath, Project.ProjectNumber, ref errorText); //чтение из файла и упаковка в XML

                                writeSQLProjectBudgetCommon(connection, Project.ProjectID, Project.ProjectNumber, xmlBudgetCommon, ref errorText); //запись общей информации по бюджету в базу

                                string xmlBudgetFull = getXMLDataFromFileFull(Project.ProjectID, Project.ProjectNumber, filePath, ref errorText); //чтение из файла и упаковка в XML

                                writeSQLProjectBudgetFull(connection, Project.ProjectNumber, xmlBudgetFull, ref errorText); //запись детальной информации по бюджету в базу
                            }
                        }

                    Settings.SQLVariables = new Settings.SQLVariablesClass();

                    connection.Close();
                }

            return errorText;
        }

    }
}