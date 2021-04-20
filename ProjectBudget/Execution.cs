using System.Linq;
using OfficeOpenXml;
using System.IO;
using System;
using System.Data;
using System.Data.SqlClient;

namespace ProjectBudget
{
    public static class Execution
    {
        static string getXMLDataFromFile(string filePath, int projectNumber, ref string errorText)
        {
            string resultXML = "";
            ExcelPackage excelPackage = null;

            var cellsDictionary = new System.Collections.Generic.Dictionary<string, Helpers.CellsDictionaryElement>(); //словарь элементов ячеек
            int id = 0; //заполняем словарь ячеек - переводим таблицу в плоскую структуру

            if (errorText == "")
                try
                {
                    var dataSet = Helpers.Sugar.GetDataSetFromXML(Settings.SQLVariables.cells);
                    foreach (DataRow row in dataSet.Tables[0].Rows)
                        foreach (DataColumn column in dataSet.Tables[0].Columns)
                        {
                            cellsDictionary.Add(id.ToString(), new Helpers.CellsDictionaryElement() { Id = id, Location = row[column].ToString(), Value = "" });
                            id++;
                        }

                    excelPackage = new ExcelPackage(new FileInfo(filePath));
                    var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetName);
                    foreach (var cell in cellsDictionary)
                    {
                        try
                        {
                            if (cell.Value.Location != ExcelAddress.GetAddress(worksheet.Cells[cell.Value.Location].Start.Row, worksheet.Cells[cell.Value.Location].Start.Column))
                                throw new Exception();
                            else
                                cell.Value.Value = (worksheet.Cells[cell.Value.Location].Value ?? "").ToString();
                        }
                        catch
                        {
                            cell.Value.Value = cell.Value.Location; //если адрес ячейки найти невозможно, то её значение будет равно её адресу
                        }
                    }

                    id = 0; //заполняем датасет из плоского словаря ячеек
                    foreach (DataRow row in dataSet.Tables[0].Rows)
                        foreach (DataColumn column in dataSet.Tables[0].Columns)
                        {
                            row[column] = cellsDictionary[id.ToString()].Value;
                            id++;
                        }

                    resultXML = dataSet.GetXml();

                    if (resultXML == "")
                        throw new Exception("Не удалось собрать XML!");
                }
                catch (Exception exception)
                {
                    errorText = "\nОшибка в общей информации бюджета проекта " + projectNumber.ToString() + ": " + exception.Message + Environment.NewLine;
                }
                finally
                {
                    if (excelPackage != null)
                        excelPackage.Dispose();
                }
            
            return resultXML;
        }

        static string getXMLDataFromFileFull(int projectID, int projectNumber, string filePath, ref string errorText)
        {
            string resultXML = "";
            ExcelPackage excelPackage = null;

            if (errorText == "")
                try
                {
                    excelPackage = new ExcelPackage(new FileInfo(filePath));
                    var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetName);
                    var dataTable = new DataTable();
                    int columnIndexStart = Helpers.Excel.GetWorksheetColumnIndexByName(worksheet, Settings.SQLVariables.ColumnNameStartFull);
                    int columnIndexEnd = Helpers.Excel.GetWorksheetColumnIndexByName(worksheet, Settings.SQLVariables.ColumnNameEndFull);
                    Helpers.SugarDataTable.AddColumn(dataTable, columnIndexEnd + 1 - columnIndexStart);
                    int rowStart = int.Parse(Settings.SQLVariables.RowStartFull);

                    for (int i = rowStart; i <= worksheet.Dimension.End.Row; i++)   //заполняем данными dataTable
                        if ((worksheet.Cells[Settings.SQLVariables.ColumnNameDeleteFull + i.ToString()].Value ?? "").ToString() != Settings.SQLVariables.ColumnValueDeleteFull)
                        {
                            dataTable.Rows.Add();
                            for (int j = columnIndexStart; j <= columnIndexEnd; j++)
                                dataTable.Rows[dataTable.Rows.Count - 1][j - 1] = worksheet.Cells[i, j].Value;
                        }

                    Helpers.SugarDataTable.AddColumn(dataTable, Settings.XMLColumnNameTypeNameFull);
                    Helpers.SugarDataTable.AddColumn(dataTable, Settings.XMLColumnNameIsSummaryTypeFull);
                    Helpers.SugarDataTable.AddColumn(dataTable, Settings.XMLColumnNameProjectIDFull);
                    string projectBudgetFullTypeName = "";
                    for (int i = 0; i < dataTable.Rows.Count; i++)  //обновляем данные dataTable, группируем
                    {
                        string rowType = projectBudgetFullGetRowType(dataTable, i);
                        string headerTemp = dataTable.Rows[i][0].ToString();
                        if (rowType == "TypeName")
                        {
                            projectBudgetFullTypeName = headerTemp;
                            dataTable.Rows.RemoveAt(i);
                            i--;
                        }
                        else
                        {
                            if (headerTemp == Settings.SQLVariables.ApprovingTextFull || headerTemp == Settings.SQLVariables.SummaryTextFull || headerTemp == Settings.SQLVariables.MasteringTextFull) //смотрим на последние 3 строки
                                projectBudgetFullTypeName = headerTemp;

                            dataTable.Rows[i][Settings.XMLColumnNameTypeNameFull] = projectBudgetFullTypeName;
                            dataTable.Rows[i][Settings.XMLColumnNameIsSummaryTypeFull] = (rowType == "Summary") ? true : false;
                            dataTable.Rows[i][Settings.XMLColumnNameProjectIDFull] = projectID;
                        }
                    }
                    dataTable.TableName = "data";
                    DataSet dataSet = new DataSet();
                    dataSet.DataSetName = "dataSet";
                    dataSet.Tables.Add(dataTable);
                    resultXML = dataSet.GetXml();

                    if (resultXML == "")
                        throw new Exception("Не удалось собрать XML!");
                }
                catch (Exception exception)
                {
                    errorText = "\nОшибка в детальной информации бюджета проекта " + projectNumber.ToString() + ": " + exception.Message + Environment.NewLine;
                }
                finally
                {
                    if (excelPackage != null)
                        excelPackage.Dispose();
                }

            return resultXML;
        }

        static string projectBudgetFullGetRowType(DataTable dataTable, int rowIndex)
        {
            string result = "";
            bool hasFirstColumnValue = false;
            bool hasSecondColumnValue = false;
            bool hasOtherColumnsValues = false;
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                if (i == 0 && (dataTable.Rows[rowIndex][i] ?? "").ToString() != "")
                    hasFirstColumnValue = true;
                else if (i == 1 && (dataTable.Rows[rowIndex][i] ?? "").ToString() != "")
                    hasSecondColumnValue = true;
                else if ((dataTable.Rows[rowIndex][i] ?? "").ToString() != "")
                {
                    hasOtherColumnsValues = true;
                    break;
                }
            }
            if (hasFirstColumnValue && !hasSecondColumnValue && !hasOtherColumnsValues)
                result = "TypeName";
            else if (hasFirstColumnValue && !hasSecondColumnValue && hasOtherColumnsValues)
                result = "Summary";
            return result;
        }

        public static void ReadFromFileToSQLTable(string projectIDList)
        {
            Settings.SQLVariables = new SQLVariablesClass();
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

                    Settings.SQLVariables = new SQLVariablesClass();

                    connection.Close();
                }

            return errorText;
        }

        static void writeSQLProjectBudgetCommon(SqlConnection connection, int projectID, int projectNumber, string xml, ref string errorText)
        {
            if (errorText == "")
                try
                {
                    if (xml == "")
                        throw new System.Exception("xml пуст!");

                    using (var command = new SqlCommand())
                    {
                        command.Connection = connection;
                        command.CommandText = "[ITProject].[spCreateProjectBudget]";
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                        command.Parameters.AddWithValue("@XML", xml);
                        command.Parameters.AddWithValue("@ProjectID", projectID);
                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception exception)
                {
                    errorText += "\nОшибка записи общей информации по бюджету проекта " + projectNumber.ToString() + ": " + exception.Message;
                }
        }

        static void writeSQLProjectBudgetFull(SqlConnection connection, int projectNumber, string xml, ref string errorText)
        {
            if (errorText == "")
                try
                {
                    if (xml == "")
                        throw new System.Exception("xml пуст!");

                    using (var command = new SqlCommand()) //записываем
                    {
                        command.Connection = connection;
                        command.CommandText = "[ITProject].[spCreateProjectBudgetFull]";
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                        command.Parameters.AddWithValue("@XML", xml);
                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception exception)
                {
                    errorText += "\nОшибка записи полной информации по бюджету проекта " + projectNumber.ToString() + ": " + exception.Message;
                }
        }

    }
}
