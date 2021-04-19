using System.Linq;
using OfficeOpenXml;
using System.IO;
using System;
using System.Data;

namespace ProjectBudget
{
    public static partial class Execution
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
    }
}
