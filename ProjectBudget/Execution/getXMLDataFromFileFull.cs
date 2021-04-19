using System.Linq;
using OfficeOpenXml;
using System.IO;
using System;
using System.Data;

namespace ProjectBudget
{
    public static partial class Execution
    {
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

    }
}
