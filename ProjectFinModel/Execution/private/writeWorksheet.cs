using OfficeOpenXml;
using System;
using System.Data;

namespace ProjectFinModel
{
    public static partial class Execution
    {
        static void writeWorksheet(ExcelWorksheet worksheet, int projectNumber, DataTable dataTable, ref string errorText)
        {
            if (errorText == "")
                try
                {
                    var department_Cell = Helpers.Excel.GetCellByValue(worksheet, Settings.SQLVariables.Department_Excel);          //ячейка Подразделение
                    var summary_Cell    = Helpers.Excel.GetCellByValue(worksheet, Settings.SQLVariables.LastColumnSummary_Excel);   //ячейка ИТОГО Объем работ(ч/д)
                    var summaryRow_Cell = Helpers.Excel.GetCellByValue(worksheet, Settings.SQLVariables.SummaryRow_Excel);          //ячейка "ИТОГО по всем внутренним ресурсам Банка: "
                    var role_Cell       = Helpers.Excel.GetCellByValue(worksheet, Settings.SQLVariables.Role_Excel);                //ячейка Роль\система\специализация

                    if ((worksheet.Dimension.Start.Row - 1 + worksheet.Dimension.Rows) > 0)
                        worksheet.DeleteRow(summaryRow_Cell.Start.Row + 1, (worksheet.Dimension.Start.Row - 1 + worksheet.Dimension.Rows) - summaryRow_Cell.Start.Row); //удаляем лишние строки

                    var distinctView = new DataView(dataTable);
                    distinctView.Sort = "Department_FullName ASC, Role_System_Specialization ASC";
                    var distinctRows = distinctView.ToTable(true, "TextIdentificator", "DepartmentBlockTypeName", "Role_System_Specialization", "Department_FullName"); //уникальные строки

                    var it_development_Table = distinctRows.Select(String.Format("DepartmentBlockTypeName = '{0}'", Settings.SQLVariables.It_development_SQL)); //разбиение на группы
                    var it_other_Table = distinctRows.Select(String.Format("DepartmentBlockTypeName = '{0}'", Settings.SQLVariables.It_other_SQL)); //разбиение на группы
                    var business_functionality_Table = distinctRows.Select(String.Format("DepartmentBlockTypeName = '{0}'", Settings.SQLVariables.Business_functionality_SQL)); //разбиение на группы
                    distinctRows = null;
                    distinctView = null;

                    ExcelRangeBase it_development_Cell = null; //ячейка Ресурсы ИТ-развития
                    ExcelRangeBase it_other_Cell = null; //ячейка Прочие ИТ-ресурсы
                    ExcelRangeBase business_functionality_Cell = null; //ячейка Бизнес- и функциональные подразделения
                    int lastNonEmptyRowIndex = summaryRow_Cell.Start.Row;

                    fillDepartmentBlock(worksheet   ////Ресурсы ИТ-развития
                        , Settings.SQLVariables.It_development_Excel
                        , it_development_Table
                        , summaryRow_Cell
                        , summary_Cell.Start.Column
                        , role_Cell.Start.Column
                        , department_Cell.Start.Column
                        , ref lastNonEmptyRowIndex
                        , ref it_development_Cell);

                    fillDepartmentBlock(worksheet   //Прочие ИТ-ресурсы
                        , Settings.SQLVariables.It_other_Excel
                        , it_other_Table
                        , summaryRow_Cell
                        , summary_Cell.Start.Column
                        , role_Cell.Start.Column
                        , department_Cell.Start.Column
                        , ref lastNonEmptyRowIndex
                        , ref it_other_Cell);

                    fillDepartmentBlock(worksheet   //Бизнес- и функциональные подразделения
                        , Settings.SQLVariables.Business_functionality_Excel
                        , business_functionality_Table
                        , summaryRow_Cell
                        , summary_Cell.Start.Column
                        , role_Cell.Start.Column
                        , department_Cell.Start.Column
                        , ref lastNonEmptyRowIndex
                        , ref business_functionality_Cell);

                    for (int i = summaryRow_Cell.Start.Column + 2; i < summary_Cell.Start.Column; i++) //вставляем формулу в итоговой верхней строке
                    {
                        string formula = "=" + ((it_development_Cell == null) ? "" : ExcelAddress.GetAddress(it_development_Cell.Start.Row, i));
                        formula += ((formula == "" || it_other_Cell == null) ? "" : "+") + ((it_other_Cell == null) ? "" : ExcelAddress.GetAddress(it_other_Cell.Start.Row, i));
                        formula += ((formula == "" || business_functionality_Cell == null) ? "" : "+") + ((business_functionality_Cell == null) ? "" : ExcelAddress.GetAddress(business_functionality_Cell.Start.Row, i));
                        worksheet.Cells[summaryRow_Cell.Start.Row, i].Formula = formula;
                    }

                    for (int i = summary_Cell.Start.Row + 1; i < lastNonEmptyRowIndex + 1; i++) //вставляем формулу в итоговом столбце
                        worksheet.Cells[i, summary_Cell.Start.Column].Formula =
                            "=SUM(" + ExcelAddress.GetAddress(i, department_Cell.Start.Column + 1) + ":" + ExcelAddress.GetAddress(i, summary_Cell.Start.Column - 1) + ")";

                    var date = new Date();    //ячейки с данными годов и кварталов
                    date.Fill(worksheet, department_Cell, summary_Cell);

                    foreach (DataRow row in dataTable.Rows) //заполняем данные
                    {
                        int rowStart = 0;
                        int rowEnd = 0;
                        string departmentBlockTypeNameData = row["DepartmentBlockTypeName"].ToString();

                        if (departmentBlockTypeNameData == Settings.SQLVariables.It_development_SQL)
                        {
                            rowStart = it_development_Cell.Start.Row + 1;
                            rowEnd = it_development_Cell.Start.Row + it_development_Table.Length;
                        }
                        else if (departmentBlockTypeNameData == Settings.SQLVariables.It_other_SQL)
                        {
                            rowStart = it_other_Cell.Start.Row + 1;
                            rowEnd = it_other_Cell.Start.Row + it_other_Table.Length;
                        }
                        else if (departmentBlockTypeNameData == Settings.SQLVariables.Business_functionality_SQL)
                        {
                            rowStart = business_functionality_Cell.Start.Row + 1;
                            rowEnd = business_functionality_Cell.Start.Row + business_functionality_Table.Length;
                        }
                        else
                            throw new Exception("\nНеизвестный тип блока подразделения (DepartmentBlockTypeName): " + departmentBlockTypeNameData + ", Ресурс: " + row["Role_System_Specialization"].ToString() + "; " + row["Department_FullName"].ToString());

                        int rowRole = getRowResourse(worksheet, row["Role_System_Specialization"].ToString(), row["Department_FullName"].ToString(), rowStart, rowEnd, role_Cell.Start.Column, department_Cell.Start.Column);
                        if (rowRole == -1)
                            throw new Exception("\nНе найден ресурс " + row["Role_System_Specialization"].ToString() + " " + row["Department_FullName"].ToString());

                        try
                        {
                            int column = date.Year[row["ResourceAllocation_Year"].ToString()]
                                .Quarter[Settings.SQLVariables.QuarterPreText + row["ResourceAllocation_Quarter"].ToString()].Start.Column;

                            try //сможем ли преобразовать значение в double
                            {
                                worksheet.Cells[rowRole, column].Value = Convert.ToDouble(row["ResourceAllocation_Allocated"].ToString());
                            }
                            catch
                            {
                                worksheet.Cells[rowRole, column].Value = null;
                            }
                        }
                        catch
                        {
                            string quarter = " " + row["ResourceAllocation_Year"].ToString() + " кв. " + row["ResourceAllocation_Quarter"].ToString() + ";";

                            if (Settings.Variables.UserMessage == "")
                                Settings.Variables.UserMessage = "В файл не записаны кварталы:";

                            if (Settings.Variables.UserMessage.IndexOf(quarter) == -1)
                                Settings.Variables.UserMessage += quarter;
                        }
                    }

                    if (it_development_Cell != null)    //заполням тире в пустые ячейки
                        Helpers.Excel.FillWorksheetEmptyValues(worksheet, it_development_Cell.Start.Row + 1, it_development_Cell.Start.Column + 2, it_development_Cell.Start.Row + it_development_Table.Length, summary_Cell.Start.Column - 1);
                    if (it_other_Cell != null)
                        Helpers.Excel.FillWorksheetEmptyValues(worksheet, it_other_Cell.Start.Row + 1, it_other_Cell.Start.Column + 2, it_other_Cell.Start.Row + it_other_Table.Length, summary_Cell.Start.Column - 1);
                    if (business_functionality_Cell != null)
                        Helpers.Excel.FillWorksheetEmptyValues(worksheet, business_functionality_Cell.Start.Row + 1, business_functionality_Cell.Start.Column + 2, business_functionality_Cell.Start.Row + business_functionality_Table.Length, summary_Cell.Start.Column - 1);

                    if (Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.ExceptionNoDateForFileQuarter))
                        foreach (var yearQuarter in date.Year)
                            foreach (var quarter in yearQuarter.Value.Quarter)
                            {
                                if (dataTable.Select(
                                        String.Format("ResourceAllocation_Year = {0} AND ResourceAllocation_Quarter = {1}"
                                            , yearQuarter.Value.Cell.Value.ToString()
                                            , quarter.Key.ToString().Replace(Settings.SQLVariables.QuarterPreText, "")))
                                        .Length == 0)
                                    throw new Exception("\nВ базе нет данных по кварталу в файле: " + yearQuarter.Value.Cell.Value.ToString() + " кв. " + quarter.Key.ToString());
                            }

                    if (Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsAutoFitColumns))
                        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }
                catch (Exception exception)
                {
                    errorText +=
                          "\nПроект "
                        + projectNumber.ToString()
                        + ": ошибка в выходе excelPackage: "
                        + exception.Message
                        ;
                }
        }
    }
}