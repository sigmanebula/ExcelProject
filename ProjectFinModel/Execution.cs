using OfficeOpenXml;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace ProjectFinModel
{
    public static class Execution
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

        static void writeWorksheet(ExcelWorksheet worksheet, int projectNumber, DataTable dataTable, ref string errorText)
        {
            if (errorText == "")
                try
                {
                    var department_Cell = Helpers.Excel.GetCellByValue(worksheet, Settings.SQLVariables.Department_Excel);          //ячейка Подразделение
                    var summary_Cell = Helpers.Excel.GetCellByValue(worksheet, Settings.SQLVariables.LastColumnSummary_Excel);   //ячейка ИТОГО Объем работ(ч/д)
                    var summaryRow_Cell = Helpers.Excel.GetCellByValue(worksheet, Settings.SQLVariables.SummaryRow_Excel);          //ячейка "ИТОГО по всем внутренним ресурсам Банка: "
                    var role_Cell = Helpers.Excel.GetCellByValue(worksheet, Settings.SQLVariables.Role_Excel);                //ячейка Роль\система\специализация

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

        static DataTable deleteJunkDataTableReport(DataTable dataTableSource, ref string errorText)
        {
            DataTable dataTable = new DataTable();

            if (errorText == "")
                try
                {
                    for (int i = 0; i > -1 && i < dataTableSource.Rows.Count; i++)
                        if (
                            (bool)(dataTableSource.Rows[i]["IsOldAllocation"] ?? false) == true
                            || dataTableSource.Rows[i]["ResourceAllocation_Allocated"] == null
                            || Convert.ToString(dataTableSource.Rows[i]["ResourceAllocation_Allocated"] ?? "") == "0"
                            || Convert.ToString(dataTableSource.Rows[i]["ResourceAllocation_TypeCode"] ?? "") == "empty"
                            || Convert.ToString(dataTableSource.Rows[i]["ResourceAllocation_TypeCode"] ?? "") == ""
                            )
                        {
                            dataTableSource.Rows[i].Delete();
                        }

                    dataTableSource.AcceptChanges();

                    dataTable = Helpers.SugarDataTable.CopyDataTableByColumnList(dataTableSource, Settings.SQLVariables.ReportProjectResourceIntensityColumnList);
                }
                catch (Exception ex)
                {
                    errorText += "\nОшибка очистки данных таблицы отчёта, причина: " + ex.Message;
                }

            return dataTable;
        }

        static void fillDepartmentBlock(
            ExcelWorksheet worksheet
            , string headerCellValue
            , DataRow[] dataTableRows
            , ExcelRangeBase summaryRow_Cell
            , int summary_Cell_Start_Column
            , int role_Cell_Start_Column
            , int department_Cell_Start_Column
            , ref int lastNonEmptyRowIndex
            , ref ExcelRangeBase headerBlockCell
            )
        {
            if (dataTableRows.Length > 0)
            {
                worksheet.InsertRow(lastNonEmptyRowIndex + 1, 1);

                Helpers.Excel.CellStyleClass cellStyle = new Helpers.Excel.CellStyleClass();
                cellStyle.SetPropertiesFromCell(summaryRow_Cell);
                cellStyle.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                cellStyle.BorderLeftStyle = OfficeOpenXml.Style.ExcelBorderStyle.None;
                cellStyle.BorderRightStyle = OfficeOpenXml.Style.ExcelBorderStyle.None;
                cellStyle.FillRange(worksheet, ExcelAddress.GetAddress(lastNonEmptyRowIndex + 1, summaryRow_Cell.Start.Column, lastNonEmptyRowIndex + 1, summary_Cell_Start_Column));

                worksheet.Row(lastNonEmptyRowIndex + 1).Height = worksheet.DefaultRowHeight;
                worksheet.Cells[lastNonEmptyRowIndex + 1, summaryRow_Cell.Start.Column].Value = headerCellValue;
                headerBlockCell = worksheet.Cells[lastNonEmptyRowIndex + 1, summaryRow_Cell.Start.Column];
                lastNonEmptyRowIndex = headerBlockCell.Start.Row;

                worksheet.InsertRow(lastNonEmptyRowIndex + 1, dataTableRows.Length);

                foreach (DataRow row in dataTableRows)  //заполняем заголовки обычных строк
                {
                    lastNonEmptyRowIndex++;
                    worksheet.Cells[lastNonEmptyRowIndex, role_Cell_Start_Column].Value = row["Role_System_Specialization"].ToString();
                    worksheet.Cells[lastNonEmptyRowIndex, department_Cell_Start_Column].Value = row["Department_FullName"].ToString();
                }

                try
                {
                    worksheet.Cells[ExcelAddress.GetAddress(headerBlockCell.Start.Row, summaryRow_Cell.Start.Column, headerBlockCell.Start.Row, summaryRow_Cell.Start.Column + 1)].Merge = true;
                }
                catch { }

                cellStyle.FillBackgroundColor = System.Drawing.Color.Transparent;
                cellStyle.FontColor = System.Drawing.Color.Black;
                cellStyle.FontBold = false;
                cellStyle.BorderLeftStyle = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                cellStyle.BorderRightStyle = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                cellStyle.BorderTopColor = System.Drawing.Color.LightGray;
                cellStyle.BorderBottomColor = System.Drawing.Color.LightGray;
                cellStyle.BorderLeftColor = System.Drawing.Color.LightGray;
                cellStyle.BorderRightColor = System.Drawing.Color.LightGray;

                cellStyle.FillRange(worksheet, ExcelAddress.GetAddress(headerBlockCell.Start.Row + 1, summaryRow_Cell.Start.Column, lastNonEmptyRowIndex, summaryRow_Cell.Start.Column + 1));   //роль и департамент
                cellStyle.FillRange(worksheet, ExcelAddress.GetAddress(headerBlockCell.Start.Row + 1, summaryRow_Cell.Start.Column + 2, lastNonEmptyRowIndex, summary_Cell_Start_Column));  //аллокации

                for (int i = headerBlockCell.Start.Column + 2; i < summary_Cell_Start_Column; i++)  //обновляем формулы в блоке
                    worksheet.Cells[headerBlockCell.Start.Row, i].Formula = "=SUM(" + ExcelAddress.GetAddress(headerBlockCell.Start.Row + 1, i) + ":" + ExcelAddress.GetAddress(headerBlockCell.Start.Row + dataTableRows.Length, i) + ")";
            }
        }

        static DataTable getReportProjectResourceIntensityDataTable(
              Helpers.ProjectIDNumberListClass.ProjectIDNumberClass Project
            , SqlConnection connection
            , ref string errorText
            )
        {
            if (errorText == "")
            {
                try
                {
                    using (var cmdIn = new SqlCommand())
                    {
                        cmdIn.Connection = connection;
                        cmdIn.CommandText = "[ITProject].[spGetReportProjectResourceIntensity]";
                        cmdIn.CommandType = CommandType.StoredProcedure;
                        cmdIn.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                        cmdIn.Parameters.AddWithValue("@ProjectID", Project.ProjectID);
                        cmdIn.Parameters.AddWithValue("@FilterOnlyProjectStateCodeNotNull", 0);
                        cmdIn.Parameters.AddWithValue("@FilterOnlyAllowedDepartments", 0);
                        cmdIn.ExecuteNonQuery();

                        var dataSet = new DataSet();
                        new SqlDataAdapter { SelectCommand = cmdIn }.Fill(dataSet);

                        if (dataSet.Tables[0].Rows.Count == 0)
                            throw new Exception("\nРаспределения ресурсов пусты");

                        return dataSet.Tables[0];
                    }
                }
                catch (Exception ex)
                {
                    errorText += "\nОшибка выполнения процедуры построения отчёта, причина: " + ex.Message;
                    return new DataTable();
                }
            }
            else
                return new DataTable();
        }

        static int getRowResourse(ExcelWorksheet worksheet, string role, string departmentName, int rowStart, int rowEnd, int columnRole, int columnDepartment)
        {
            int result = -1;
            for (int i = rowStart; i <= rowEnd; i++)
                if ((worksheet.Cells[i, columnRole].Value ?? "").ToString() == role && (worksheet.Cells[i, columnDepartment].Value ?? "").ToString() == departmentName)
                {
                    result = i;
                    break;
                }
            return result;
        }

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
                catch (Exception exception)
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
                    catch (Exception exception)
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
