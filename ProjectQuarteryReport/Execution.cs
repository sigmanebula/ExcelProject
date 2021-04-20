using System;
using System.IO;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.Data;

namespace ProjectQuarteryReport
{
    public static class Execution
    {
        public static Helpers.ReturnClass GetFromSQLToFile(string projectID, string productionCalendarID)
        {
            Settings.SQLVariables = new SQLVariablesClass();
            Settings.Variables    = new VariablesClass();
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

                Settings.SQLVariables = new SQLVariablesClass();
                Settings.Variables = new VariablesClass();
                connection.Close();

                GC.Collect();

                userMessage = Helpers.Sugar.GetUserMessageAndErrorText(userMessage, errorText, isGetErrorMessage);
                
                return new Helpers.ReturnClass() { FileData = fileData, UserMessage = userMessage };
            }
        }



        static void worksheetMainData_BudgetFull_Write(ExcelWorksheet worksheet, SqlConnection connection)
        {
            DataTable dataTable = new DataTable();
            using (var cmd = new SqlCommand())
            {
                cmd.Connection = connection;
                cmd.CommandText = "[ITProject].[spGetProjectBudgetFullQuartersAll]";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                cmd.Parameters.AddWithValue("@ProjectID", Settings.Variables.ProjectID);
                cmd.ExecuteNonQuery();
                var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                var dataSet = new DataSet();
                dataAdapter.Fill(dataSet);
                dataTable = dataSet.Tables[0];
            }

            DataTable dataTableSorted = Helpers.SugarDataTable.CopyDataTableByColumnList(dataTable, Settings.SQLVariables.Data_BudgetFull_FieldList);

            Helpers.Excel.WriteDataTableToWorkSheet(
                  Settings.SQLVariables.Data_BudgetFull_Column + Settings.SQLVariables.Data_BudgetFull_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.Data_BudgetFull_IsDataTableHasHeaders)
                , dataTableSorted
                , worksheet
                );
        }

        static void writeWorksheetMainData(SqlConnection connection, ExcelWorksheet worksheet, ref string errorText)
        {
            if (errorText == "")
            {
                worksheetMainData_Project_Write(worksheet, connection);

                //Current Waterfall Methodology
                worksheetMainData_KPIWaterfall_Write(
                      worksheet
                    , connection
                    , Settings.Variables.ProductionCalendarID
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Waterfall_Mtd_Data_Column
                    , Settings.SQLVariables.Data_KPI_Waterfall_Mtd_Data_FieldList
                    );

                //Current Waterfall Methodology Ratings
                worksheetMainData_KPIWaterfall_Ratings_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Waterfall_Mtd_Ratings_Column
                    );

                //Current Waterfall Methodology Comment
                worksheetMainData_KPIWaterfall_Comment_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Waterfall_Mtd_CommentField_Column
                    , Settings.SQLVariables.Data_KPI_Waterfall_Mtd_CommentField_FieldList
                    );

                //Current Waterfall GUP
                worksheetMainData_KPIWaterfall_Write(
                      worksheet
                    , connection
                    , Settings.Variables.ProductionCalendarID
                    , Settings.CodeGUP
                    , Settings.SQLVariables.Data_KPI_Waterfall_GUP_Data_Column
                    , Settings.SQLVariables.Data_KPI_Waterfall_GUP_Data_FieldList
                    );

                //Current Waterfall GUP Ratings
                worksheetMainData_KPIWaterfall_Ratings_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeGUP
                    , Settings.SQLVariables.Data_KPI_Waterfall_GUP_Ratings_Column
                    );

                //Current Waterfall GUP Comment
                worksheetMainData_KPIWaterfall_Comment_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeGUP
                    , Settings.SQLVariables.Data_KPI_Waterfall_GUP_CommentField_Column
                    , Settings.SQLVariables.Data_KPI_Waterfall_GUP_CommentField_FieldList
                    );

                //////////////////////////NEXT

                //Next Waterfall Methodology
                worksheetMainData_KPIWaterfall_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Next", "ProductionCalendarID")
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Waterfall_Mtd_Data_Next_Column
                    , Settings.SQLVariables.Data_KPI_Waterfall_Mtd_Data_Next_FieldList
                    );

                //////////////////////////Dynamic

                //Current Dynamic Methodology
                worksheetMainData_KPIDynamic_Write(
                      worksheet
                    , connection
                    , Settings.Variables.ProductionCalendarID
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Dynamic_Mtd_Data_Column
                    , Settings.SQLVariables.Data_KPI_Dynamic_Mtd_Data_FieldList
                    );

                //Current Dynamic Methodology Ratings
                worksheetMainData_KPIDynamic_Ratings_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Dynamic_Mtd_Ratings_Column
                    );

                //Current Dynamic Methodology Comment
                worksheetMainData_KPIDynamic_Comment_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Dynamic_Mtd_CommentField_Column
                    , Settings.SQLVariables.Data_KPI_Dynamic_Mtd_CommentField_FieldList
                    );

                //Current Dynamic GUP
                worksheetMainData_KPIDynamic_Write(
                      worksheet
                    , connection
                    , Settings.Variables.ProductionCalendarID
                    , Settings.CodeGUP
                    , Settings.SQLVariables.Data_KPI_Dynamic_GUP_Data_Column
                    , Settings.SQLVariables.Data_KPI_Dynamic_GUP_Data_FieldList
                    );

                //Current Dynamic GUP Ratings
                worksheetMainData_KPIDynamic_Ratings_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeGUP
                    , Settings.SQLVariables.Data_KPI_Dynamic_GUP_Ratings_Column
                    );

                //Current Dynamic GUP Comment
                worksheetMainData_KPIDynamic_Comment_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeGUP
                    , Settings.SQLVariables.Data_KPI_Dynamic_GUP_CommentField_Column
                    , Settings.SQLVariables.Data_KPI_Dynamic_GUP_CommentField_FieldList
                    );

                //////////////////////////NEXT

                //Next Dynamic Methodology
                worksheetMainData_KPIDynamic_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Next", "ProductionCalendarID")
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Dynamic_Mtd_Data_Next_Column
                    , Settings.SQLVariables.Data_KPI_Dynamic_Mtd_Data_Next_FieldList
                    );

                //Risks
                worksheetMainData_Risks_Write(worksheet, connection);

                //Lessons
                worksheetMainData_Lessons_Write(worksheet, connection);

                //////////////////////////Budget

                worksheetMainData_BudgetGeneral_Write(
                      worksheet
                    , connection
                    , Settings.SQLVariables.Data_BudgetGeneral_Base_Name
                    , Settings.SQLVariables.Data_BudgetGeneral_Base_Column
                    );

                worksheetMainData_BudgetGeneral_Write(
                      worksheet
                    , connection
                    , Settings.SQLVariables.Data_BudgetGeneral_Actual_Name
                    , Settings.SQLVariables.Data_BudgetGeneral_Actual_Column
                    );

                worksheetMainData_BudgetGeneral_Write(
                      worksheet
                    , connection
                    , Settings.SQLVariables.Data_BudgetGeneral_Estimate_PartName
                    , Settings.SQLVariables.Data_BudgetGeneral_Estimate_Column
                    );

                //BudgetFull
                worksheetMainData_BudgetFull_Write(worksheet, connection);

                if (Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsAutoFitColumns))
                    worksheet.Cells.AutoFitColumns();
            }
        }

        static void worksheetMainData_Risks_Write(ExcelWorksheet worksheet, SqlConnection connection)
        {
            DataTable dataTable = Helpers.SugarSQLConnection.ExecuteSQLCommand(
                  connection
                , String.Format(
                      Settings.SQLCommandGetProjectQuarterlyRisks
                    , Settings.Variables.ProjectID
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    )
                , ""
                );

            Helpers.Excel.WriteDataTableToWorkSheet(
                  Settings.SQLVariables.Data_Risks_Column + Settings.SQLVariables.Data_Risks_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsDataTableHasHeaders)
                , dataTable
                , worksheet
                );
        }

        static void worksheetMainData_Project_Write(ExcelWorksheet worksheet, SqlConnection connection)
        {
            DataTable dataTable = new DataTable();
            using (var cmd = new SqlCommand())
            {
                cmd.Connection = connection;
                cmd.CommandText = "[ITProject].[spGetExcelProjectQuarteryReportMainPart]";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                cmd.Parameters.AddWithValue("@ProjectID", Settings.Variables.ProjectID);
                cmd.Parameters.AddWithValue("@ProductionCalendarID", Settings.Variables.ProductionCalendarID);
                cmd.Parameters.AddWithValue("@StuffPrefix", Settings.SQLVariables.StuffPrefix);
                cmd.ExecuteNonQuery();
                var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                var dataSet = new DataSet();
                dataAdapter.Fill(dataSet);
                dataTable = dataSet.Tables[0];
            }

            Helpers.Excel.WriteDataTableToWorkSheet(
                  Settings.SQLVariables.Data_Project_Column + Settings.SQLVariables.MainData_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsDataTableHasHeaders)
                , dataTable
                , worksheet
                );
        }

        static void worksheetMainData_Lessons_Write(ExcelWorksheet worksheet, SqlConnection connection)
        {
            DataTable dataTable = Helpers.SugarSQLConnection.ExecuteSQLCommand(
                  connection
                , String.Format(
                      Settings.SQLCommandGetProjectQuarterlyLessons
                    , Settings.Variables.ProjectID
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    )
                , ""
                );

            Helpers.Excel.WriteDataTableToWorkSheet(
                  Settings.SQLVariables.Data_Lessons_Column + Settings.SQLVariables.Data_Lessons_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsDataTableHasHeaders)
                , dataTable
                , worksheet
                );
        }

        static void worksheetMainData_KPIWaterfall_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string productionCalendarID
            , string methodType
            , string startColumn
            , string columnList
            )
        {
            DataTable dataTable = new DataTable();
            using (var cmd = new SqlCommand())
            {
                cmd.Connection = connection;
                cmd.CommandText = "[ITProject].[spGetExcelProjectQuarteryReportKPIWaterfallPart]";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                cmd.Parameters.AddWithValue("@ProjectID", Settings.Variables.ProjectID);
                cmd.Parameters.AddWithValue("@ProductionCalendarID", productionCalendarID);
                cmd.Parameters.AddWithValue("@MethodType", methodType);
                cmd.ExecuteNonQuery();
                var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                var dataSet = new DataSet();
                dataAdapter.Fill(dataSet);
                dataTable = dataSet.Tables[0];
            }

            DataTable dataTableSorted = Helpers.SugarDataTable.CopyDataTableByColumnList(dataTable, columnList);

            Helpers.Excel.WriteDataTableToWorkSheet(
                  startColumn + Settings.SQLVariables.KPIData_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsDataTableHasHeaders)
                , dataTableSorted
                , worksheet
                );
        }

        static void worksheetMainData_KPIWaterfall_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string productionCalendarID
            , string methodType
            , string startColumn
            )
        {
            worksheetMainData_KPIWaterfall_Write(
                  worksheet
                , connection
                , productionCalendarID
                , methodType
                , startColumn
                , ""
                );
        }

        static void worksheetMainData_KPIWaterfall_Ratings_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string year
            , string quarter
            , string methodType
            , string startColumn
            , string columnList
            )
        {
            DataTable dataTable = new DataTable();
            using (var cmd = new SqlCommand())
            {
                cmd.Connection = connection;
                cmd.CommandText = "[ITProject].[spGetKPIWaterfallRating]";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                cmd.Parameters.AddWithValue("@ProjectID", Settings.Variables.ProjectID);
                cmd.Parameters.AddWithValue("@Year", year);
                cmd.Parameters.AddWithValue("@Quarter", quarter);
                cmd.Parameters.AddWithValue("@MethodType", methodType);
                cmd.ExecuteNonQuery();
                var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                var dataSet = new DataSet();
                dataAdapter.Fill(dataSet);
                dataTable = dataSet.Tables[0];
            }

            DataTable dataTableSorted = Helpers.SugarDataTable.CopyDataTableByColumnList(dataTable, columnList);

            Helpers.Excel.WriteDataTableToWorkSheet(
                  startColumn + Settings.SQLVariables.KPIData_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsDataTableHasHeaders)
                , dataTableSorted
                , worksheet
                );
        }

        static void worksheetMainData_KPIWaterfall_Ratings_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string year
            , string quarter
            , string methodType
            , string startColumn
            )
        {
            worksheetMainData_KPIWaterfall_Ratings_Write(
              worksheet
            , connection
            , year
            , quarter
            , methodType
            , startColumn
            , ""
            );
        }

        static void worksheetMainData_KPIWaterfall_Comment_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string year
            , string quarter
            , string methodType
            , string startColumn
            , string columnList
            )
        {
            DataTable dataTable = new DataTable();
            using (var cmd = new SqlCommand())
            {
                cmd.Connection = connection;
                cmd.CommandText = "[ITProject].[spGetKPIWaterfallComment]";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                cmd.Parameters.AddWithValue("@ProjectID", Settings.Variables.ProjectID);
                cmd.Parameters.AddWithValue("@Year", year);
                cmd.Parameters.AddWithValue("@Quarter", quarter);
                cmd.Parameters.AddWithValue("@MethodType", methodType);
                cmd.ExecuteNonQuery();
                var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                var dataSet = new DataSet();
                dataAdapter.Fill(dataSet);
                dataTable = dataSet.Tables[0];
            }

            DataTable dataTableSorted = Helpers.SugarDataTable.CopyDataTableByColumnList(dataTable, columnList);

            Helpers.Excel.WriteDataTableToWorkSheet(
                  startColumn + Settings.SQLVariables.KPIData_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsDataTableHasHeaders)
                , dataTableSorted
                , worksheet
                );
        }

        static void worksheetMainData_KPIWaterfall_Comment_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string year
            , string quarter
            , string methodType
            , string startColumn
            )
        {
            worksheetMainData_KPIWaterfall_Comment_Write(
                  worksheet
                , connection
                , year
                , quarter
                , methodType
                , startColumn
                , ""
                );
        }


        static void worksheetMainData_KPIDynamic_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string productionCalendarID
            , string methodType
            , string startColumn
            , string columnList
            )
        {
            DataTable dataTable = new DataTable();
            using (var cmd = new SqlCommand())
            {
                cmd.Connection = connection;
                cmd.CommandText = "[ITProject].[spGetExcelProjectQuarteryReportKPIDynamicPart]";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                cmd.Parameters.AddWithValue("@ProjectID", Settings.Variables.ProjectID);
                cmd.Parameters.AddWithValue("@ProductionCalendarID", productionCalendarID);
                cmd.Parameters.AddWithValue("@MethodType", methodType);
                cmd.ExecuteNonQuery();
                var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                var dataSet = new DataSet();
                dataAdapter.Fill(dataSet);
                dataTable = dataSet.Tables[0];
            }

            DataTable dataTableSorted = Helpers.SugarDataTable.CopyDataTableByColumnList(dataTable, columnList);

            Helpers.Excel.WriteDataTableToWorkSheet(
                  startColumn + Settings.SQLVariables.KPIData_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsDataTableHasHeaders)
                , dataTableSorted
                , worksheet
                );
        }

        static void worksheetMainData_KPIDynamic_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string productionCalendarID
            , string methodType
            , string startColumn
            )
        {
            worksheetMainData_KPIDynamic_Write(
                  worksheet
                , connection
                , productionCalendarID
                , methodType
                , startColumn
                , ""
                );
        }

        static void worksheetMainData_KPIDynamic_Ratings_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string year
            , string quarter
            , string methodType
            , string startColumn
            , string columnList
            )
        {
            DataTable dataTable = new DataTable();
            using (var cmd = new SqlCommand())
            {
                cmd.Connection = connection;
                cmd.CommandText = "[ITProject].[spGetKPIDynamicRating]";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                cmd.Parameters.AddWithValue("@ProjectID", Settings.Variables.ProjectID);
                cmd.Parameters.AddWithValue("@Year", year);
                cmd.Parameters.AddWithValue("@Quarter", quarter);
                cmd.Parameters.AddWithValue("@MethodType", methodType);
                cmd.ExecuteNonQuery();
                var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                var dataSet = new DataSet();
                dataAdapter.Fill(dataSet);
                dataTable = dataSet.Tables[0];
            }

            DataTable dataTableSorted = Helpers.SugarDataTable.CopyDataTableByColumnList(dataTable, columnList);

            Helpers.Excel.WriteDataTableToWorkSheet(
                  startColumn + Settings.SQLVariables.KPIData_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsDataTableHasHeaders)
                , dataTableSorted
                , worksheet
                );
        }

        static void worksheetMainData_KPIDynamic_Ratings_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string year
            , string quarter
            , string methodType
            , string startColumn
            )
        {
            worksheetMainData_KPIDynamic_Ratings_Write(
              worksheet
            , connection
            , year
            , quarter
            , methodType
            , startColumn
            , ""
            );
        }

        static void worksheetMainData_KPIDynamic_Comment_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string year
            , string quarter
            , string methodType
            , string startColumn
            , string columnList
            )
        {
            DataTable dataTable = new DataTable();
            using (var cmd = new SqlCommand())
            {
                cmd.Connection = connection;
                cmd.CommandText = "[ITProject].[spGetKPIDynamicComment]";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                cmd.Parameters.AddWithValue("@ProjectID", Settings.Variables.ProjectID);
                cmd.Parameters.AddWithValue("@Year", year);
                cmd.Parameters.AddWithValue("@Quarter", quarter);
                cmd.Parameters.AddWithValue("@MethodType", methodType);
                cmd.ExecuteNonQuery();
                var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                var dataSet = new DataSet();
                dataAdapter.Fill(dataSet);
                dataTable = dataSet.Tables[0];
            }

            DataTable dataTableSorted = Helpers.SugarDataTable.CopyDataTableByColumnList(dataTable, columnList);

            Helpers.Excel.WriteDataTableToWorkSheet(
                  startColumn + Settings.SQLVariables.KPIData_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsDataTableHasHeaders)
                , dataTableSorted
                , worksheet
                );
        }

        static void worksheetMainData_KPIDynamic_Comment_Write(
              ExcelWorksheet worksheet
            , SqlConnection connection
            , string year
            , string quarter
            , string methodType
            , string startColumn
            )
        {
            worksheetMainData_KPIDynamic_Comment_Write(
                  worksheet
                , connection
                , year
                , quarter
                , methodType
                , startColumn
                , ""
                );
        }

        static void worksheetMainData_BudgetGeneral_Write(ExcelWorksheet worksheet, SqlConnection connection, string budgetName, string startColumn)
        {
            DataTable dataTable = Helpers.SugarSQLConnection.ExecuteSQLCommand(
                  connection
                , String.Format(Settings.SQLCommandGetProjectBudgetGeneral, Settings.Variables.ProjectID, budgetName)
                , ""
                );

            DataTable dataTableSorted = Helpers.SugarDataTable.CopyDataTableByColumnList(dataTable, Settings.SQLVariables.Data_BudgetGeneral_FieldList);

            Helpers.Excel.WriteDataTableToWorkSheet(
                  startColumn + Settings.SQLVariables.Data_BudgetGeneral_StartRow
                , Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsDataTableHasHeaders)
                , dataTableSorted
                , worksheet
                );
        }

        static void getProductionCalendarData(SqlConnection connection, ref string errorText)
        {
            if (errorText == "")
                try
                {
                    using (var cmd = new SqlCommand())
                    {
                        cmd.Connection = connection;
                        cmd.CommandText = "[ITProject].[spGetExcelProductionCalendar]";
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                        cmd.Parameters.AddWithValue("@ProductionCalendarID", Settings.Variables.ProductionCalendarID);
                        cmd.ExecuteNonQuery();
                        var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                        var dataSet = new DataSet();
                        dataAdapter.Fill(dataSet);
                        Settings.Variables.ProductionCalendar = dataSet.Tables[0];
                    }
                }
                catch (Exception ex)
                {
                    errorText += "Ошибка: не удалось получить данные по календарю, причина: " + ex.Message;
                }
        }

        static string getFileShortNameNew(SqlConnection connection, ref string errorText)
        {
            if (errorText == "")
            {
                try
                {
                    DataTable dataTable = Helpers.SugarSQLConnection.ExecuteSQLCommand(
                        connection
                        , String.Format(Settings.SQLCommandGetProjectNumberShortName, Settings.Variables.ProjectID)
                        , ""
                        );

                    Settings.Variables.ProjectNumberShortName = dataTable.Rows[0][0].ToString();

                    return Settings.SQLVariables.NewFileNamePrefix
                        + Settings.Variables.ProjectNumberShortName
                        + " за "
                        + Settings.Variables.GetProductionCalendar("Current", "Year")
                        + " г. "
                        + Settings.Variables.GetProductionCalendar("Current", "Quarter")
                        + "кв.."
                        + Settings.FileExtention;
                }
                catch (Exception ex)
                {
                    errorText += "Ошибка: не удалось получить данные по проекту и календарю, причина: " + ex.Message;
                    return "";
                }
            }
            else
                return "";
        }






    }
}
