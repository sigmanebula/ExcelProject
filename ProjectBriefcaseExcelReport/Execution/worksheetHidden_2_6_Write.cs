using System;
using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void worksheetHidden_2_6_Write(ExcelPackage excelPackage, SqlConnection connection, string dateStart, string dateEnd, string stateIdList, string projectTypeIdList)
        {
            //////////////////Включены в портфель
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_6_Name);
            
            try
            {
                DataTable dataTable = new DataTable();
                using (var cmd = new SqlCommand())
                {
                    cmd.Connection = connection;
                    cmd.CommandText = "[ITProject].[spGetExcelReportBriefcaseDynamicPortfolio]";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                    cmd.Parameters.AddWithValue("@StartDateString", dateStart);
                    cmd.Parameters.AddWithValue("@EndDateString", dateEnd);
                    cmd.Parameters.AddWithValue("@ProjectStateIDList", stateIdList);
                    cmd.Parameters.AddWithValue("@ProjectTypeIDList", projectTypeIdList);
                    cmd.Parameters.AddWithValue("@Mode", "ChangedDate_portfolio");

                    cmd.ExecuteNonQuery();
                    var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                    var dataSet = new DataSet();
                    dataAdapter.Fill(dataSet);
                    dataTable = dataSet.Tables[0];
                }

                Helpers.Excel.WriteDataTableToWorkSheet(dataTable, worksheet);

                //считаем число проектов, пригодится
                for (int i = 2; i <= worksheet.Cells.Rows; i++)
                    if (Convert.ToString(worksheet.Cells[i, 1].Value ?? "") != "")
                        Settings.Variables.WorksheetHidden_2_6_ProjectCount++;
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetHidden_2_6_Name));
            }
        }
    }
}

