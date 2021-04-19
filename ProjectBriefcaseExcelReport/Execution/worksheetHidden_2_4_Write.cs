using System;
using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void worksheetHidden_2_4_Write(ExcelPackage excelPackage, SqlConnection connection, string dateStart, string dateEnd, string productionCalendarIDStart, string productionCalendarIDEnd)
        {
            //////////////////Здоровье динамических и водопадных проектов
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_4_Name);
            
            try
            {
                DataTable dataTable = new DataTable();
                using (var cmd = new SqlCommand())
                {
                    cmd.Connection = connection;
                    cmd.CommandText = "[ITProject].[spGetExcelReportProjectScore]";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                    cmd.Parameters.AddWithValue("@StartDateString", dateStart);
                    cmd.Parameters.AddWithValue("@EndDateString", dateEnd);
                    cmd.Parameters.AddWithValue("@ProductionCalendarIDStart", productionCalendarIDStart);
                    cmd.Parameters.AddWithValue("@ProductionCalendarIDEnd", productionCalendarIDEnd);
                    cmd.Parameters.AddWithValue("@IsFinishedOnly", true);
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
                        Settings.Variables.WorksheetHidden_2_4_ProjectCount++;
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetHidden_2_4_Name));
            }
        }
    }
}