using System;
using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void worksheetHidden_4_0_Write(ExcelPackage excelPackage, SqlConnection connection, string dateStart, string dateEnd, int startPeriodNumber, int endPeriodNumber, string stateIdList, string projectTypeIdList)
        {
            //////////////////Статистика запросов на изменение портфеля
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_4_0_Name);

            try
            {
                DataTable dataTable = new DataTable();
                using (var cmd = new SqlCommand())
                {
                    cmd.Connection = connection;
                    cmd.CommandText = "[ITProject].[spGetExcelReportProjectIntegralRatings]";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                    cmd.Parameters.AddWithValue("@StartDateString",    dateStart);
                    cmd.Parameters.AddWithValue("@EndDateString",      dateEnd);
                    cmd.Parameters.AddWithValue("@StartPeriodNumber",  startPeriodNumber);
                    cmd.Parameters.AddWithValue("@EndPeriodNumber",    endPeriodNumber);
                    cmd.Parameters.AddWithValue("@ProjectStateIDList", stateIdList);
                    cmd.Parameters.AddWithValue("@ProjectTypeIDList",  projectTypeIdList);
                    cmd.ExecuteNonQuery();
                    var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                    var dataSet = new DataSet();
                    dataAdapter.Fill(dataSet);
                    dataTable = dataSet.Tables[0];
                }

                Helpers.Excel.WriteDataTableToWorkSheet(dataTable, worksheet);
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetHidden_4_0_Name));
            }
        }
    }
}