using System;
using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void worksheetHidden_3_1_Write(ExcelPackage excelPackage, SqlConnection connection, string dateStart, string dateEnd, string stateIdList, string projectTypeIdList)
        {
            //////////////////Статистика запросов на изменение портфеля
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_3_1_Name);

            try
            {
                DataTable dataTable = new DataTable();
                using (var cmd = new SqlCommand())
                {
                    cmd.Connection = connection;
                    cmd.CommandText = "[ITProject].[spGetExcelReportPortfolioChangeStatistic]";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                    cmd.Parameters.AddWithValue("@StartDateString", dateStart);
                    cmd.Parameters.AddWithValue("@EndDateString", dateEnd);
                    cmd.Parameters.AddWithValue("@ProjectStateIDList", stateIdList);
                    cmd.Parameters.AddWithValue("@ProjectTypeIDList", projectTypeIdList);
                    cmd.ExecuteNonQuery();
                    var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                    var dataSet = new DataSet();
                    dataAdapter.Fill(dataSet);
                    dataTable = dataSet.Tables[0];
                }

                foreach (DataRow row in dataTable.Rows)
                    row[0] = Convert.ToString(row[0]).Replace(Environment.NewLine, null);

                int projectCountRowIndex = Helpers.SugarDataTable.GetRowIndex(dataTable, 0, "Temp_ProjectCount");
                int waitingCountRowIndex = Helpers.SugarDataTable.GetRowIndex(dataTable, 0, "Temp_WaitingCount");

                //записываем число проектов на изменения
                Settings.Variables.WorksheetHidden_3_1_ProjectCount = Convert.ToInt16(Convert.ToString(dataTable.Rows[projectCountRowIndex][1]));
                Settings.Variables.WorksheetHidden_3_1_WaitCount = Convert.ToInt16(Convert.ToString(dataTable.Rows[waitingCountRowIndex][1]));

                foreach (DataRow row in dataTable.Rows)
                    if (Convert.ToString(row[0]).IndexOf("Temp_") > -1)
                        row.Delete();
                dataTable.AcceptChanges();
                
                Helpers.Excel.WriteDataTableToWorkSheet(dataTable, worksheet);

                //считаем число параметров
                /*
                for (int i = 2; i <= worksheet.Cells.Rows; i++)
                    if (Convert.ToString(worksheet.Cells[i, 1].Value ?? "") != "")
                        Settings.Variables.WorksheetHidden_3_1_CountLine++;
                */
                
                Settings.Variables.WorksheetHidden_3_1_CountLine = dataTable.Rows.Count;
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetHidden_3_1_Name));
            }
        }
    }
}
   
