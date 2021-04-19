using System;
using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;

namespace ProjectBriefcaseExcelReport
{
    public static partial class Execution
    {
        static void worksheetHidden_2_1_Write(ExcelPackage excelPackage, SqlConnection connection, string dateEnd, string stateIdList, string projectTypeIdList)
        {
            //////////////////Структура портфеля технологических задач
            var worksheet = Helpers.Excel.GetExcelWorksheetByName(excelPackage, Settings.SQLVariables.WorksheetHidden_2_1_Name);
            
            try
            {
                DataTable dataTable = new DataTable();
                using (var cmd = new SqlCommand())
                {
                    cmd.Connection = connection;
                    cmd.CommandText = "[ITProject].[spGetExcelReportBriefcaseStructure]";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                    cmd.Parameters.AddWithValue("@EndDateString", dateEnd);
                    cmd.Parameters.AddWithValue("@ProjectStateIDList", stateIdList);
                    cmd.Parameters.AddWithValue("@ProjectTypeIDList", projectTypeIdList);
                    cmd.ExecuteNonQuery();
                    var dataAdapter = new SqlDataAdapter { SelectCommand = cmd };
                    var dataSet = new DataSet();
                    dataAdapter.Fill(dataSet);
                    dataTable = dataSet.Tables[0];
                }

                Helpers.Excel.WriteDataTableToWorkSheet(dataTable, worksheet);

                Settings.Variables.WorksheetHidden_2_1_DataStartCell = ExcelAddress.GetAddress(2, 1);
                Settings.Variables.WorksheetHidden_2_1_DataEndCell = ExcelAddress.GetAddress(dataTable.Rows.Count + 1, dataTable.Columns.Count);
            }
            catch (Exception ex)
            {
                throw new Exception(Helpers.Excel.GetWorksheetError(ex.Message, Settings.SQLVariables.WorksheetHidden_2_1_Name));
            }
        }
    }
}