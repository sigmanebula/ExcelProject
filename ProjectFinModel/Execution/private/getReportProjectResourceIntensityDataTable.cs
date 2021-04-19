using System;
using System.Data;
using System.Data.SqlClient;

namespace ProjectFinModel
{
    public static partial class Execution
    {
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

    }
}
