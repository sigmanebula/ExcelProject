namespace Helpers
{
    public static partial class SugarSQLConnection
    {
        public static System.Data.DataTable ExecuteSQLCommand(System.Data.SqlClient.SqlConnection connection, string commandSQL, string errorPrefix)
        {
            try
            {
                var dataSet = new System.Data.DataSet();

                using (var cmdIn = new System.Data.SqlClient.SqlCommand())
                {
                    cmdIn.Connection = connection;
                    cmdIn.CommandText = commandSQL;
                    cmdIn.CommandType = System.Data.CommandType.Text;
                    cmdIn.CommandTimeout = Helpers.SugarSQLConnection.TimeOutSql;
                    cmdIn.ExecuteNonQuery();
                    new System.Data.SqlClient.SqlDataAdapter { SelectCommand = cmdIn }.Fill(dataSet);
                }
                
                return dataSet.Tables[0];
            }
            catch (System.Exception exception)
            {
                throw new System.Exception(errorPrefix + exception.Message);
            }
        }
    }
}