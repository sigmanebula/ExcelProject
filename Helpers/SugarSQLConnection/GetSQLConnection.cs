namespace Helpers
{
    public static partial class SugarSQLConnection
    {
        public static System.Data.SqlClient.SqlConnection GetSQLConnection()
        {
            return new System.Data.SqlClient.SqlConnection(Helpers.SugarSQLConnection.GetSQLConnectionStringFromServiceInstance());
        }
    }
}