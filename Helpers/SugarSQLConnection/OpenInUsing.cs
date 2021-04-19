namespace Helpers
{
    public static partial class SugarSQLConnection
    {
        public static void OpenInUsing(System.Data.SqlClient.SqlConnection connection, ref string errorText)
        {
            if (errorText == "")
                try
                {
                    connection.Open();
                }
                catch (System.Exception exception)
                {
                    errorText +=
                        "\nОшибка при открытии подключения к базе данных "
                        + connection.ConnectionString
                        + ", причина: "
                        + exception.Message
                        ;
                }
        }

        public static void OpenInUsing(System.Data.SqlClient.SqlConnection connection)
        {
            string errorText = "";

            OpenInUsing(connection, ref errorText);

            if (errorText != "")
                throw new System.Exception(errorText);
        }
    }
}