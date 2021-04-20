namespace Helpers
{
    public static class SugarSQLConnection
    {
        const string ConnectionString = "Integrated=True;IsPrimaryLogin=True;Authenticate=True;EncryptedPassword=False;Host=localhost;Port=5555";
        public static int TimeOutSql = 2000;   // Time Out запросов к sql

        public static void OpenInUsing(System.Data.SqlClient.SqlConnection connection)
        {
                try
                {
                    connection.Open();
                }
                catch (System.Exception exception)
                {
                    exception.Message +=
                        "\nОшибка при открытии подключения к базе данных "
                        + connection.ConnectionString
                        + ", причина: "
                        + exception.Message
                        ;
                }                
        }


        public static string GetSQLConnectionStringFromServiceInstance()
        {
            try
            {
                string Server = System.String.Empty;
                string Database = System.String.Empty;
                SourceCode.SmartObjects.Services.Management.ServiceManagementServer serviceManagementServer = new SourceCode.SmartObjects.Services.Management.ServiceManagementServer();
                serviceManagementServer.CreateConnection();
                serviceManagementServer.Connection.Open(ConnectionString);
                System.Xml.XmlDocument xmlDocument = new System.Xml.XmlDocument();
                xmlDocument.LoadXml(serviceManagementServer.GetServiceInstance(new System.Guid("748988a7-ac80-4a45-83c3-ed911cf9d096"))); // Service Instance "K2_MDM_REQUEST on msk-sqldb01"
                //var serviceconfig = xm.SelectNodes("serviceconfig");
                foreach (System.Xml.XmlNode node1 in xmlDocument.GetElementsByTagName("settings"))
                    foreach (System.Xml.XmlNode node2 in node1)
                    {
                        var key = node2.Attributes["name"].Value;
                        if (key.Equals("Server"))
                            Server = node2.InnerText;
                        if (key.Equals("Database"))
                            Database = node2.InnerText;
                        if (key.Equals("Command Timeout"))
                            int.TryParse(node2.InnerText, out TimeOutSql);
                    }
                return $"Data Source={Server};Initial Catalog={Database};Integrated Security=True"; //sqlConnectionString
            }
            catch (System.Exception exception)
            {
                throw new System.Exception("Не удалось получить строку подключения. Причина: " + exception.Message);
            }
        }

        public static System.Data.SqlClient.SqlConnection GetSQLConnection()
        {
            return new System.Data.SqlClient.SqlConnection(Helpers.SugarSQLConnection.GetSQLConnectionStringFromServiceInstance());
        }

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
