namespace Helpers
{
    public static partial class SugarSQLConnection
    {
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
    }
}