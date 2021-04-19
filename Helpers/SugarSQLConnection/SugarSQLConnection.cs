namespace Helpers
{
    public static partial class SugarSQLConnection
    {
        const string ConnectionString = "Integrated=True;IsPrimaryLogin=True;Authenticate=True;EncryptedPassword=False;Host=localhost;Port=5555";
        public static int TimeOutSql = 2000;   // Time Out запросов к sql
    }
}
