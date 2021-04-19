namespace Helpers
{
    public static partial class SugarDataTable
    {
        public static void AddColumn(System.Data.DataTable dataTable, string columnNameList)
        {
            AddColumn(dataTable, columnNameList, ",");
        }

        public static void AddColumn(System.Data.DataTable dataTable, string columnNameList, string columnListDelimeter)
        {
            AddColumn(dataTable, columnNameList.Split(new string[] { columnListDelimeter }, System.StringSplitOptions.RemoveEmptyEntries));
        }

        public static void AddColumn(System.Data.DataTable dataTable, string[] columnNameArray)
        {
            for (int i = 0; i < columnNameArray.Length; i++)
                dataTable.Columns.Add(columnNameArray[i]);
        }
        public static void AddColumn(System.Data.DataTable dataTable)
        {
            dataTable.Columns.Add(dataTable.Columns.Count.ToString());
        }

        public static void AddColumn(System.Data.DataTable dataTable, int columnCount)
        {
            for (int i = 0; i < columnCount; i++)
                dataTable.Columns.Add("column" + (i + 1).ToString());
        }
    }
}
