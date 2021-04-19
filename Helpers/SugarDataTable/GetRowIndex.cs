namespace Helpers
{
    public static partial class SugarDataTable
    {
        public static int GetRowIndex(System.Data.DataTable dataTable, string columnName, string rowValue)
        {
            int columnIndex = -1;
            for (int i = 0; i < dataTable.Columns.Count; i++)
                if (dataTable.Columns[i].ColumnName == columnName)
                    columnIndex = i;
            return GetRowIndex(dataTable, columnIndex, rowValue);
        }

        public static int GetRowIndex(System.Data.DataTable dataTable, int columnIndex, string rowValue)
        {
            int result = -1;
            for (int i = 0; i < dataTable.Rows.Count; i++)
                if (dataTable.Rows[i][columnIndex].ToString() == rowValue)
                    return i;
            return result;
        }
    }
}
