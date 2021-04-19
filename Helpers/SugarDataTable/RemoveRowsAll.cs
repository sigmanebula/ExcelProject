namespace Helpers
{
    public static partial class SugarDataTable
    {
        public static void RemoveRowsAll(System.Data.DataTable dataTable)
        {
            while (dataTable.Rows.Count > 0)
                dataTable.Rows.RemoveAt(0);
        }
    }
}
