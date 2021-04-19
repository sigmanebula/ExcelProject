namespace Helpers
{
    public static partial class SugarDataTable
    {
        public static System.Data.DataTable CopyDataTableByColumnList(System.Data.DataTable dataTableSource, string columnList)
        {
            if (columnList != "")
            {
                System.Data.DataTable dataTableNew = new System.Data.DataTable();
                Helpers.SugarDataTable.AddColumn(dataTableNew, columnList, ",");
                Helpers.SugarDataTable.CopyColumnNamesIdentity(dataTableSource, dataTableNew, false);

                return dataTableNew;
            }
            else
                return dataTableSource;
        }
    }
}