namespace Helpers
{
    public static partial class SugarDataTable
    {
        public static void CopyColumnNamesIdentity(
              System.Data.DataTable dataTableSource
            , System.Data.DataTable dataTableTarget
            , bool isExceptionIfNotExist
            )
        {
            try
            {
                RemoveRowsAll(dataTableTarget);

                for (int i = 0; i < dataTableSource.Rows.Count; i++)
                {
                    dataTableTarget.Rows.Add();

                    for (int j = 0; j < dataTableSource.Columns.Count; j++)
                    {
                        if (dataTableTarget.Columns.Contains(dataTableSource.Columns[j].ColumnName))
                            dataTableTarget.Rows[i][dataTableSource.Columns[j].ColumnName] = dataTableSource.Rows[i][dataTableSource.Columns[j].ColumnName];
                        else if (isExceptionIfNotExist)
                            throw new System.Exception("колонка не существует - " + dataTableSource.Columns[j].ColumnName);
                    }
                }
            }
            catch (System.Exception exception)
            {
                throw new System.Exception("Не удалось скопировать DataTable: " + exception.Message);
            }
        }
    }
}