namespace Helpers
{
    public static class SugarDataTable
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

        public static void RemoveRowsAll(System.Data.DataTable dataTable)
        {
            while (dataTable.Rows.Count > 0)
                dataTable.Rows.RemoveAt(0);
        }

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
