using System;
using System.Data;

namespace ProjectFinModel
{
    public static partial class Execution
    {
        static DataTable deleteJunkDataTableReport(DataTable dataTableSource, ref string errorText)
        {
            DataTable dataTable = new DataTable();

            if (errorText == "")
                try
                {
                    for (int i = 0; i > -1 && i < dataTableSource.Rows.Count; i++)
                        if (
                            (bool)(dataTableSource.Rows[i]["IsOldAllocation"] ?? false) == true
                            || dataTableSource.Rows[i]["ResourceAllocation_Allocated"] == null
                            || Convert.ToString(dataTableSource.Rows[i]["ResourceAllocation_Allocated"] ?? "") == "0"
                            || Convert.ToString(dataTableSource.Rows[i]["ResourceAllocation_TypeCode"] ?? "") == "empty"
                            || Convert.ToString(dataTableSource.Rows[i]["ResourceAllocation_TypeCode"] ?? "") == ""
                            )
                        {
                            dataTableSource.Rows[i].Delete();
                        }

                    dataTableSource.AcceptChanges();

                    dataTable = Helpers.SugarDataTable.CopyDataTableByColumnList(dataTableSource, Settings.SQLVariables.ReportProjectResourceIntensityColumnList);
                }
                catch (Exception ex)
                {
                    errorText += "\nОшибка очистки данных таблицы отчёта, причина: " + ex.Message;
                }

            return dataTable;
        }
    }
}
