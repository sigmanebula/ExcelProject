using System;
using OfficeOpenXml;
using System.Data;

namespace Helpers
{
    public static partial class Excel
    {
        public static void WriteDataTableToWorkSheet(DataTable dataTable, ExcelWorksheet worksheet)
        {
            WriteDataTableToWorkSheet(1, 1, true, dataTable, worksheet);
        }

        public static void WriteDataTableToWorkSheet(string startLocation, bool hasHeaders, DataTable dataTable, ExcelWorksheet worksheet)
        {
            WriteDataTableToWorkSheet(worksheet.Cells[startLocation].Start.Row, worksheet.Cells[startLocation].Start.Column, hasHeaders, dataTable, worksheet);
        }

        public static void WriteDataTableToWorkSheet(int rowStart, int columnStart, bool hasHeaders, DataTable dataTable, ExcelWorksheet worksheet)
        {
            int headersRowCount = 1;

            if (hasHeaders)
                for (int j = 0; j < dataTable.Columns.Count; j++)
                    worksheet.Cells[rowStart, j + columnStart].Value = Convert.ToString(dataTable.Columns[j].ColumnName ?? ""); //пишем заголовки
            else
                headersRowCount = 0;
            
            for (int i = 0; i < dataTable.Rows.Count; i++)
                for (int j = 0; j < dataTable.Columns.Count; j++)
                    worksheet.Cells[i + rowStart + headersRowCount, j + columnStart].Value = dataTable.Rows[i][j];
        }

    }
}
