using OfficeOpenXml;
using System.Data;

namespace ProjectFinModel
{
    public static partial class Execution
    {
        static void fillDepartmentBlock(
            ExcelWorksheet worksheet
            , string headerCellValue
            , DataRow[] dataTableRows
            , ExcelRangeBase summaryRow_Cell
            , int summary_Cell_Start_Column
            , int role_Cell_Start_Column
            , int department_Cell_Start_Column
            , ref int lastNonEmptyRowIndex
            , ref ExcelRangeBase headerBlockCell
            )
        {
            if (dataTableRows.Length > 0)
            {
                worksheet.InsertRow(lastNonEmptyRowIndex + 1, 1);

                Helpers.Excel.CellStyleClass cellStyle = new Helpers.Excel.CellStyleClass();
                cellStyle.SetPropertiesFromCell(summaryRow_Cell);
                cellStyle.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                cellStyle.BorderLeftStyle = OfficeOpenXml.Style.ExcelBorderStyle.None;
                cellStyle.BorderRightStyle = OfficeOpenXml.Style.ExcelBorderStyle.None;
                cellStyle.FillRange(worksheet, ExcelAddress.GetAddress(lastNonEmptyRowIndex + 1, summaryRow_Cell.Start.Column, lastNonEmptyRowIndex + 1, summary_Cell_Start_Column));

                worksheet.Row(lastNonEmptyRowIndex + 1).Height = worksheet.DefaultRowHeight;
                worksheet.Cells[lastNonEmptyRowIndex + 1, summaryRow_Cell.Start.Column].Value = headerCellValue;
                headerBlockCell = worksheet.Cells[lastNonEmptyRowIndex + 1, summaryRow_Cell.Start.Column];
                lastNonEmptyRowIndex = headerBlockCell.Start.Row;
                
                worksheet.InsertRow(lastNonEmptyRowIndex + 1, dataTableRows.Length);
                
                foreach (DataRow row in dataTableRows)  //заполняем заголовки обычных строк
                {
                    lastNonEmptyRowIndex++;
                    worksheet.Cells[lastNonEmptyRowIndex, role_Cell_Start_Column].Value = row["Role_System_Specialization"].ToString();
                    worksheet.Cells[lastNonEmptyRowIndex, department_Cell_Start_Column].Value = row["Department_FullName"].ToString();
                }
                
                try
                {
                    worksheet.Cells[ExcelAddress.GetAddress(headerBlockCell.Start.Row, summaryRow_Cell.Start.Column, headerBlockCell.Start.Row, summaryRow_Cell.Start.Column + 1)].Merge = true;
                }
                catch { }
                
                cellStyle.FillBackgroundColor = System.Drawing.Color.Transparent;
                cellStyle.FontColor = System.Drawing.Color.Black;
                cellStyle.FontBold = false;
                cellStyle.BorderLeftStyle = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                cellStyle.BorderRightStyle = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                cellStyle.BorderTopColor = System.Drawing.Color.LightGray;
                cellStyle.BorderBottomColor = System.Drawing.Color.LightGray;
                cellStyle.BorderLeftColor = System.Drawing.Color.LightGray;
                cellStyle.BorderRightColor = System.Drawing.Color.LightGray;

                cellStyle.FillRange(worksheet, ExcelAddress.GetAddress(headerBlockCell.Start.Row + 1, summaryRow_Cell.Start.Column, lastNonEmptyRowIndex, summaryRow_Cell.Start.Column + 1));   //роль и департамент
                cellStyle.FillRange(worksheet, ExcelAddress.GetAddress(headerBlockCell.Start.Row + 1, summaryRow_Cell.Start.Column + 2, lastNonEmptyRowIndex, summary_Cell_Start_Column));  //аллокации

                for (int i = headerBlockCell.Start.Column + 2; i < summary_Cell_Start_Column; i++)  //обновляем формулы в блоке
                    worksheet.Cells[headerBlockCell.Start.Row, i].Formula = "=SUM(" + ExcelAddress.GetAddress(headerBlockCell.Start.Row + 1, i) + ":" + ExcelAddress.GetAddress(headerBlockCell.Start.Row + dataTableRows.Length, i) + ")";
            }
        }
    }
}
