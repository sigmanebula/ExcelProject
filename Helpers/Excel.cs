using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using OfficeOpenXml;

namespace Helpers
{
    public static class Excel
    {
        public static void CheckExcelPackage(ExcelPackage excelPackage, ref string errorText)
        {
            if (errorText == "")
            {
                try
                {
                    var worksheet = excelPackage.Workbook.Worksheets[1];
                }
                catch (Exception exception)
                {
                    errorText += "Ошибка в Excel: " + exception.Message;
                }
            }
        }

        public static void CheckExcelPackage(ExcelPackage excelPackage)
        {
            string errorText = "";

            CheckExcelPackage(excelPackage, ref errorText);

            if (errorText != "")
                throw new Exception(errorText);
        }



        public static void CheckExistedCellLocation(ExcelWorksheet worksheet, string location)
        {
            if (location != ExcelAddress.GetAddress(worksheet.Cells[location].Start.Row, worksheet.Cells[location].Start.Column))
                throw new Exception("Некорректный адрес ячейки: " + location);
        }


        public static void CopyDataFromAnotherWorkSheet(
              string startCellFrom
            , string endCellFrom
            , string startCellTo
            , ExcelPackage excelPackage
            , string worksheetFromName
            , string worksheetToName
            , bool byLinks
        )
        {
            CopyDataFromAnotherWorkSheet(
                  startCellFrom
                , endCellFrom
                , startCellTo

                , GetExcelWorksheetByName(excelPackage, worksheetFromName)
                , GetExcelWorksheetByName(excelPackage, worksheetToName)

                , byLinks
                );
        }

        public static void CopyDataFromAnotherWorkSheet(
              string startCellFrom
            , string endCellFrom
            , string startCellTo
            , ExcelWorksheet worksheetFrom
            , ExcelWorksheet worksheetTo
            , bool byLinks
            )
        {
            CopyDataFromAnotherWorkSheet(
                  worksheetFrom.Cells[startCellFrom].Start.Row
                , worksheetFrom.Cells[startCellFrom].Start.Column
                , worksheetFrom.Cells[endCellFrom].End.Row
                , worksheetFrom.Cells[endCellFrom].End.Column

                , worksheetTo.Cells[startCellTo].Start.Row
                , worksheetTo.Cells[startCellTo].Start.Column

                , worksheetFrom
                , worksheetTo
                , byLinks
                );
        }

        public static void CopyDataFromAnotherWorkSheet(
              int startRowFrom
            , int startColumnFrom
            , int endRowFrom
            , int endColumnFrom

            , int startRowTo
            , int startColumnTo

            , ExcelWorksheet worksheetFrom
            , ExcelWorksheet worksheetTo
            , bool byLinks
            )
        {
            int differenceRowTo = startRowTo - startRowFrom;
            int differenceColumnTo = startColumnTo - startColumnFrom;

            for (int i = startRowFrom; i <= endRowFrom; i++)
                for (int j = startColumnFrom; j <= endColumnFrom; j++)
                {
                    if (byLinks)
                        worksheetTo.Cells[i + differenceRowTo, j + differenceColumnTo].Formula = "=" + worksheetFrom.Name + "!" + ExcelAddress.GetAddress(i, j);
                    else
                        worksheetTo.Cells[i + differenceRowTo, j + differenceColumnTo].Value = worksheetFrom.Cells[i, j].Value;
                }
        }


        public static void FillWorksheetEmptyValues(ExcelWorksheet worksheet, int rowStart, int columnStart, int rowEnd, int columnEnd)
        {
            for (int i = rowStart; i <= rowEnd; i++)
                for (int j = columnStart; j <= columnEnd; j++)
                    if (worksheet.Cells[i, j].Value == null)
                        worksheet.Cells[i, j].Value = "-";
        }


        public static ExcelRangeBase GetCellByValue(ExcelWorksheet worksheet, string value)
        {
            return GetCellByValue(worksheet, value, 0, 0);
        }

        public static ExcelRangeBase GetCellByValue(ExcelWorksheet worksheet, string value, int rowStart, int rowEnd)
        {
            foreach (var cell in worksheet.Cells)
                if ((cell.Value ?? "").ToString() == value)
                    if (rowStart == 0 || cell.Start.Row >= rowStart)
                        if (rowEnd == 0 || cell.Start.Row <= rowEnd)
                            return cell;
            return null;
        }

        public static int[] GetCellDifference(ExcelWorksheet worksheet, string cellFirst, string cellSecond)
        {
            return new int[] {
                  worksheet.Cells[cellSecond].Start.Row     - worksheet.Cells[cellFirst].Start.Row
                , worksheet.Cells[cellSecond].Start.Column  - worksheet.Cells[cellFirst].Start.Column
            };  //row, col
        }


        public static System.Drawing.Color GetExcelColor(OfficeOpenXml.Style.ExcelColor excelColor)
        {
            return GetExcelColor(excelColor, System.Drawing.Color.Transparent);
        }

        public static System.Drawing.Color GetExcelColor(OfficeOpenXml.Style.ExcelColor excelColor, System.Drawing.Color defaultColor)
        {
            try
            {
                return System.Drawing.ColorTranslator.FromHtml("#" + excelColor.Rgb.ToString());
                //return System.Drawing.ColorTranslator.FromHtml(excelColor.LookupColor());
            }
            catch
            {
                return defaultColor;
            }
        }


        public static ExcelWorksheet GetExcelWorksheetByName(ExcelPackage excelPackage, string worksheetName, ref string errorText)
        {
            ExcelWorksheet worksheet = null;

            if (errorText == "")
                try
                {
                    worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault(w => w.Name == worksheetName);

                    if (worksheet == null)
                        throw new Exception("worksheet is null ");
                }
                catch (Exception exception)
                {
                    errorText += "Вкладка excel не найдена: " + worksheetName + ", " + exception.Message;
                }

            return worksheet;
        }

        public static ExcelWorksheet GetExcelWorksheetByName(ExcelPackage excelPackage, string worksheetName)
        {
            string errorText = "";
            ExcelWorksheet worksheet = GetExcelWorksheetByName(excelPackage, worksheetName, ref errorText);

            if (errorText != "")
                throw new Exception(errorText);

            return worksheet;
        }


        public static string GetLocationNewMoveCell(ExcelWorksheet worksheet, string cell, int[] rowsColumns)
        {
            return GetLocationNewMoveCell(worksheet, cell, rowsColumns[0], rowsColumns[1]);
        }

        public static string GetLocationNewMoveCell(ExcelWorksheet worksheet, string cell, int rows, int columns)
        {
            return ExcelAddress.GetAddress(worksheet.Cells[cell].Start.Row + rows, worksheet.Cells[cell].Start.Column + columns);
        }


        public static int GetWorksheetColumnIndexByName(ExcelWorksheet worksheet, string columnName)
        {
            return worksheet.Cells[columnName + "1"].Start.Column;
        }

        public static string GetWorksheetError(string exceptionMessage, string workSheetName)
        {
            return "\nОшибка в листе " + workSheetName + ", текст ошибки: " + exceptionMessage;
        }

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


        public class CellStyleClass
        {
            public OfficeOpenXml.Style.ExcelFillStyle FillPatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

            public System.Drawing.Color FillBackgroundColor { get; set; }

            public float FontSize { get; set; }
            public string FontName { get; set; }
            public bool FontBold { get; set; }
            public System.Drawing.Color FontColor { get; set; }

            public bool WrapText { get; set; }

            public OfficeOpenXml.Style.ExcelBorderStyle BorderTopStyle { get; set; }
            public OfficeOpenXml.Style.ExcelBorderStyle BorderBottomStyle { get; set; }
            public OfficeOpenXml.Style.ExcelBorderStyle BorderLeftStyle { get; set; }
            public OfficeOpenXml.Style.ExcelBorderStyle BorderRightStyle { get; set; }

            public System.Drawing.Color BorderTopColor { get; set; }
            public System.Drawing.Color BorderBottomColor { get; set; }
            public System.Drawing.Color BorderLeftColor { get; set; }
            public System.Drawing.Color BorderRightColor { get; set; }

            public int BorderTopIndexed { get; set; }
            public int BorderBottomIndexed { get; set; }
            public int BorderLeftIndexed { get; set; }
            public int BorderRightIndexed { get; set; }

            public decimal BorderTopTint { get; set; }
            public decimal BorderBottomTint { get; set; }
            public decimal BorderLeftTint { get; set; }
            public decimal BorderRightTint { get; set; }

            public OfficeOpenXml.Style.ExcelVerticalAlignment VerticalAlignment { get; set; }
            public OfficeOpenXml.Style.ExcelHorizontalAlignment HorizontalAlignment { get; set; }

            public void FillRange(ExcelWorksheet worksheet, int startRow, int startColumn)
            {
                FillRange(worksheet, ExcelAddress.GetAddress(startRow, startColumn));
            }

            public void FillRange(ExcelWorksheet worksheet, int startRow, int startColumn, int endRow, int endColumn)
            {
                FillRange(worksheet, ExcelAddress.GetAddress(startRow, startColumn, endRow, endColumn));
            }

            public void FillRange(ExcelWorksheet worksheet, string rangeAddress)
            {
                if (FillBackgroundColor == System.Drawing.Color.Transparent)
                    FillPatternType = OfficeOpenXml.Style.ExcelFillStyle.None;
                else
                {
                    worksheet.Cells[rangeAddress].Style.Fill.PatternType = FillPatternType;
                    worksheet.Cells[rangeAddress].Style.Fill.BackgroundColor.SetColor(FillBackgroundColor);
                }

                worksheet.Cells[rangeAddress].Style.Font.Size = FontSize;
                worksheet.Cells[rangeAddress].Style.Font.Name = FontName;
                worksheet.Cells[rangeAddress].Style.Font.Bold = FontBold;
                worksheet.Cells[rangeAddress].Style.Font.Color.SetColor(FontColor);

                worksheet.Cells[rangeAddress].Style.Border.Top.Style = BorderTopStyle;
                worksheet.Cells[rangeAddress].Style.Border.Bottom.Style = BorderBottomStyle;
                worksheet.Cells[rangeAddress].Style.Border.Left.Style = BorderLeftStyle;
                worksheet.Cells[rangeAddress].Style.Border.Right.Style = BorderRightStyle;

                if (BorderTopStyle != OfficeOpenXml.Style.ExcelBorderStyle.None)
                {
                    worksheet.Cells[rangeAddress].Style.Border.Top.Color.SetColor(BorderTopColor);
                    if (BorderTopColor == null)
                    {
                        worksheet.Cells[rangeAddress].Style.Border.Top.Color.Indexed = BorderTopIndexed;
                        worksheet.Cells[rangeAddress].Style.Border.Top.Color.Tint = BorderTopTint;
                    }
                }

                if (BorderBottomStyle != OfficeOpenXml.Style.ExcelBorderStyle.None)
                {
                    worksheet.Cells[rangeAddress].Style.Border.Bottom.Color.SetColor(BorderBottomColor);
                    if (BorderBottomColor == null)
                    {
                        worksheet.Cells[rangeAddress].Style.Border.Bottom.Color.Indexed = BorderBottomIndexed;
                        worksheet.Cells[rangeAddress].Style.Border.Bottom.Color.Tint = BorderBottomTint;
                    }
                }

                if (BorderLeftStyle != OfficeOpenXml.Style.ExcelBorderStyle.None)
                {
                    worksheet.Cells[rangeAddress].Style.Border.Left.Color.SetColor(BorderLeftColor);
                    if (BorderLeftColor == null)
                    {
                        worksheet.Cells[rangeAddress].Style.Border.Left.Color.Indexed = BorderLeftIndexed;
                        worksheet.Cells[rangeAddress].Style.Border.Left.Color.Tint = BorderLeftTint;
                    }
                }

                if (BorderRightStyle != OfficeOpenXml.Style.ExcelBorderStyle.None)
                {
                    worksheet.Cells[rangeAddress].Style.Border.Right.Color.SetColor(BorderRightColor);
                    if (BorderRightColor == null)
                    {
                        worksheet.Cells[rangeAddress].Style.Border.Right.Color.Indexed = BorderRightIndexed;
                        worksheet.Cells[rangeAddress].Style.Border.Right.Color.Tint = BorderRightTint;
                    }
                }

                worksheet.Cells[rangeAddress].Style.VerticalAlignment = VerticalAlignment;
                worksheet.Cells[rangeAddress].Style.HorizontalAlignment = HorizontalAlignment;

                worksheet.Cells[rangeAddress].Style.WrapText = WrapText;
            }

            public void SetCellBorderColorFromCellValuesRGB(ExcelWorksheet worksheet, string rangeAddress, string splitter)
            {
                SetCellBorderColorFromCellValuesRGB(worksheet.Cells[rangeAddress], splitter);
            }

            public void SetCellBorderColorFromCellValuesRGB(ExcelRangeBase cell, string splitter)
            {
                System.Drawing.Color newColor = System.Drawing.Color.Transparent;
                if (Convert.ToString(cell.Value ?? "") != "")
                    try
                    {
                        string[] rgb = Convert.ToString(cell.Value).Split(new string[1] { splitter }, StringSplitOptions.RemoveEmptyEntries);
                        if (rgb.Length == 3 || rgb.Length == 4)
                        {
                            newColor = (rgb.Length == 3) ?
                                  System.Drawing.Color.FromArgb(int.Parse(rgb[0]), int.Parse(rgb[1]), int.Parse(rgb[2]))
                                : System.Drawing.Color.FromArgb(int.Parse(rgb[0]), int.Parse(rgb[1]), int.Parse(rgb[2]), int.Parse(rgb[3]));
                        }
                        else
                            throw new Exception();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Не получилось получить RGB, строка: " + Convert.ToString(cell.Value ?? "") + ", разделитель: " + splitter + ", ошибка: " + ex.Message);
                    }

                BorderTopColor = newColor;
                BorderBottomColor = newColor;
                BorderLeftColor = newColor;
                BorderRightColor = newColor;
            }

            public void SetPropertiesFromCell(ExcelWorksheet worksheet, string rangeAddress)
            {
                SetPropertiesFromCell(worksheet.Cells[rangeAddress]);
            }

            public void SetPropertiesFromCell(ExcelRangeBase cell)
            {
                FillBackgroundColor = GetExcelColor(cell.Style.Fill.BackgroundColor);

                FontSize = cell.Style.Font.Size;
                FontName = cell.Style.Font.Name;
                FontBold = cell.Style.Font.Bold;
                FontColor = GetExcelColor(cell.Style.Font.Color);

                WrapText = cell.Style.WrapText;

                BorderTopStyle = cell.Style.Border.Top.Style;
                BorderBottomStyle = cell.Style.Border.Bottom.Style;
                BorderLeftStyle = cell.Style.Border.Left.Style;
                BorderRightStyle = cell.Style.Border.Right.Style;

                BorderTopColor = GetExcelColor(cell.Style.Border.Top.Color);
                BorderBottomColor = GetExcelColor(cell.Style.Border.Bottom.Color);
                BorderLeftColor = GetExcelColor(cell.Style.Border.Left.Color);
                BorderRightColor = GetExcelColor(cell.Style.Border.Right.Color);

                BorderTopIndexed = cell.Style.Border.Top.Color.Indexed;
                BorderBottomIndexed = cell.Style.Border.Bottom.Color.Indexed;
                BorderLeftIndexed = cell.Style.Border.Left.Color.Indexed;
                BorderRightIndexed = cell.Style.Border.Right.Color.Indexed;

                BorderTopTint = cell.Style.Border.Top.Color.Tint;
                BorderBottomTint = cell.Style.Border.Bottom.Color.Tint;
                BorderLeftTint = cell.Style.Border.Left.Color.Tint;
                BorderRightTint = cell.Style.Border.Right.Color.Tint;

                VerticalAlignment = cell.Style.VerticalAlignment;
                HorizontalAlignment = cell.Style.HorizontalAlignment;
            }

        }

    }
}
