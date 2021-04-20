using System;

namespace Helpers
{
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
