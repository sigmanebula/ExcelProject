using System.Linq;
using OfficeOpenXml;
using System.IO;
using System;
using System.Data;
using System.Collections.Generic;
//using System.Windows.Forms;

namespace Helpers
{
    public static partial class Excel
    {
        public partial class CellStyleClass
        {
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

                worksheet.Cells[rangeAddress].Style.VerticalAlignment   = VerticalAlignment;
                worksheet.Cells[rangeAddress].Style.HorizontalAlignment = HorizontalAlignment;

                worksheet.Cells[rangeAddress].Style.WrapText            = WrapText;
            }

        }
    }
}
