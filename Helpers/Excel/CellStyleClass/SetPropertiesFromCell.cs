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
            public void SetPropertiesFromCell(ExcelWorksheet worksheet, string rangeAddress)
            {
                SetPropertiesFromCell(worksheet.Cells[rangeAddress]);
            }

            public void SetPropertiesFromCell(ExcelRangeBase cell)
            {
                FillBackgroundColor = GetExcelColor(cell.Style.Fill.BackgroundColor);

                FontSize            = cell.Style.Font.Size;
                FontName            = cell.Style.Font.Name;
                FontBold            = cell.Style.Font.Bold;
                FontColor           = GetExcelColor(cell.Style.Font.Color);

                WrapText            = cell.Style.WrapText;

                BorderTopStyle      = cell.Style.Border.Top.Style;
                BorderBottomStyle   = cell.Style.Border.Bottom.Style;
                BorderLeftStyle     = cell.Style.Border.Left.Style;
                BorderRightStyle    = cell.Style.Border.Right.Style;
                
                BorderTopColor      = GetExcelColor(cell.Style.Border.Top.Color);
                BorderBottomColor   = GetExcelColor(cell.Style.Border.Bottom.Color);
                BorderLeftColor     = GetExcelColor(cell.Style.Border.Left.Color);
                BorderRightColor    = GetExcelColor(cell.Style.Border.Right.Color);

                BorderTopIndexed    = cell.Style.Border.Top.Color.Indexed;
                BorderBottomIndexed = cell.Style.Border.Bottom.Color.Indexed;
                BorderLeftIndexed   = cell.Style.Border.Left.Color.Indexed;
                BorderRightIndexed  = cell.Style.Border.Right.Color.Indexed;

                BorderTopTint       = cell.Style.Border.Top.Color.Tint;
                BorderBottomTint    = cell.Style.Border.Bottom.Color.Tint;
                BorderLeftTint      = cell.Style.Border.Left.Color.Tint;
                BorderRightTint     = cell.Style.Border.Right.Color.Tint;

                VerticalAlignment = cell.Style.VerticalAlignment;
                HorizontalAlignment = cell.Style.HorizontalAlignment;
            }

        }
    }
}
