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
        }
    }
}
