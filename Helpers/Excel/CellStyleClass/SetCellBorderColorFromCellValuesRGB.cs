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
                    catch(Exception ex)
                    {
                        throw new Exception("Не получилось получить RGB, строка: " + Convert.ToString(cell.Value ?? "") + ", разделитель: " + splitter + ", ошибка: " + ex.Message);
                    }

                BorderTopColor      = newColor;
                BorderBottomColor   = newColor;
                BorderLeftColor     = newColor;
                BorderRightColor    = newColor;
            }
        }
    }
}

