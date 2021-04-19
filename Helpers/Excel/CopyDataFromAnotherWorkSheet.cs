using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Data;
//using System.Windows.Forms;

namespace Helpers
{
    public static partial class Excel
    {
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
            int differenceRowTo     = startRowTo    - startRowFrom;
            int differenceColumnTo  = startColumnTo - startColumnFrom;

            for (int i = startRowFrom; i <= endRowFrom; i++)
                for (int j = startColumnFrom; j <= endColumnFrom; j++)
                {
                    if (byLinks)
                        worksheetTo.Cells[i + differenceRowTo, j + differenceColumnTo].Formula = "=" + worksheetFrom.Name + "!" + ExcelAddress.GetAddress(i, j);
                    else
                        worksheetTo.Cells[i + differenceRowTo, j + differenceColumnTo].Value = worksheetFrom.Cells[i, j].Value;
                }
        }
    }
}
