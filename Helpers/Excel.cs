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
    public static void CheckExcelPackage(ExcelPackage excelPackage)
    {
      try
      {
        var worksheet = excelPackage.Workbook.Worksheets[1];
      }
      catch (Exception exception)
      {
        exception.Message = "Ошибка в Excel: " + exception.Message;
      }

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


    public static ExcelWorksheet GetExcelWorksheetByName(ExcelPackage excelPackage, string worksheetName)
    {
      ExcelWorksheet worksheet = null;


        try
        {
          worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault(w => w.Name == worksheetName);

          if (worksheet == null)
            throw new Exception("worksheet is null ");
        }
        catch (Exception exception)
        {
          exception.Message = "Вкладка excel не найдена: " + worksheetName + ", " + exception.Message;
        }

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

  }
}
