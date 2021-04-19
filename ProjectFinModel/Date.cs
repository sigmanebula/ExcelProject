using OfficeOpenXml;
using System.Collections.Generic;

namespace ProjectFinModel
{
    public class Date
    {
        public Dictionary<string, YearQuarter> Year = new Dictionary<string, YearQuarter>();
        
        public void Fill(ExcelWorksheet worksheet, ExcelRangeBase departmentCell, ExcelRangeBase summaryCell)
        {
            Year = new Dictionary<string, YearQuarter>();
            int rowYear = departmentCell.Start.Row;
            int columnStart = departmentCell.Start.Column;
            int columnEnd = summaryCell.Start.Column;
            YearQuarter yearQuarter = null;

            for (int i = columnStart + 1; i < columnEnd; i++)
            {
                var cell = worksheet.Cells[ExcelAddress.GetAddress(rowYear, i)];
                if (cell.Value != null)
                {
                    if (yearQuarter != null && yearQuarter.Quarter.Count > 0)
                        Year.Add(yearQuarter.Cell.Value.ToString(), yearQuarter);
                    yearQuarter = new YearQuarter();
                    yearQuarter.Cell = cell;
                }

                cell = worksheet.Cells[ExcelAddress.GetAddress(rowYear + 1, i)];
                if (cell.Value != null)
                    yearQuarter.Quarter.Add(cell.Value.ToString(), cell);
            }
            if (yearQuarter != null && yearQuarter.Quarter.Count > 0)
                Year.Add(yearQuarter.Cell.Value.ToString(), yearQuarter);

            //debug();
        }
        /*
        void debug()
        {
            string result = "";
            foreach (var keyY in Year)
                foreach (var keyQ in Year[keyY.Key].Quarter)
                    result += keyY.Key + " | " + Year[keyY.Key].Cell.Value + " | " + keyQ.Key + " | " + keyQ.Value + "\n";
            throw new Exception(result);
        }
        */
    }
}
