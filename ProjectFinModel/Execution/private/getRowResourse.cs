using OfficeOpenXml;

namespace ProjectFinModel
{
    public static partial class Execution
    {
        static int getRowResourse(ExcelWorksheet worksheet, string role, string departmentName, int rowStart, int rowEnd, int columnRole, int columnDepartment)
        {
            int result = -1;
            for (int i = rowStart; i <= rowEnd; i++)
                if ((worksheet.Cells[i, columnRole].Value ?? "").ToString() == role && (worksheet.Cells[i, columnDepartment].Value ?? "").ToString() == departmentName)
                {
                    result = i;
                    break;
                }
            return result;
        }
    }
}
