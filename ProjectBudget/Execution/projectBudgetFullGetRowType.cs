using System.Data;

namespace ProjectBudget
{
    public static partial class Execution
    {
        static string projectBudgetFullGetRowType(DataTable dataTable, int rowIndex)
        {
            string result = "";
            bool hasFirstColumnValue = false;
            bool hasSecondColumnValue = false;
            bool hasOtherColumnsValues = false;
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                if (i == 0 && (dataTable.Rows[rowIndex][i] ?? "").ToString() != "")
                    hasFirstColumnValue = true;
                else if (i == 1 && (dataTable.Rows[rowIndex][i] ?? "").ToString() != "")
                    hasSecondColumnValue = true;
                else if ((dataTable.Rows[rowIndex][i] ?? "").ToString() != "")
                {
                    hasOtherColumnsValues = true;
                    break;
                }
            }
            if (hasFirstColumnValue && !hasSecondColumnValue && !hasOtherColumnsValues)
                result = "TypeName";
            else if (hasFirstColumnValue && !hasSecondColumnValue && hasOtherColumnsValues)
                result = "Summary";
            return result;
        }
    }
}
