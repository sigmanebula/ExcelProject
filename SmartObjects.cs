namespace ExcelProject
{
    public static class SmartObjects
    {
        public static void ReadExcelProjectBudget(int projectID, int projectNumber)
        {
            ProjectBudget.Execution.ReadFromFileToSQLTable(projectID.ToString());
        }

        public static void ReadExcelProjectBudgetAll()
        {
            ProjectBudget.Execution.ReadFromFileToSQLTable("");
        }

        public static void ReadExcelProjectBudgetList(string projectIDList)   //"11, 12, 13, 115"
        {
            ProjectBudget.Execution.ReadFromFileToSQLTable(projectIDList);
        }
        
        public static Helpers.ReturnClass WriteExcelProjectFinModelSaveLoad(string projectID, string fileData)  //11, XML
        {
            return ProjectFinModel.Execution.WriteFromSQLToFileSingle(projectID, fileData);
        }

        public static Helpers.ReturnClass GetProjectBriefcaseExcelReport(string dateStart, string dateEnd, string projectTypeCode, string stateCode)
        {
            return ProjectBriefcaseExcelReport.Execution.GetFromSQLToFile(dateStart, dateEnd, projectTypeCode, stateCode);
        }
        
        public static Helpers.ReturnClass GetProjectQuarteryReport(string projectID, string productionCalendarID)
        {
            return ProjectQuarteryReport.Execution.GetFromSQLToFile(projectID, productionCalendarID);
        }
    }
}