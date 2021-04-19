namespace ProjectBudget
{
    public static partial class Settings
    {
        public static Helpers.ProjectIDNumberListClass ProjectIDNumber = new Helpers.ProjectIDNumberListClass();

        public static SQLVariablesClass SQLVariables = new SQLVariablesClass();
        
        public static string FileExtention = "xlsx"; //не меняется
        public static string XMLColumnNameTypeNameFull = "ProjectBudgetFullTypeName"; //не меняется
        public static string XMLColumnNameIsSummaryTypeFull = "IsSummaryType"; //не меняется
        public static string XMLColumnNameProjectIDFull = "ProjectID"; //не меняется
        public static string SettingsTypeCodeList = "'project_budget','project_budget_cells'";
        
    }
}
