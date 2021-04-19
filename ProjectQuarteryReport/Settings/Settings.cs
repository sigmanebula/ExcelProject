namespace ProjectQuarteryReport
{
    public static partial class Settings
    {
        public static SQLVariablesClass SQLVariables = new SQLVariablesClass();
        public static ProjectQuarteryReport.Settings.VariablesClass Variables = new ProjectQuarteryReport.Settings.VariablesClass();

        public static string FileExtention = "xlsx";

        public static string CodeMethodology = "Methodology";
        public static string CodeGUP = "GUP";
        
        public static string ProductionCalendarCurrent = "Current";
        public static string ProductionCalendarPrevious = "Previous";
        public static string ProductionCalendarNext = "Next";
        
        public static string SettingsTypeCodeList = "'ProjectQuarteryReport'";

        public static string SQLCommandGetProductionCalendarData = @"
            SELECT * FROM [ITProject].[ProductionCalendar] WITH(NOLOCK) WHERE [ProductionCalendarID] = {0}
        ";

        public static string SQLCommandGetProjectNumberShortName = @"
            SELECT TOP 1 [ProjectNumberShortName]
                = CAST(ISNULL([Number], 0) AS NVARCHAR(10))
                + ' '
                + ISNULL(
                     CASE
                        WHEN LEN(RTRIM(LTRIM([ShortName]))) IN (0, 1)
                            THEN NULL
                            ELSE RTRIM(LTRIM([ShortName])) END
                    ,CAST([Name] AS NVARCHAR(150))
                )
           FROM [ITProject].[Project] WITH(NOLOCK) WHERE [ProjectID] = {0}
        ";

        public static string SQLCommandGetProjectQuarterlyRisks = @"
            SELECT
                 [Description]     = [vwProjectQuarterlyRisks].[Description]     --поле Риск
                ,[PropabilityName] = [ProjectQuarterlyRisksPropabilities].[Name] --поле Вероятность наступления
                ,[MeasureName]     = [ProjectQuarterlyRisksMeasures].[Name]      --поле Мера реагирования
                ,[Comment]         = [vwProjectQuarterlyRisks].[Comment]         --поле Комментарий
            FROM        [ITProject].[vwProjectQuarterlyRisks]            AS [vwProjectQuarterlyRisks]
            LEFT JOIN   [ITProject].[ProjectQuarterlyRisksMeasures]      AS [ProjectQuarterlyRisksMeasures]      WITH(NOLOCK) ON
                    [ProjectQuarterlyRisksMeasures].[ID]      = [vwProjectQuarterlyRisks].[MeasureID]
            LEFT JOIN   [ITProject].[ProjectQuarterlyRisksPropabilities] AS [ProjectQuarterlyRisksPropabilities] WITH(NOLOCK) ON
                    [ProjectQuarterlyRisksPropabilities].[ID] = [vwProjectQuarterlyRisks].[PropabilityID]
            WHERE   [vwProjectQuarterlyRisks].[ProjectID]     IN ({0})
                AND [vwProjectQuarterlyRisks].[Year]          IN ({1})
                AND [vwProjectQuarterlyRisks].[Quarter]       IN ({2})
            ORDER BY [vwProjectQuarterlyRisks].[RowIndex] ASC
        ";

        public static string SQLCommandGetProjectQuarterlyLessons = @"
            SELECT
                 [Description]    = [vwProjectQuarterlyLessons].[Description]    --поле Описание ситуации и последствия
                ,[Cause]          = [vwProjectQuarterlyLessons].[Cause]          --поле Причина реализовавшегося события
                ,[Solution]       = [vwProjectQuarterlyLessons].[Solution]       --поле Решение (что было сделано)
                ,[Recommendation] = [vwProjectQuarterlyLessons].[Recommendation] --поле Рекомендация
            FROM [ITProject].[vwProjectQuarterlyLessons] AS [vwProjectQuarterlyLessons]
            WHERE   [vwProjectQuarterlyLessons].[ProjectID] IN ({0})
                AND [vwProjectQuarterlyLessons].[Year]      IN ({1})
                AND [vwProjectQuarterlyLessons].[Quarter]   IN ({2})
            ORDER BY [vwProjectQuarterlyLessons].[RowIndex] ASC
        ";
        
        public static string SQLCommandGetProjectBudgetGeneral = @"
            SELECT *
            FROM [ITProject].[ProjectFM] WITH(NOLOCK)
            WHERE   [ProjectID] = {0}
                AND [RowHeader] LIKE '%' + '{1}' + '%'
        ";
    }
}
