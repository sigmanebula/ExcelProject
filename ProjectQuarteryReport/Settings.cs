using System;
using System.Data;

namespace ProjectQuarteryReport
{
    public static class Settings
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

        public class VariablesClass : Helpers.VariablesClass
        {
            public string ProjectNumberShortName { get; set; }

            public string ProjectID { get; set; }
            public string ProductionCalendarID { get; set; }
            public DataTable ProductionCalendar { get; set; }

            public string GetProductionCalendar(string code, string columnName)
            {
                string result = "";
                for (int i = 0; i < ProductionCalendar.Rows.Count; i++)
                    if (Convert.ToString(ProductionCalendar.Rows[i]["Code"] ?? "") == code)
                        return Convert.ToString(ProductionCalendar.Rows[i][columnName] ?? "");
                return result;
            }
        }

        public class SQLVariablesClass : Helpers.SQLVariablesClass
        {
            public string TemplateFileShortName { get; set; }
            public string NewFileNamePrefix { get; set; }
            public string IsGetErrorMessage { get; set; }
            public string IsAutoFitColumns { get; set; }

            public string WorksheetMainData_Name { get; set; }
            public string MainData_StartRow { get; set; }
            public string IsDataTableHasHeaders { get; set; }
            public string StuffPrefix { get; set; }
            public string KPIData_StartRow { get; set; }
            public string Data_Project_Column { get; set; }


            public string Data_KPI_Waterfall_Mtd_Data_Column { get; set; }
            public string Data_KPI_Waterfall_Mtd_Data_FieldList { get; set; }

            public string Data_KPI_Waterfall_Mtd_Ratings_Column { get; set; }
            public string Data_KPI_Waterfall_Mtd_CommentField_Column { get; set; }
            public string Data_KPI_Waterfall_Mtd_CommentField_FieldList { get; set; }

            public string Data_KPI_Waterfall_GUP_Data_Column { get; set; }
            public string Data_KPI_Waterfall_GUP_Data_FieldList { get; set; }

            public string Data_KPI_Waterfall_GUP_Ratings_Column { get; set; }
            public string Data_KPI_Waterfall_GUP_CommentField_Column { get; set; }
            public string Data_KPI_Waterfall_GUP_CommentField_FieldList { get; set; }

            public string Data_KPI_Waterfall_Mtd_Data_Next_Column { get; set; }
            public string Data_KPI_Waterfall_Mtd_Data_Next_FieldList { get; set; }


            public string Data_KPI_Dynamic_Mtd_Data_Column { get; set; }
            public string Data_KPI_Dynamic_Mtd_Data_FieldList { get; set; }
            public string Data_KPI_Dynamic_Mtd_Ratings_Column { get; set; }
            public string Data_KPI_Dynamic_Mtd_CommentField_Column { get; set; }
            public string Data_KPI_Dynamic_Mtd_CommentField_FieldList { get; set; }

            public string Data_KPI_Dynamic_GUP_Data_Column { get; set; }
            public string Data_KPI_Dynamic_GUP_Data_FieldList { get; set; }
            public string Data_KPI_Dynamic_GUP_Ratings_Column { get; set; }
            public string Data_KPI_Dynamic_GUP_CommentField_Column { get; set; }
            public string Data_KPI_Dynamic_GUP_CommentField_FieldList { get; set; }

            public string Data_KPI_Dynamic_Mtd_Data_Next_Column { get; set; }
            public string Data_KPI_Dynamic_Mtd_Data_Next_FieldList { get; set; }

            public string Data_Risks_Column { get; set; }
            public string Data_Risks_StartRow { get; set; }

            public string Data_Lessons_Column { get; set; }
            public string Data_Lessons_StartRow { get; set; }

            public string Data_BudgetGeneral_StartRow { get; set; }
            public string Data_BudgetGeneral_FieldList { get; set; }

            public string Data_BudgetGeneral_Base_Name { get; set; }
            public string Data_BudgetGeneral_Base_Column { get; set; }

            public string Data_BudgetGeneral_Actual_Name { get; set; }
            public string Data_BudgetGeneral_Actual_Column { get; set; }

            public string Data_BudgetGeneral_Estimate_PartName { get; set; }
            public string Data_BudgetGeneral_Estimate_Column { get; set; }

            public string Data_BudgetFull_StartRow { get; set; }
            public string Data_BudgetFull_IsDataTableHasHeaders { get; set; }
            public string Data_BudgetFull_Column { get; set; }
            public string Data_BudgetFull_FieldList { get; set; }
        }



    }


}
