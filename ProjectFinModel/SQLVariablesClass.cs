namespace ProjectFinModel
{
    public class SQLVariablesClass : Helpers.SQLVariablesClass
        {
            public string WorksheetName { get; set; }
            public string IsGetErrorMessage { get; set; }
            public string IsDebugSQL { get; set; }
            public string IsAutoFitColumns { get; set; }
            public string ProjectNumberDelimeter { get; set; }
            public string ReportProjectResourceIntensityColumnList { get; set; }

            public string It_development_SQL { get; set; }
            public string It_other_SQL { get; set; }
            public string Business_functionality_SQL { get; set; }

            public string It_development_Excel { get; set; }
            public string It_other_Excel { get; set; }
            public string Business_functionality_Excel { get; set; }

            public string Role_Excel { get; set; }
            public string Department_Excel { get; set; }
            public string LastColumnSummary_Excel { get; set; }
            public string SummaryRow_Excel { get; set; }
            public string QuarterPreText { get; set; }

            public string ExcelPassword { get; set; }
            public string ExceptionNoDateForFileQuarter { get; set; }
        }



    }




