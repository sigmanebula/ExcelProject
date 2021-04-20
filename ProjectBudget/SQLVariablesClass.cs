namespace ProjectBudget
{
    public class SQLVariablesClass : Helpers.SQLVariablesClass
        {
            public string cells { get; set; }

            public string ProjectNumberDelimeter { get; set; }
            public string WorksheetName { get; set; }
            public string WorksheetNameFull { get; set; }
            public string RowStartFull { get; set; }

            public string ColumnNameStartFull { get; set; }
            public string ColumnNameEndFull { get; set; }
            public string ColumnNameDeleteFull { get; set; }
            public string ColumnValueDeleteFull { get; set; }

            public string ApprovingTextFull { get; set; }
            public string SummaryTextFull { get; set; }
            public string MasteringTextFull { get; set; }
        }


    }
