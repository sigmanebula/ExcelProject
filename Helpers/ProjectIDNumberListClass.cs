namespace Helpers
{
  public class ProjectIDNumberListClass
    {
        public System.Collections.Generic.List<ProjectIDNumberClass> List { get; set; }
        
        public string SQLCommandSelectProjectList = @"
            SELECT
                 [ProjectID]
				,[Number]
			FROM [ITProject].[Project] WITH(NOLOCK)
			WHERE	[IsDeleted] = 0
				AND	[Number]    <> 0
				AND [Number]    IS NOT NULL
                AND [ProjectID] IN ({0})
        ";

        public string SQLCommandSelectProjectsAll = @"
            SELECT
                 [ProjectID]
                ,[Number]
			FROM [ITProject].[Project] WITH(NOLOCK)
			WHERE	[IsDeleted] = 0
				AND	[Number] <> 0
				AND [Number] IS NOT NULL
        ";


        public void Refresh()
        {
            this.List = new System.Collections.Generic.List<ProjectIDNumberClass>();
        }


        public string GetStringListData(string prefix)
        {
            string result = "";

            for (int i = 0; i < List.Count; i++)
                result +=
                    prefix
                    + "ProjectID/Number["
                    + i.ToString()
                    + "] "
                    + List[i].ProjectID.ToString()
                    + "/"
                    + List[i].ProjectNumber.ToString()
                    ;

            return result;
        }

        public string GetStringListData()
        {
            return GetStringListData("\n");
        }


        public void GetData(System.Data.SqlClient.SqlConnection connection, string projectIDList, ref string errorText)
        {
            if (errorText == "")
                try
                {
                    this.List = new System.Collections.Generic.List<ProjectIDNumberClass>();

                    string SQLCommand = (projectIDList != "") ? System.String.Format(this.SQLCommandSelectProjectList, projectIDList) : this.SQLCommandSelectProjectsAll;

                    System.Data.DataTable dataTable = Helpers.SugarSQLConnection.ExecuteSQLCommand(
                          connection
                        , SQLCommand
                        , ""
                        );

                    if (dataTable.Rows.Count == 0)
                        throw new System.Exception("\nОшибка входных данных: ID проекта не найден(ы) в базе");

                    foreach (System.Data.DataRow row in dataTable.Rows)
                        this.List.Add(new ProjectIDNumberClass()
                        {
                            ProjectID = (row["ProjectID"].ToString() == "") ? 0 : int.Parse(row["ProjectID"].ToString()),
                            ProjectNumber = (row["Number"].ToString() == "") ? 0 : int.Parse(row["Number"].ToString())
                        });
                }
                catch (System.Exception exception)
                {
                    errorText += "\nОшибка: не удалось получить данные, причина: " + exception.Message;
                }
        }

    }
}