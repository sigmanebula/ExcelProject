namespace Helpers
{
    public partial class ProjectIDNumberListClass
    {
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
