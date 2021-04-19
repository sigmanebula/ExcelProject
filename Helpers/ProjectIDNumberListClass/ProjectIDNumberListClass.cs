namespace Helpers
{
    public partial class ProjectIDNumberListClass
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

    }
}