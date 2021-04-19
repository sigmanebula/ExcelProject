namespace Helpers
{
    public static partial class SettingsGlobal
    {
        public static string SQLCommandGetSettings = @"
            SELECT
                 [Settings].[Code]
                ,[Settings].[Value]
            FROM        [ITProject].[Settings]     AS [Settings]     WITH(NOLOCK)
            INNER JOIN  [ITProject].[SettingsType] AS [SettingsType] WITH(NOLOCK) ON
                    [SettingsType].[ID]   = [Settings].[SettingsTypeID]
                AND [SettingsType].[Code] IN ({0})
        ";
    }
}
