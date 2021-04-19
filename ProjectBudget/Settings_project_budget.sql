USE [K2_MDM_REQUEST]
GO

--ProjectQuarteryReport
DECLARE @SettingsTypeCode NVARCHAR(200) = 'project_budget'
DECLARE @SettingsTypeName NVARCHAR(500) = 'Бюджет проекта'

IF (NOT EXISTS(SELECT TOP 1 [ID] FROM [ITProject].[SettingsType] WHERE [Code] = @SettingsTypeCode))
    INSERT INTO [ITProject].[SettingsType] VALUES (@SettingsTypeName, @SettingsTypeCode)
--ELSE UPDATE [ITProject].[SettingsType] SET [Name] = @SettingsTypeName WHERE [Code] = @SettingsTypeCode

DECLARE @SettingsTypeID INT = (SELECT TOP 1 [ID] FROM [ITProject].[SettingsType] WHERE [Code] = @SettingsTypeCode)

--DELETE FROM [ITProject].[Settings] WHERE [SettingsTypeID] = @SettingsTypeID

DECLARE @SettingsTable TABLE(
     [SettingsTypeID]   INT
    ,[Name]             NVARCHAR(500)
    ,[Code]             NVARCHAR(50)
    ,[Value]            XML
)
INSERT INTO @SettingsTable VALUES
 (@SettingsTypeID ,'Папка с файлами'                            ,'FolderPath'               ,'\\run.all.corp\dfs\Проектный офис\Рабочий файл по бюджету\админы')
,(@SettingsTypeID ,'Разделитель номера проекта и его названия'  ,'ProjectNumberDelimeter'   ,'_')
,(@SettingsTypeID ,'Название вкладки'                           ,'WorksheetName'            ,'для К2')
,(@SettingsTypeID ,'Название вкладки полного бюджета'           ,'WorksheetNameFull'        ,'для К2 (деталка)')
,(@SettingsTypeID ,'Строка начала данных'                       ,'RowStartFull'             ,'42')

,(@SettingsTypeID ,'Колонка начала данных'                      ,'ColumnNameStartFull'      ,'A')
,(@SettingsTypeID ,'Колонка конца данных'                       ,'ColumnNameEndFull'        ,'O')
,(@SettingsTypeID ,'Колонка данных по удалению'                 ,'ColumnNameDeleteFull'     ,'P')
,(@SettingsTypeID ,'Значение признака удаления'                 ,'ColumnValueDeleteFull'    ,'убрать')

,(@SettingsTypeID ,'Ячейка Платежи на согласовании'             ,'ApprovingTextFull'        ,'Платежи на согласовании')
,(@SettingsTypeID ,'Ячейка ВСЕГО:'                              ,'SummaryTextFull'          ,'ВСЕГО:')
,(@SettingsTypeID ,'Ячейка Освоение'                            ,'MasteringTextFull'        ,'Освоение')


INSERT INTO [ITProject].[Settings]
SELECT
     [SettingsTable].[SettingsTypeID]
    ,[SettingsTable].[Name]
    ,[SettingsTable].[Code]
    ,[SettingsTable].[Value]
FROM @SettingsTable AS [SettingsTable]
WHERE   NOT EXISTS(
    SELECT TOP 1 [Settings].[ID]
    FROM [ITProject].[Settings] AS [Settings]
    WHERE   [Settings].[Code]           = [SettingsTable].[Code]
        AND [Settings].[SettingsTypeID] = [SettingsTable].[SettingsTypeID])

UPDATE [Settings] SET
      [Settings].[Name]             = [SettingsTable].[Name]
     ,[Settings].[Value]            = [SettingsTable].[Value]
FROM        [ITProject].[Settings]  AS [Settings]
INNER JOIN  @SettingsTable          AS [SettingsTable] ON
        [Settings].[Code]           =   [SettingsTable].[Code]
    AND [Settings].[SettingsTypeID] =   [SettingsTable].[SettingsTypeID]

DELETE FROM [ITProject].[Settings] WHERE [SettingsTypeID] = @SettingsTypeID AND [Code] NOT IN (SELECT [Code] FROM @SettingsTable)

--SELECT * FROM [ITProject].[SettingsType] WHERE [Code]            = @SettingsTypeCode
SELECT * FROM [ITProject].[Settings]     WHERE [SettingsTypeID]  = @SettingsTypeID
