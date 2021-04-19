USE [K2_MDM_REQUEST]
GO

--ProjectQuarteryReport
DECLARE @SettingsTypeCode NVARCHAR(200) = 'project_budget'
DECLARE @SettingsTypeName NVARCHAR(500) = '������ �������'

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
 (@SettingsTypeID ,'����� � �������'                            ,'FolderPath'               ,'\\run.all.corp\dfs\��������� ����\������� ���� �� �������\������')
,(@SettingsTypeID ,'����������� ������ ������� � ��� ��������'  ,'ProjectNumberDelimeter'   ,'_')
,(@SettingsTypeID ,'�������� �������'                           ,'WorksheetName'            ,'��� �2')
,(@SettingsTypeID ,'�������� ������� ������� �������'           ,'WorksheetNameFull'        ,'��� �2 (�������)')
,(@SettingsTypeID ,'������ ������ ������'                       ,'RowStartFull'             ,'42')

,(@SettingsTypeID ,'������� ������ ������'                      ,'ColumnNameStartFull'      ,'A')
,(@SettingsTypeID ,'������� ����� ������'                       ,'ColumnNameEndFull'        ,'O')
,(@SettingsTypeID ,'������� ������ �� ��������'                 ,'ColumnNameDeleteFull'     ,'P')
,(@SettingsTypeID ,'�������� �������� ��������'                 ,'ColumnValueDeleteFull'    ,'������')

,(@SettingsTypeID ,'������ ������� �� ������������'             ,'ApprovingTextFull'        ,'������� �� ������������')
,(@SettingsTypeID ,'������ �����:'                              ,'SummaryTextFull'          ,'�����:')
,(@SettingsTypeID ,'������ ��������'                            ,'MasteringTextFull'        ,'��������')


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
