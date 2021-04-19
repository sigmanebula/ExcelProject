USE [K2_MDM_REQUEST]
GO

--ProjectQuarteryReport
DECLARE @SettingsTypeCode NVARCHAR(200) = 'project_finmodel'
DECLARE @SettingsTypeName NVARCHAR(500) = '���. ������ �������'

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
 (@SettingsTypeID ,'����� � �������'                                        ,'FolderPath'                                   ,'C:\temp\FinModel')
,(@SettingsTypeID ,'�������� �������'                                       ,'WorksheetName'                                ,'������� ���������')
,(@SettingsTypeID ,'�������� ������ ��� ������� ���������?'                 ,'IsGetErrorMessage'                            ,'1')
,(@SettingsTypeID ,'�������� ��������� ������ ������ SQL?'                  ,'IsDebugSQL'                                   ,'0')
,(@SettingsTypeID ,'����������� ������� �������������?'                     ,'IsAutoFitColumns'                             ,'0')
,(@SettingsTypeID ,'����������� ������ ������� � ��������'                  ,'ProjectNumberDelimeter'                       ,'_')
,(@SettingsTypeID ,'������ ������� ������ ������������ �������������'      ,'ReportProjectResourceIntensityColumnList'     ,'TextIdentificator,IsOldAllocation,DepartmentBlockTypeName,Department_FullName,Role_System_Specialization,ResourceAllocation_Year,ResourceAllocation_Quarter,ResourceAllocation_Allocated')
,(@SettingsTypeID ,'������� ��-��������'                                    ,'It_development_SQL'                           ,'������� ��-��������')
,(@SettingsTypeID ,'������ ��-�������'                                      ,'It_other_SQL'                                 ,'������ ��-�������')
,(@SettingsTypeID ,'������- � �������������� �������������'                 ,'Business_functionality_SQL'                   ,'������- � �������������� �������������')
,(@SettingsTypeID ,'������'                                                 ,'It_development_Excel'                         ,'������� �������� ��:')
,(@SettingsTypeID ,'������'                                                 ,'It_other_Excel'                               ,'������ ������� ��:')
,(@SettingsTypeID ,'������'                                                 ,'Business_functionality_Excel'                 ,'������� ������-������������� | �������������� �������������:')
,(@SettingsTypeID ,'������'                                                 ,'Role_Excel'                                   ,'����\�����������')
,(@SettingsTypeID ,'������'                                                 ,'Department_Excel'                             ,'�������������')
,(@SettingsTypeID ,'������'                                                 ,'LastColumnSummary_Excel'                      ,'����� ����� �����(�/�)')
,(@SettingsTypeID ,'������'                                                 ,'SummaryRow_Excel'                             ,'����� �� ���� ���������� �������� �����:�')
,(@SettingsTypeID ,'������'                                                 ,'QuarterPreText'                               ,'������� ')
,(@SettingsTypeID ,'������'                                                 ,'ExcelPassword'                                ,'prof')
,(@SettingsTypeID ,'������, ���� � ����� ��������� ������, ��� � ������?'   ,'ExceptionNoDateForFileQuarter'                ,'0')



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
