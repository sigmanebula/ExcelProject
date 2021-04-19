USE [K2_MDM_REQUEST]
GO

--Settings_ProjectBriefcaseExcelReport
DECLARE @SettingsTypeCode NVARCHAR(200) = 'ProjectBriefcaseExcelReport'
DECLARE @SettingsTypeName NVARCHAR(500) = '����� �� ��������: ����� ������ � excel'

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
 (@SettingsTypeID, '���� � ����� � ��������',                                                                   'FolderPath',                                       'C:\Temp\ProjectBriefcaseExcelReport')
,(@SettingsTypeID, '������������ ����� �������',                                                                'TemplateFileShortName',                            'Template.xlsx')
,(@SettingsTypeID, '������� ����� �����',                                                                       'NewFileNamePrefix',                                '������� ����� � ���������� �������� ��������������� ����� ����, �������� ������ ')
,(@SettingsTypeID, '���� ����������� �������',                                                                  'WorksheetHidden_Debug_Name',                       'Debug')
,(@SettingsTypeID, '���� ����������� �������, ������ �����, ��������� �������',                                 'WorksheetHidden_Debug_StyleCellLabel',             'B2')
,(@SettingsTypeID, '���� ����������� �������, ������ �����, ��������� �������',                                 'WorksheetHidden_Debug_StyleCellTableHeader',       'B4')
,(@SettingsTypeID, '���� ����������� �������, ������ �����, ��������� �������',                                 'WorksheetHidden_Debug_StyleCellColumnHeader',      'B6')
,(@SettingsTypeID, '���� ����������� �������, ������ �����, ������ �������',                                    'WorksheetHidden_Debug_StyleCellInner',             'B8')
,(@SettingsTypeID, '���� ����������� �������, ������ ����� ������, ��������� �������',                          'WorksheetHidden_Debug_CellForBorderLabel',         'C2')
,(@SettingsTypeID, '���� ����������� �������, ������ ����� ������, ��������� �������',                          'WorksheetHidden_Debug_CellForBorderTableHeader',   'C4')
,(@SettingsTypeID, '���� ����������� �������, ������ ����� ������, ��������� �������',                          'WorksheetHidden_Debug_CellForBorderColumnHeader',  'C6')
,(@SettingsTypeID, '���� ����������� �������, ������ ����� ������, ������ �������',                             'WorksheetHidden_Debug_CellForBorderInner',         'C8')
,(@SettingsTypeID, '���� ����������� �������, ������ ������ ������',                                            'WorksheetHidden_Debug_Row',                        '10')
,(@SettingsTypeID, '����1 ��������, ���',                                                                       'WorksheetShown_1_Name',                            '���������')
,(@SettingsTypeID, '����1 ��������, ����������, �������� ���� � ���������',                                     'WorksheetShown_1_HeaderEndDateLocation',           'E3')
,(@SettingsTypeID, '����1 ��������, ����������, ������',                                                        'WorksheetShown_1_PeriodLocation',                  'D10')
,(@SettingsTypeID, '����1 ��������, ����������, �������� ���� � ������� �����',                                 'WorksheetShown_1_SmallLabelEndDateLocation',       'C12')
,(@SettingsTypeID, '����2_1 �������, ��������� �������� ��������������� �����, ���',                            'WorksheetHidden_2_1_Name',                         '�������1')
,(@SettingsTypeID, '����2_2 �������, ���',                                                                      'WorksheetHidden_2_2_Name',                         '�������2')
,(@SettingsTypeID, '����2_3 �������, ���',                                                                      'WorksheetHidden_2_3_Name',                         '�������3')
,(@SettingsTypeID, '����2_4 �������, ���',                                                                      'WorksheetHidden_2_4_Name',                         '�������4')
,(@SettingsTypeID, '����2_5 �������, ���',                                                                      'WorksheetHidden_2_5_Name',                         '�������5')
,(@SettingsTypeID, '����2_6 �������, ���',                                                                      'WorksheetHidden_2_6_Name',                         '�������6')
,(@SettingsTypeID, '����2 ��������, ���',                                                                       'WorksheetShown_2_Name',                            '��������� ��������')
,(@SettingsTypeID, '����2 ��������, ������� ���������, ������ �������',                                         'WorksheetShown_2_Header_LabelStartCell',           'B2')
,(@SettingsTypeID, '����2 ��������, ��������� �������� ���������, ������ �������',                              'WorksheetShown_2_3_Header_LabelStartCell',         'B4')
,(@SettingsTypeID, '����2 ��������, ��������� ���������� ���������, ������ �������',                            'WorksheetShown_2_2_Header_LabelStartCell',         'K4')
,(@SettingsTypeID, '����2 ��������, �������� ������������ � ���������� ��������, ������ �������',               'WorksheetShown_2_3_4_LabelStartCell',              'B23')
,(@SettingsTypeID, '����2 ��������, �������� ������������ � ���������� ��������, ����� ����������� � ������',   'WorksheetShown_2_3_4_MergeCellCount',              '6')
,(@SettingsTypeID, '����2 ��������, ��������� �� ��������, ������ �������',                                     'WorksheetShown_2_5_LabelStartCell',                'K17')

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
