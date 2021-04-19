USE [K2_MDM_REQUEST]
GO

--ProjectQuarteryReport
DECLARE @SettingsTypeCode NVARCHAR(200) = 'ProjectQuarteryReport'
DECLARE @SettingsTypeName NVARCHAR(500) = '����������� ����� �� ������� � Excel'

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
 (@SettingsTypeID ,'���� � ����� � ��������'                                               ,'FolderPath'                                    ,'C:\Temp\ProjectQuarteryReport')
,(@SettingsTypeID ,'������������ ����� �������'                                            ,'TemplateFileShortName'                         ,'Template.xlsx'                )
,(@SettingsTypeID ,'������� ����� �����'                                                   ,'NewFileNamePrefix'                             ,'����������� ����� �� ������� ')
,(@SettingsTypeID ,'�������� ������ ��� ������� ���������?'                                ,'IsGetErrorMessage'                             ,'1'                            )
,(@SettingsTypeID, '����������� �������?'                                                  ,'IsAutoFitColumns'                              ,'0'                            )
                                                                                           
,(@SettingsTypeID ,'�������, Excel ����'                                                   ,'WorksheetMainData_Name'                        ,'��������_����� ����������'    )
,(@SettingsTypeID, '������ ����������� ���������?'                                         ,'IsDataTableHasHeaders'                         ,'0'                            )
,(@SettingsTypeID, '����������� ������������'                                              ,'StuffPrefix'                                   ,', '                           ) --';' + CAST(CHAR(10) AS NVARCHAR(50))
,(@SettingsTypeID, '�������, ��������� ������'                                             ,'MainData_StartRow'                             ,'2'                            )
,(@SettingsTypeID, '�������, ������, ��������� �������'                                    ,'Data_Project_Column'                           ,'A'                            )

,(@SettingsTypeID, '�������, ���, ������'                                                  ,'KPIData_StartRow'                              ,'3'                            )
                                                                                        
,(@SettingsTypeID, '�������, ��� ������� ����������� ������, ��������� �������'            ,'Data_KPI_Waterfall_Mtd_Data_Column'            ,'AP'                           )
,(@SettingsTypeID, '�������, ��� ������� ����������� ������, ������ �������'               ,'Data_KPI_Waterfall_Mtd_Data_FieldList'         ,'Description,CrossProjReqYN,PlanDate,FactDate,DeviationForReport,RaitingForReport,Comment')
,(@SettingsTypeID, '�������, ��� ������� ����������� ������, ��������� �������'            ,'Data_KPI_Waterfall_Mtd_Ratings_Column'         ,'AW'                           )
,(@SettingsTypeID, '�������, ��� ������� ����������� ���� �����������, ��������� �������'  ,'Data_KPI_Waterfall_Mtd_CommentField_Column'    ,'AZ'                           )
,(@SettingsTypeID, '�������, ��� ������� ����������� ���� �����������, ������ �������'     ,'Data_KPI_Waterfall_Mtd_CommentField_FieldList' ,'Comment'                      )
                                                                                           
,(@SettingsTypeID, '�������, ��� ������� ��� ������, ��������� �������'                    ,'Data_KPI_Waterfall_GUP_Data_Column'            ,'BA'                           )
,(@SettingsTypeID, '�������, ��� ������� ��� ������, ������ �������'                       ,'Data_KPI_Waterfall_GUP_Data_FieldList'         ,'Description,CrossProjReqYN,PlanDate,FactDate,DeviationForReport,RaitingForReport,Comment')
,(@SettingsTypeID, '�������, ��� ������� ��� ������, ��������� �������'                    ,'Data_KPI_Waterfall_GUP_Ratings_Column'         ,'BH'                           )
,(@SettingsTypeID, '�������, ��� ������� ��� ���� �����������, ��������� �������'          ,'Data_KPI_Waterfall_GUP_CommentField_Column'    ,'BK'                           )
,(@SettingsTypeID, '�������, ��� ������� ��� ���� �����������, ������ �������'             ,'Data_KPI_Waterfall_GUP_CommentField_FieldList' ,'Comment'                      )
                                                                                           
,(@SettingsTypeID, '�������, ��� ������� ����������� ������ ����. ��., ��������� �������'  ,'Data_KPI_Waterfall_Mtd_Data_Next_Column'       ,'BL'                           )
,(@SettingsTypeID, '�������, ��� ������� ����������� ������ ����. ��., ������ �������'     ,'Data_KPI_Waterfall_Mtd_Data_Next_FieldList'    ,'Description,PlanDate,CrossProjReqYN,ProjectNumberNameInitiator,ProjectNumberNamePerformer,CrossProjReqIconInitiator,CrossProjReqIconPerformer')
                                                                                           
,(@SettingsTypeID, '�������, ��� �������� ����������� ������, ��������� �������'           ,'Data_KPI_Dynamic_Mtd_Data_Column'              ,'BS'                           )
,(@SettingsTypeID, '�������, ��� �������� ����������� ������, ������ �������'              ,'Data_KPI_Dynamic_Mtd_Data_FieldList'           ,'CategoryName,Weight,Description,CrossProjReqYN,PlanVal,FactVal,Deviation,RaitingForReport,ElementCommmentIcon')
,(@SettingsTypeID, '�������, ��� �������� ����������� ������, ��������� �������'           ,'Data_KPI_Dynamic_Mtd_Ratings_Column'           ,'CB'                           )
,(@SettingsTypeID, '�������, ��� �������� ����������� ���� �����������, ��������� �������' ,'Data_KPI_Dynamic_Mtd_CommentField_Column'      ,'CC'                           )
,(@SettingsTypeID, '�������, ��� �������� ����������� ���� �����������, ������ �������'    ,'Data_KPI_Dynamic_Mtd_CommentField_FieldList'   ,'Comment'                      )
                                                                                           
,(@SettingsTypeID, '�������, ��� �������� ��� ������, ��������� �������'                   ,'Data_KPI_Dynamic_GUP_Data_Column'              ,'CD'                           )
,(@SettingsTypeID, '�������, ��� �������� ��� ������, ������ �������'                      ,'Data_KPI_Dynamic_GUP_Data_FieldList'           ,'CategoryName,Weight,Description,CrossProjReqYN,PlanVal,FactVal,Deviation,RaitingForReport,ElementCommmentIcon')
,(@SettingsTypeID, '�������, ��� �������� ��� ������, ��������� �������'                   ,'Data_KPI_Dynamic_GUP_Ratings_Column'           ,'CM'                           )
,(@SettingsTypeID, '�������, ��� �������� ��� ���� �����������, ��������� �������'         ,'Data_KPI_Dynamic_GUP_CommentField_Column'      ,'CN'                           )
,(@SettingsTypeID, '�������, ��� �������� ��� ���� �����������, ������ �������'            ,'Data_KPI_Dynamic_GUP_CommentField_FieldList'   ,'Comment'                      )
                                                                                           
,(@SettingsTypeID, '�������, ��� �������� ����������� ������ ����. ��., ��������� �������' ,'Data_KPI_Dynamic_Mtd_Data_Next_Column'         ,'CO'                           )
,(@SettingsTypeID, '�������, ��� �������� ����������� ������ ����. ��., ������ �������'    ,'Data_KPI_Dynamic_Mtd_Data_Next_FieldList'      ,'CategoryName,Weight,Description,CrossProjReqYN,PlanVal,ProjectNumberNameInitiator,ProjectNumberNamePerformer,CrossProjReqIconInitiator,CrossProjReqIconPerformer')
                                                                                           
,(@SettingsTypeID, '�������, �����, ��������� �������'                                     ,'Data_Risks_Column'                             ,'CX'                           )
,(@SettingsTypeID, '�������, �����, ������'                                                ,'Data_Risks_StartRow'                           ,'3'                            )
                                                                                           
,(@SettingsTypeID, '�������, �����, ��������� �������'                                     ,'Data_Lessons_Column'                           ,'DB'                           )
,(@SettingsTypeID, '�������, �����, ������'                                                ,'Data_Lessons_StartRow'                         ,'3'                            )

,(@SettingsTypeID, '�������, ������ ������ ���-�, ������'                                  ,'Data_BudgetGeneral_StartRow'                   ,'3'                            )
,(@SettingsTypeID, '�������, ������ ������ ���-�, ������ �������'                          ,'Data_BudgetGeneral_FieldList'                  ,'AHR,KV,OREH,FOT,Summary,FACT,Mastering')
                                                                                           
,(@SettingsTypeID, '�������, ������ ������ ���-�, ����, ������������'                      ,'Data_BudgetGeneral_Base_Name'                  ,'����� ������, �������, ���. ���.')
,(@SettingsTypeID, '�������, ������ ������ ���-�, ����, ��������� �������'                 ,'Data_BudgetGeneral_Base_Column'                ,'DF'                           )

,(@SettingsTypeID, '�������, ������ ������ ���-�, ����������, ������������'                ,'Data_BudgetGeneral_Actual_Name'                ,'����� ������, ����������, ���. ���.')
,(@SettingsTypeID, '�������, ������ ������ ���-�, ����������, ��������� �������'           ,'Data_BudgetGeneral_Actual_Column'              ,'DM'                           )

,(@SettingsTypeID, '�������, ������ ������ ���-�, �����, ����� ������������ (LIKE)'        ,'Data_BudgetGeneral_Estimate_PartName'          ,'�������� � �����')
,(@SettingsTypeID, '�������, ������ ������ ���-�, �����, ��������� �������'                ,'Data_BudgetGeneral_Estimate_Column'            ,'DT'                           )

,(@SettingsTypeID, '�������, ������ ���������, ������'                                     ,'Data_BudgetFull_StartRow'                      ,'3'                            )
,(@SettingsTypeID, '�������, ������ ���������, ������ ����������� ���������?'              ,'Data_BudgetFull_IsDataTableHasHeaders'         ,'1'                            )
,(@SettingsTypeID, '�������, ������ ���������, ��������� �������'                          ,'Data_BudgetFull_Column'                        ,'EA'                           )
,(@SettingsTypeID, '�������, ������ ���������, ������ �������'                             ,'Data_BudgetFull_FieldList'                     ,'ProjectBudgetFullTypeName,Code,Name,Plane1Q,Fact1Q,Left1Q,Plane2Q,Fact2Q,Left2Q,Plane3Q,Fact3Q,Left3Q,Plane4Q,Fact4Q,Left4Q,YearSummary')



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
