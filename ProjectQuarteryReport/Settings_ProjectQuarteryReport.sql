USE [K2_MDM_REQUEST]
GO

--ProjectQuarteryReport
DECLARE @SettingsTypeCode NVARCHAR(200) = 'ProjectQuarteryReport'
DECLARE @SettingsTypeName NVARCHAR(500) = 'Квартальный отчет по проекту в Excel'

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
 (@SettingsTypeID ,'Путь к папке с шаблоном'                                               ,'FolderPath'                                    ,'C:\Temp\ProjectQuarteryReport')
,(@SettingsTypeID ,'Наименование файла шаблона'                                            ,'TemplateFileShortName'                         ,'Template.xlsx'                )
,(@SettingsTypeID ,'Префикс имени файла'                                                   ,'NewFileNamePrefix'                             ,'Квартальный отчет по проекту ')
,(@SettingsTypeID ,'Выводить ошибку как обычное сообщение?'                                ,'IsGetErrorMessage'                             ,'1'                            )
,(@SettingsTypeID, 'Выравнивать колонки?'                                                  ,'IsAutoFitColumns'                              ,'0'                            )
                                                                                           
,(@SettingsTypeID ,'Главная, Excel лист'                                                   ,'WorksheetMainData_Name'                        ,'Выгрузка_Общая информация'    )
,(@SettingsTypeID, 'Писать технические заголовки?'                                         ,'IsDataTableHasHeaders'                         ,'0'                            )
,(@SettingsTypeID, 'Разделитель перечисления'                                              ,'StuffPrefix'                                   ,', '                           ) --';' + CAST(CHAR(10) AS NVARCHAR(50))
,(@SettingsTypeID, 'Главная, начальная строка'                                             ,'MainData_StartRow'                             ,'2'                            )
,(@SettingsTypeID, 'Главная, проект, начальный столбец'                                    ,'Data_Project_Column'                           ,'A'                            )

,(@SettingsTypeID, 'Главная, КПЭ, строка'                                                  ,'KPIData_StartRow'                              ,'3'                            )
                                                                                        
,(@SettingsTypeID, 'Главная, КПЭ водопад методология данные, начальный столбец'            ,'Data_KPI_Waterfall_Mtd_Data_Column'            ,'AP'                           )
,(@SettingsTypeID, 'Главная, КПЭ водопад методология данные, список колонок'               ,'Data_KPI_Waterfall_Mtd_Data_FieldList'         ,'Description,CrossProjReqYN,PlanDate,FactDate,DeviationForReport,RaitingForReport,Comment')
,(@SettingsTypeID, 'Главная, КПЭ водопад методология оценки, начальный столбец'            ,'Data_KPI_Waterfall_Mtd_Ratings_Column'         ,'AW'                           )
,(@SettingsTypeID, 'Главная, КПЭ водопад методология поле комментарий, начальный столбец'  ,'Data_KPI_Waterfall_Mtd_CommentField_Column'    ,'AZ'                           )
,(@SettingsTypeID, 'Главная, КПЭ водопад методология поле комментарий, список колонок'     ,'Data_KPI_Waterfall_Mtd_CommentField_FieldList' ,'Comment'                      )
                                                                                           
,(@SettingsTypeID, 'Главная, КПЭ водопад ГУП данные, начальный столбец'                    ,'Data_KPI_Waterfall_GUP_Data_Column'            ,'BA'                           )
,(@SettingsTypeID, 'Главная, КПЭ водопад ГУП данные, список колонок'                       ,'Data_KPI_Waterfall_GUP_Data_FieldList'         ,'Description,CrossProjReqYN,PlanDate,FactDate,DeviationForReport,RaitingForReport,Comment')
,(@SettingsTypeID, 'Главная, КПЭ водопад ГУП оценки, начальный столбец'                    ,'Data_KPI_Waterfall_GUP_Ratings_Column'         ,'BH'                           )
,(@SettingsTypeID, 'Главная, КПЭ водопад ГУП поле комментарий, начальный столбец'          ,'Data_KPI_Waterfall_GUP_CommentField_Column'    ,'BK'                           )
,(@SettingsTypeID, 'Главная, КПЭ водопад ГУП поле комментарий, список колонок'             ,'Data_KPI_Waterfall_GUP_CommentField_FieldList' ,'Comment'                      )
                                                                                           
,(@SettingsTypeID, 'Главная, КПЭ водопад методология данные след. кв., начальный столбец'  ,'Data_KPI_Waterfall_Mtd_Data_Next_Column'       ,'BL'                           )
,(@SettingsTypeID, 'Главная, КПЭ водопад методология данные след. кв., список колонок'     ,'Data_KPI_Waterfall_Mtd_Data_Next_FieldList'    ,'Description,PlanDate,CrossProjReqYN,ProjectNumberNameInitiator,ProjectNumberNamePerformer,CrossProjReqIconInitiator,CrossProjReqIconPerformer')
                                                                                           
,(@SettingsTypeID, 'Главная, КПЭ динамика методология данные, начальный столбец'           ,'Data_KPI_Dynamic_Mtd_Data_Column'              ,'BS'                           )
,(@SettingsTypeID, 'Главная, КПЭ динамика методология данные, список колонок'              ,'Data_KPI_Dynamic_Mtd_Data_FieldList'           ,'CategoryName,Weight,Description,CrossProjReqYN,PlanVal,FactVal,Deviation,RaitingForReport,ElementCommmentIcon')
,(@SettingsTypeID, 'Главная, КПЭ динамика методология оценки, начальный столбец'           ,'Data_KPI_Dynamic_Mtd_Ratings_Column'           ,'CB'                           )
,(@SettingsTypeID, 'Главная, КПЭ динамика методология поле комментарий, начальный столбец' ,'Data_KPI_Dynamic_Mtd_CommentField_Column'      ,'CC'                           )
,(@SettingsTypeID, 'Главная, КПЭ динамика методология поле комментарий, список колонок'    ,'Data_KPI_Dynamic_Mtd_CommentField_FieldList'   ,'Comment'                      )
                                                                                           
,(@SettingsTypeID, 'Главная, КПЭ динамика ГУП данные, начальный столбец'                   ,'Data_KPI_Dynamic_GUP_Data_Column'              ,'CD'                           )
,(@SettingsTypeID, 'Главная, КПЭ динамика ГУП данные, список колонок'                      ,'Data_KPI_Dynamic_GUP_Data_FieldList'           ,'CategoryName,Weight,Description,CrossProjReqYN,PlanVal,FactVal,Deviation,RaitingForReport,ElementCommmentIcon')
,(@SettingsTypeID, 'Главная, КПЭ динамика ГУП оценки, начальный столбец'                   ,'Data_KPI_Dynamic_GUP_Ratings_Column'           ,'CM'                           )
,(@SettingsTypeID, 'Главная, КПЭ динамика ГУП поле комментарий, начальный столбец'         ,'Data_KPI_Dynamic_GUP_CommentField_Column'      ,'CN'                           )
,(@SettingsTypeID, 'Главная, КПЭ динамика ГУП поле комментарий, список колонок'            ,'Data_KPI_Dynamic_GUP_CommentField_FieldList'   ,'Comment'                      )
                                                                                           
,(@SettingsTypeID, 'Главная, КПЭ динамика методология данные след. кв., начальный столбец' ,'Data_KPI_Dynamic_Mtd_Data_Next_Column'         ,'CO'                           )
,(@SettingsTypeID, 'Главная, КПЭ динамика методология данные след. кв., список колонок'    ,'Data_KPI_Dynamic_Mtd_Data_Next_FieldList'      ,'CategoryName,Weight,Description,CrossProjReqYN,PlanVal,ProjectNumberNameInitiator,ProjectNumberNamePerformer,CrossProjReqIconInitiator,CrossProjReqIconPerformer')
                                                                                           
,(@SettingsTypeID, 'Главная, Риски, начальный столбец'                                     ,'Data_Risks_Column'                             ,'CX'                           )
,(@SettingsTypeID, 'Главная, Риски, строка'                                                ,'Data_Risks_StartRow'                           ,'3'                            )
                                                                                           
,(@SettingsTypeID, 'Главная, Уроки, начальный столбец'                                     ,'Data_Lessons_Column'                           ,'DB'                           )
,(@SettingsTypeID, 'Главная, Уроки, строка'                                                ,'Data_Lessons_StartRow'                         ,'3'                            )

,(@SettingsTypeID, 'Главная, Бюджет общаяя инф-я, строка'                                  ,'Data_BudgetGeneral_StartRow'                   ,'3'                            )
,(@SettingsTypeID, 'Главная, Бюджет общаяя инф-я, список колонок'                          ,'Data_BudgetGeneral_FieldList'                  ,'AHR,KV,OREH,FOT,Summary,FACT,Mastering')
                                                                                           
,(@SettingsTypeID, 'Главная, Бюджет общаяя инф-я, база, наименование'                      ,'Data_BudgetGeneral_Base_Name'                  ,'Общий бюджет, базовый, тыс. руб.')
,(@SettingsTypeID, 'Главная, Бюджет общаяя инф-я, база, начальный столбец'                 ,'Data_BudgetGeneral_Base_Column'                ,'DF'                           )

,(@SettingsTypeID, 'Главная, Бюджет общаяя инф-я, актуальный, наименование'                ,'Data_BudgetGeneral_Actual_Name'                ,'Общий бюджет, актуальный, тыс. руб.')
,(@SettingsTypeID, 'Главная, Бюджет общаяя инф-я, актуальный, начальный столбец'           ,'Data_BudgetGeneral_Actual_Column'              ,'DM'                           )

,(@SettingsTypeID, 'Главная, Бюджет общаяя инф-я, смета, часть наименования (LIKE)'        ,'Data_BudgetGeneral_Estimate_PartName'          ,'выделено в смету')
,(@SettingsTypeID, 'Главная, Бюджет общаяя инф-я, смета, начальный столбец'                ,'Data_BudgetGeneral_Estimate_Column'            ,'DT'                           )

,(@SettingsTypeID, 'Главная, Бюджет детальный, строка'                                     ,'Data_BudgetFull_StartRow'                      ,'3'                            )
,(@SettingsTypeID, 'Главная, Бюджет детальный, писать технические заголовки?'              ,'Data_BudgetFull_IsDataTableHasHeaders'         ,'1'                            )
,(@SettingsTypeID, 'Главная, Бюджет детальный, начальный столбец'                          ,'Data_BudgetFull_Column'                        ,'EA'                           )
,(@SettingsTypeID, 'Главная, Бюджет детальный, список колонок'                             ,'Data_BudgetFull_FieldList'                     ,'ProjectBudgetFullTypeName,Code,Name,Plane1Q,Fact1Q,Left1Q,Plane2Q,Fact2Q,Left2Q,Plane3Q,Fact3Q,Left3Q,Plane4Q,Fact4Q,Left4Q,YearSummary')



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
