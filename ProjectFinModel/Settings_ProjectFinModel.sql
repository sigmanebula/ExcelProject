USE [K2_MDM_REQUEST]
GO

--ProjectQuarteryReport
DECLARE @SettingsTypeCode NVARCHAR(200) = 'project_finmodel'
DECLARE @SettingsTypeName NVARCHAR(500) = 'Фин. модель проекта'

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
 (@SettingsTypeID ,'Папка с файлами'                                        ,'FolderPath'                                   ,'C:\temp\FinModel')
,(@SettingsTypeID ,'Название вкладки'                                       ,'WorksheetName'                                ,'РЕСУРСЫ ВНЕДРЕНИЯ')
,(@SettingsTypeID ,'Выводить ошибку как обычное сообщение?'                 ,'IsGetErrorMessage'                            ,'1')
,(@SettingsTypeID ,'Выводить сообщение дебага данных SQL?'                  ,'IsDebugSQL'                                   ,'0')
,(@SettingsTypeID ,'Выравнивать колонки автоматически?'                     ,'IsAutoFitColumns'                             ,'0')
,(@SettingsTypeID ,'Разделитель номера проекта и названия'                  ,'ProjectNumberDelimeter'                       ,'_')
,(@SettingsTypeID ,'Список колонок отчёта формирования ресурсоёмкости'      ,'ReportProjectResourceIntensityColumnList'     ,'TextIdentificator,IsOldAllocation,DepartmentBlockTypeName,Department_FullName,Role_System_Specialization,ResourceAllocation_Year,ResourceAllocation_Quarter,ResourceAllocation_Allocated')
,(@SettingsTypeID ,'Ресурсы ИТ-развития'                                    ,'It_development_SQL'                           ,'Ресурсы ИТ-развития')
,(@SettingsTypeID ,'Прочие ИТ-ресурсы'                                      ,'It_other_SQL'                                 ,'Прочие ИТ-ресурсы')
,(@SettingsTypeID ,'Бизнес- и функциональные подразделения'                 ,'Business_functionality_SQL'                   ,'Бизнес- и функциональные подразделения')
,(@SettingsTypeID ,'Ячейка'                                                 ,'It_development_Excel'                         ,'Ресурсы развития ИТ:')
,(@SettingsTypeID ,'Ячейка'                                                 ,'It_other_Excel'                               ,'Прочие ресурсы ИТ:')
,(@SettingsTypeID ,'Ячейка'                                                 ,'Business_functionality_Excel'                 ,'Ресурсы бизнес-подразделений | функциональных подразделений:')
,(@SettingsTypeID ,'Ячейка'                                                 ,'Role_Excel'                                   ,'Роль\направление')
,(@SettingsTypeID ,'Ячейка'                                                 ,'Department_Excel'                             ,'Подразделение')
,(@SettingsTypeID ,'Ячейка'                                                 ,'LastColumnSummary_Excel'                      ,'ИТОГО Объем работ(ч/д)')
,(@SettingsTypeID ,'Ячейка'                                                 ,'SummaryRow_Excel'                             ,'ИТОГО по всем внутренним ресурсам Банка: ')
,(@SettingsTypeID ,'Ячейка'                                                 ,'QuarterPreText'                               ,'Квартал ')
,(@SettingsTypeID ,'Пароль'                                                 ,'ExcelPassword'                                ,'prof')
,(@SettingsTypeID ,'Ошибка, если в файле кварталов больше, чем в данных?'   ,'ExceptionNoDateForFileQuarter'                ,'0')



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
