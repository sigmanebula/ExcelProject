USE [K2_MDM_REQUEST]
GO

--Settings_ProjectBriefcaseExcelReport
DECLARE @SettingsTypeCode NVARCHAR(200) = 'ProjectBriefcaseExcelReport'
DECLARE @SettingsTypeName NVARCHAR(500) = 'Отчет по портфелю: Вывод данных в excel'

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
 (@SettingsTypeID, 'Путь к папке с шаблоном',                                                                   'FolderPath',                                       'C:\Temp\ProjectBriefcaseExcelReport')
,(@SettingsTypeID, 'Наименование файла шаблона',                                                                'TemplateFileShortName',                            'Template.xlsx')
,(@SettingsTypeID, 'Префикс имени файла',                                                                       'NewFileNamePrefix',                                'Сводный отчёт о реализации портфеля технологических задач БФКО, отчётный период ')
,(@SettingsTypeID, 'Лист технической отладки',                                                                  'WorksheetHidden_Debug_Name',                       'Debug')
,(@SettingsTypeID, 'Лист технической отладки, ячейка стиля, заголовок таблицы',                                 'WorksheetHidden_Debug_StyleCellLabel',             'B2')
,(@SettingsTypeID, 'Лист технической отладки, ячейка стиля, заголовок таблицы',                                 'WorksheetHidden_Debug_StyleCellTableHeader',       'B4')
,(@SettingsTypeID, 'Лист технической отладки, ячейка стиля, заголовок столбца',                                 'WorksheetHidden_Debug_StyleCellColumnHeader',      'B6')
,(@SettingsTypeID, 'Лист технической отладки, ячейка стиля, ячейка таблицы',                                    'WorksheetHidden_Debug_StyleCellInner',             'B8')
,(@SettingsTypeID, 'Лист технической отладки, ячейка цвета границ, заголовок таблицы',                          'WorksheetHidden_Debug_CellForBorderLabel',         'C2')
,(@SettingsTypeID, 'Лист технической отладки, ячейка цвета границ, заголовок таблицы',                          'WorksheetHidden_Debug_CellForBorderTableHeader',   'C4')
,(@SettingsTypeID, 'Лист технической отладки, ячейка цвета границ, заголовок столбца',                          'WorksheetHidden_Debug_CellForBorderColumnHeader',  'C6')
,(@SettingsTypeID, 'Лист технической отладки, ячейка цвета границ, ячейка таблицы',                             'WorksheetHidden_Debug_CellForBorderInner',         'C8')
,(@SettingsTypeID, 'Лист технической отладки, строка данных дебага',                                            'WorksheetHidden_Debug_Row',                        '10')
,(@SettingsTypeID, 'Лист1 открытый, имя',                                                                       'WorksheetShown_1_Name',                            'Титульный')
,(@SettingsTypeID, 'Лист1 открытый, координата, конечная дата у заголовка',                                     'WorksheetShown_1_HeaderEndDateLocation',           'E3')
,(@SettingsTypeID, 'Лист1 открытый, координата, период',                                                        'WorksheetShown_1_PeriodLocation',                  'D10')
,(@SettingsTypeID, 'Лист1 открытый, координата, конечная дата у надписи внизу',                                 'WorksheetShown_1_SmallLabelEndDateLocation',       'C12')
,(@SettingsTypeID, 'Лист2_1 скрытый, Структура портфеля технологических задач, имя',                            'WorksheetHidden_2_1_Name',                         'Скрытый1')
,(@SettingsTypeID, 'Лист2_2 скрытый, имя',                                                                      'WorksheetHidden_2_2_Name',                         'Скрытый2')
,(@SettingsTypeID, 'Лист2_3 скрытый, имя',                                                                      'WorksheetHidden_2_3_Name',                         'Скрытый3')
,(@SettingsTypeID, 'Лист2_4 скрытый, имя',                                                                      'WorksheetHidden_2_4_Name',                         'Скрытый4')
,(@SettingsTypeID, 'Лист2_5 скрытый, имя',                                                                      'WorksheetHidden_2_5_Name',                         'Скрытый5')
,(@SettingsTypeID, 'Лист2_6 скрытый, имя',                                                                      'WorksheetHidden_2_6_Name',                         'Скрытый6')
,(@SettingsTypeID, 'Лист2 открытый, имя',                                                                       'WorksheetShown_2_Name',                            'Структура портфеля')
,(@SettingsTypeID, 'Лист2 открытый, Главный заголовок, ячейка надписи',                                         'WorksheetShown_2_Header_LabelStartCell',           'B2')
,(@SettingsTypeID, 'Лист2 открытый, Заголовок круговой диаграммы, ячейка надписи',                              'WorksheetShown_2_3_Header_LabelStartCell',         'B4')
,(@SettingsTypeID, 'Лист2 открытый, Заголовок столбчатой диаграммы, ячейка надписи',                            'WorksheetShown_2_2_Header_LabelStartCell',         'K4')
,(@SettingsTypeID, 'Лист2 открытый, Здоровье динамических и водопадных проектов, ячейка надписи',               'WorksheetShown_2_3_4_LabelStartCell',              'B23')
,(@SettingsTypeID, 'Лист2 открытый, Здоровье динамических и водопадных проектов, ячеек объединения в строке',   'WorksheetShown_2_3_4_MergeCellCount',              '6')
,(@SettingsTypeID, 'Лист2 открытый, Исключены из портфеля, ячейка надписи',                                     'WorksheetShown_2_5_LabelStartCell',                'K17')

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
