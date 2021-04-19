USE [K2_MDM_REQUEST]
GO

--ProjectQuarteryReport
DECLARE @SettingsTypeCode NVARCHAR(200) = 'project_budget_cells'
DECLARE @SettingsTypeName NVARCHAR(500) = 'Бюджет проекта (ячейки)'

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
 (@SettingsTypeID ,'Настройки ячеек файла бюджета проекта'    ,'cells'    ,'<settings xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <data xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <ID>1</ID>
    <Description>общий бюджет из версии паспорта проекта по итогам включения в Портфель</Description>
    <RowHeader>A6</RowHeader>
    <AHR>B6</AHR>
    <KV>C6</KV>
    <OREH>D6</OREH>
    <FOT>E6</FOT>
    <Summary>F6</Summary>
    <FACT>G6</FACT>
    <Mastering>H6</Mastering>
  </data>
  <data xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <ID>2</ID>
    <Description>общий бюджет из последней версии паспорта проекта с учетом изменений уровня Спонсора или КРБИ</Description>
    <RowHeader>A7</RowHeader>
    <AHR>B7</AHR>
    <KV>C7</KV>
    <OREH>D7</OREH>
    <FOT>E7</FOT>
    <Summary>F7</Summary>
    <FACT>G7</FACT>
    <Mastering>H7</Mastering>
  </data>
  <data xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <ID>4</ID>
    <Description>доступный объем средств в смете проекта на данный момент с учетом проведенных корректировок</Description>
    <RowHeader>A9</RowHeader>
    <AHR>B9</AHR>
    <KV>C9</KV>
    <OREH>D9</OREH>
    <FOT>E9</FOT>
    <Summary>F9</Summary>
    <FACT>G9</FACT>
    <Mastering>H9</Mastering>
  </data>
</settings>')

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
