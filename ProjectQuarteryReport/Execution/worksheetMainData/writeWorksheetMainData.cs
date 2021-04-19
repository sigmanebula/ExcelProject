using OfficeOpenXml;
using System.Data.SqlClient;

namespace ProjectQuarteryReport
{
    public static partial class Execution
    {
        static void writeWorksheetMainData(SqlConnection connection, ExcelWorksheet worksheet, ref string errorText)
        {
            if (errorText == "")
            {
                worksheetMainData_Project_Write(worksheet, connection);

                //Current Waterfall Methodology
                worksheetMainData_KPIWaterfall_Write(
                      worksheet
                    , connection
                    , Settings.Variables.ProductionCalendarID
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Waterfall_Mtd_Data_Column
                    , Settings.SQLVariables.Data_KPI_Waterfall_Mtd_Data_FieldList
                    );

                //Current Waterfall Methodology Ratings
                worksheetMainData_KPIWaterfall_Ratings_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Waterfall_Mtd_Ratings_Column
                    );

                //Current Waterfall Methodology Comment
                worksheetMainData_KPIWaterfall_Comment_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Waterfall_Mtd_CommentField_Column
                    , Settings.SQLVariables.Data_KPI_Waterfall_Mtd_CommentField_FieldList
                    );

                //Current Waterfall GUP
                worksheetMainData_KPIWaterfall_Write(
                      worksheet
                    , connection
                    , Settings.Variables.ProductionCalendarID
                    , Settings.CodeGUP
                    , Settings.SQLVariables.Data_KPI_Waterfall_GUP_Data_Column
                    , Settings.SQLVariables.Data_KPI_Waterfall_GUP_Data_FieldList
                    );

                //Current Waterfall GUP Ratings
                worksheetMainData_KPIWaterfall_Ratings_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeGUP
                    , Settings.SQLVariables.Data_KPI_Waterfall_GUP_Ratings_Column
                    );

                //Current Waterfall GUP Comment
                worksheetMainData_KPIWaterfall_Comment_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeGUP
                    , Settings.SQLVariables.Data_KPI_Waterfall_GUP_CommentField_Column
                    , Settings.SQLVariables.Data_KPI_Waterfall_GUP_CommentField_FieldList
                    );

                //////////////////////////NEXT

                //Next Waterfall Methodology
                worksheetMainData_KPIWaterfall_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Next", "ProductionCalendarID")
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Waterfall_Mtd_Data_Next_Column
                    , Settings.SQLVariables.Data_KPI_Waterfall_Mtd_Data_Next_FieldList
                    );

                //////////////////////////Dynamic

                //Current Dynamic Methodology
                worksheetMainData_KPIDynamic_Write(
                      worksheet
                    , connection
                    , Settings.Variables.ProductionCalendarID
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Dynamic_Mtd_Data_Column
                    , Settings.SQLVariables.Data_KPI_Dynamic_Mtd_Data_FieldList
                    );

                //Current Dynamic Methodology Ratings
                worksheetMainData_KPIDynamic_Ratings_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Dynamic_Mtd_Ratings_Column
                    );

                //Current Dynamic Methodology Comment
                worksheetMainData_KPIDynamic_Comment_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Dynamic_Mtd_CommentField_Column
                    , Settings.SQLVariables.Data_KPI_Dynamic_Mtd_CommentField_FieldList
                    );

                //Current Dynamic GUP
                worksheetMainData_KPIDynamic_Write(
                      worksheet
                    , connection
                    , Settings.Variables.ProductionCalendarID
                    , Settings.CodeGUP
                    , Settings.SQLVariables.Data_KPI_Dynamic_GUP_Data_Column
                    , Settings.SQLVariables.Data_KPI_Dynamic_GUP_Data_FieldList
                    );

                //Current Dynamic GUP Ratings
                worksheetMainData_KPIDynamic_Ratings_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeGUP
                    , Settings.SQLVariables.Data_KPI_Dynamic_GUP_Ratings_Column
                    );

                //Current Dynamic GUP Comment
                worksheetMainData_KPIDynamic_Comment_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Current", "Year")
                    , Settings.Variables.GetProductionCalendar("Current", "Quarter")
                    , Settings.CodeGUP
                    , Settings.SQLVariables.Data_KPI_Dynamic_GUP_CommentField_Column
                    , Settings.SQLVariables.Data_KPI_Dynamic_GUP_CommentField_FieldList
                    );

                //////////////////////////NEXT

                //Next Dynamic Methodology
                worksheetMainData_KPIDynamic_Write(
                      worksheet
                    , connection
                    , Settings.Variables.GetProductionCalendar("Next", "ProductionCalendarID")
                    , Settings.CodeMethodology
                    , Settings.SQLVariables.Data_KPI_Dynamic_Mtd_Data_Next_Column
                    , Settings.SQLVariables.Data_KPI_Dynamic_Mtd_Data_Next_FieldList
                    );

                //Risks
                worksheetMainData_Risks_Write(worksheet, connection);

                //Lessons
                worksheetMainData_Lessons_Write(worksheet, connection);

                //////////////////////////Budget

                worksheetMainData_BudgetGeneral_Write(
                      worksheet
                    , connection
                    , Settings.SQLVariables.Data_BudgetGeneral_Base_Name
                    , Settings.SQLVariables.Data_BudgetGeneral_Base_Column
                    );

                worksheetMainData_BudgetGeneral_Write(
                      worksheet
                    , connection
                    , Settings.SQLVariables.Data_BudgetGeneral_Actual_Name
                    , Settings.SQLVariables.Data_BudgetGeneral_Actual_Column
                    );

                worksheetMainData_BudgetGeneral_Write(
                      worksheet
                    , connection
                    , Settings.SQLVariables.Data_BudgetGeneral_Estimate_PartName
                    , Settings.SQLVariables.Data_BudgetGeneral_Estimate_Column
                    );

                //BudgetFull
                worksheetMainData_BudgetFull_Write(worksheet, connection);

                if (Helpers.Sugar.ConvertStringToBool(Settings.SQLVariables.IsAutoFitColumns))
                    worksheet.Cells.AutoFitColumns();
            }
        }
    }
}



