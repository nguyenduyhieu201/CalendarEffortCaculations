
using static CalendarEffortCalculationsWinform.Models.Models;
using OfficeOpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;

namespace CalendarEffortCalculationsWinform.BusinessLogic
{
    public class ExcelService
    {
        public List<WorkRangeTimeModel> GetActualWork(ExcelWorksheet worksheet, out int accountColumn)
        {
            var lastCell = worksheet.Cells.End;
            int lastRow = worksheet.Dimension.End.Row;
            int lastColumn = 40;

            int projectCodeColumn = 0;
            accountColumn = 0;
            int fromDateColumn = 0;
            int toDateColumn = 0;
            int hoursPerDateColumn = 0;
            int startRow = 0;
            object[,] data = new object[400, 30];
            ExcelRangeBase titleRange = worksheet.Cells[1, 1, 4, 80];
            object[,] titleExcelValue = titleRange.Value as object[,];
            for (int i = 0; i < titleExcelValue.GetLength(0); i ++)
            {
                for (int j = 0; j < titleExcelValue.GetLength(1); j++)
                {
                    if (titleExcelValue[i, j] is null) continue;
                    if (titleExcelValue[i, j].ToString().ToLower() == "project code")
                    {
                        projectCodeColumn = j;
                        startRow = i + 3;
                    }
                    if (titleExcelValue[i, j].ToString().ToLower() == "username") accountColumn = j;
                    if (titleExcelValue[i, j].ToString().ToLower() == "from date") fromDateColumn = j;
                    if (titleExcelValue[i, j].ToString().ToLower() == "to date") toDateColumn = j;
                    if (titleExcelValue[i, j].ToString().ToLower() .Contains( "hours")) hoursPerDateColumn = j;
                }
            }
            titleRange.Dispose();
            if (worksheet.Cells[lastRow, accountColumn].Value is null)
            {
                while (worksheet.Cells[lastRow, accountColumn].Value is null) lastRow--;
            }
            ExcelRange range = worksheet.Cells[startRow, 1, lastRow, 80];

            // Load data from the range into the DataTable
            object[,] excelValue = range.Value as object[,];
            List<WorkRangeTimeModel> lstRangeModel = new List<WorkRangeTimeModel>();
            for (int i = 0; i < excelValue.GetLength(0); i ++)
            {
                Console.WriteLine($"i la {i}");
                try
                {
                    lstRangeModel.Add(new WorkRangeTimeModel
                    {
                        FromDate = DateTime.FromOADate(int.Parse(excelValue[i, fromDateColumn].ToString())),
                        ToDate = DateTime.FromOADate(int.Parse(excelValue[i, toDateColumn].ToString())),
                        Row = i + startRow,
                        HoursPerDay = double.Parse(excelValue[i, hoursPerDateColumn].ToString()),
                        Account = excelValue[i, accountColumn].ToString(),
                        ProjectCode = excelValue[i, projectCodeColumn].ToString()
                    });
                }
                catch{
                    lstRangeModel.Add(new WorkRangeTimeModel
                    {
                        FromDate = DateTime.Parse(excelValue[i, fromDateColumn].ToString()),
                        ToDate = DateTime.Parse(excelValue[i, toDateColumn].ToString()),
                        Row = i + startRow,
                        HoursPerDay = double.Parse(excelValue[i, hoursPerDateColumn].ToString()),
                        Account = excelValue[i, accountColumn].ToString(),
                        ProjectCode = excelValue[i, projectCodeColumn].ToString()
                    });
                }

            }
            return lstRangeModel;
        }

        public List<PersonalLeaveDay> GetLeave(ExcelWorksheet tmsworksheet)
        {
            var lastCell = tmsworksheet.Cells.End;
            int lastRow = tmsworksheet.Dimension.End.Row;
            int lastColumn = tmsworksheet.Dimension.End.Column;
            int startRow = 1;
            int startRowLeaveSheet = 0;
            int accountColumnLeaveSheet = 0;
            int sumDaysColumn = 0;
            int leaveFromColumn = 0;
            int leaveToColumn = 0;
            int leaveColumn = 0;
            int leaveTypeColumn = 0;

            ExcelRangeBase titleRange = tmsworksheet.Cells[1, 1, 4, 80];
            object[,] titleExcelValue = titleRange.Value as object[,];
            for (int i = 0; i < titleExcelValue.GetLength(0); i++)
            {
                for (int j = 0; j < titleExcelValue.GetLength(1); j++)
                {
                    if (titleExcelValue[i, j] is null) continue;

                    if (titleExcelValue[i, j].ToString().ToLower() == "account")
                    { 
                        accountColumnLeaveSheet = j;
                        startRow = i + 2;
                    }
                    if (titleExcelValue[i, j].ToString().ToLower() == "start date") leaveFromColumn = j;
                    if (titleExcelValue[i, j].ToString().ToLower() == "end date") leaveToColumn = j;
                    if (titleExcelValue[i, j].ToString().ToLower() == ("sum days")) sumDaysColumn = j;
                    if (titleExcelValue[i, j].ToString().ToLower() == ("request type")) leaveColumn = j;
                    if (titleExcelValue[i, j].ToString().ToLower() == ("partialday")) leaveTypeColumn = j;
                }
            }
            titleRange.Dispose();

            ExcelRange range = tmsworksheet.Cells[startRow, 1, lastRow, lastColumn];

            // Load data from the range into the DataTable
            object[,] excelLeaveValues = range.Value as object[,];
            List<PersonalLeaveDay> lstPersonalLeaveModel = new List<PersonalLeaveDay>();
            for (int i = 0; i < excelLeaveValues.GetLength(0); i++)
            {
                if (excelLeaveValues[i, leaveColumn].ToString().ToLower().Contains("nghỉ") || excelLeaveValues[i, leaveColumn].ToString().ToLower().Contains("tạm hoãn"))
                lstPersonalLeaveModel.Add(new PersonalLeaveDay
                {
                    Account = excelLeaveValues[i, accountColumnLeaveSheet].ToString(),
                    SumsLeaveDays = double.Parse(excelLeaveValues[i, sumDaysColumn].ToString()),
                    LeaveFrom = DateTime.FromOADate(int.Parse(excelLeaveValues[i, leaveFromColumn].ToString())),
                    LeaveTo = DateTime.FromOADate(int.Parse(excelLeaveValues[i, leaveToColumn].ToString())),
                    LeaveType = excelLeaveValues[i, leaveTypeColumn].ToString().Contains("Buổi")
                                ? Constants.PartialDayLeave : Constants.FullDayLeave,
                    SumLeaveHours = double.Parse(excelLeaveValues[i, sumDaysColumn].ToString()) * 8
                }); ;
            }
            range.Dispose();
            return lstPersonalLeaveModel;
        }
    
        public List<OTModels> GetOTTimes(ExcelWorksheet otWorksheet)
        {
            int lastRow = otWorksheet.Dimension.End.Row;
            int lastColumn = otWorksheet.Dimension.End.Column;
            int accountColumnInOTSheet = 0;
            int OTSummaryColumnDay = 0;
            int OTSummaryColumnNight = 0;

            int MonthColumn = 0;
            int startRowOTSheet = 0;
            ExcelRangeBase titleRange = otWorksheet.Cells[1, 1, 4, 80];
            object[,] titleExcelValue = titleRange.Value as object[,];
            for (int i = 0; i < titleExcelValue.GetLength(0); i++)
            {
                for (int j = 0; j < titleExcelValue.GetLength(1); j++)
                {
                    if (titleExcelValue[i, j] is null) continue;
                    if (titleExcelValue[i, j].ToString().ToLower().Contains("account")) { 
                        accountColumnInOTSheet = j;
                        startRowOTSheet = i + 2;
                    }

                    if (titleExcelValue[i, j].ToString().ToLower() == "day time") OTSummaryColumnDay = j;
                    if (titleExcelValue[i, j].ToString().ToLower() == "night time") OTSummaryColumnNight = j;

                    if (titleExcelValue[i, j].ToString().ToLower() == ("tháng")) MonthColumn = j;
                }
            }

            titleRange.Dispose();
            List<OTModels> lstOtModel = new List<OTModels>();
            ExcelRange range = otWorksheet.Cells[startRowOTSheet, 1, lastRow, lastColumn];
            // Load data from the range into the DataTable
            object[,] excelLeaveValues = range.Value as object[,];
            //add data 
            for (int i = 0; i < excelLeaveValues.GetLength(0);i ++)
            {
                try
                {
                    lstOtModel.Add(new OTModels
                    {
                        Account = excelLeaveValues[i, accountColumnInOTSheet].ToString(),
                        OverTimeHoursSummary = double.Parse(excelLeaveValues[i, OTSummaryColumnDay].ToString()) + double.Parse(excelLeaveValues[i, OTSummaryColumnNight].ToString()),
                        //Month = int.Parse(excelLeaveValues[i, MonthColumn].ToString()),
                        Month = DateTime.Parse(excelLeaveValues[i, MonthColumn].ToString()).Month

                    });
                }
                catch
                {
                    continue;
                }
            }
            range.Dispose();

            return lstOtModel;
        }

        
    }
}
