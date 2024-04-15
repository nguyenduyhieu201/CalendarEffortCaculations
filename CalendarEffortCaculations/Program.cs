// See https://aka.ms/new-console-template for more information
using static System.Net.Mime.MediaTypeNames;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using CalendarEffortCaculations;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Reflection;
using Microsoft.VisualBasic;
using System.Xml.Linq;

Console.WriteLine("Hello, World!");


//truy cập vào file excel đang chạy
static Tuple<Excel.Workbook, Excel.Worksheet, Excel.Range> GetOpeningExcelFile(string name, string sheetname)
{
    bool wasFoundRunning = false;
    Excel.Workbook workbook = null;
    Excel.Worksheet worksheet = null;
    Excel.Range range = null;
    try
    {
        var xlApp = (Application)Marshal2.GetActiveObject("Excel.Application");
        Excel.Workbooks xlBooks = xlApp.Workbooks;
        foreach (Excel.Workbook xlbook in xlBooks)
        {
            if (xlbook.Name.ToLower().Trim().Contains(name.ToLower().Trim()))
            {
                workbook = xlbook;
                worksheet = (Excel.Worksheet)workbook.Worksheets[sheetname];
                range = worksheet.Cells[1, 1];

            }
        }
        var numBooks = xlBooks.Count;
        wasFoundRunning = true;
        xlApp.Visible = true;
        Marshal.ReleaseComObject(xlApp);

        // Call the garbage collector to release any remaining references
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
    catch (Exception e)
    {
        //Log.Error(e.Message);
        Console.WriteLine(e.ToString());
        wasFoundRunning = false;
    }

    return Tuple.Create(workbook, worksheet, range);
}

//tính số ngày làm việc trong tháng
static int WorkingDaysInMonth(int year, int month, DateTime startDate, DateTime endDate)
{
    var listHoliday = new List<DateTime>
    {
        new DateTime(2024,1,1),
        new DateTime(2024,2,8),
        new DateTime(2024,2,9),
        new DateTime(2024,2,10),
        new DateTime(2024,2,11),
        new DateTime(2024,2,12),
        new DateTime(2024,2,13),
        new DateTime(2024,2,14),
        new DateTime(2024,4,18),
        new DateTime(2024,4,29),
        new DateTime(2024,4,30),
        new DateTime(2024,5,1),
        new DateTime(2024,9,2),
        new DateTime(2024,9,3),
    };
    // Get the number of days in the month
    int numDays = DateTime.DaysInMonth(year, month);

    // Initialize a counter for working days
    int workingDays = 0;

    // Iterate through each day in the month
    for (int day = 1; day <= numDays; day++)
    {
        DateTime date = new DateTime(year, month, day);
        if (month == startDate.Month && day < startDate.Day)
        {
            continue;
        }

        if (month == endDate.Month && day > endDate.Day)
        {
            continue;
        }

        // Check if the day is a weekday (Monday to Friday)
        if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday && !listHoliday.Any(holiday => holiday.Equals(date)))
        {
            workingDays++;
        }
    }

    return workingDays;
}
//tính số ngày làm việc trong một quãng thời gian, chia ra theo tháng
static Dictionary<Tuple<int, int>, Tuple<int, int>> WorkingDaysInDuration(DateTime startDate, DateTime endDate)
{
    // Initialize a dictionary to store the count of working days for each month
    Dictionary<Tuple<int, int>, Tuple<int, int>> workingDaysPerMonth = new Dictionary<Tuple<int, int>, Tuple<int, int>>();

    // Iterate through each month within the duration
    while (startDate <= endDate)
    {
        int year = startDate.Year;
        int month = startDate.Month;
        // Calculate the number of working days in the current month
        int workingDays = WorkingDaysInMonth(year, month, startDate, endDate);
        int expectedWorkingHours = workingDays * 8;
        // Store the count in the dictionary
        workingDaysPerMonth.Add(new Tuple<int, int>(year, month), new Tuple<int, int> (workingDays, expectedWorkingHours));
        // Move to the next month
        startDate = new DateTime(year, month, 1).AddMonths(1);
    }

    return workingDaysPerMonth;
}

//tính số ngày nghỉ trong tháng
static int LeaveDaysInMonth(int year, int month, DateTime startDate, DateTime endDate)
{
    var listHoliday = new List<DateTime>
    {
        new DateTime(2024,1,1),
        new DateTime(2024,2,8),
        new DateTime(2024,2,9),
        new DateTime(2024,2,10),
        new DateTime(2024,2,11),
        new DateTime(2024,2,12),
        new DateTime(2024,2,13),
        new DateTime(2024,2,14),
        new DateTime(2024,4,18),
        new DateTime(2024,4,29),
        new DateTime(2024,4,30),
        new DateTime(2024,5,1),
        new DateTime(2024,9,2),
        new DateTime(2024,9,3),
    };
    // Get the number of days in the month
    int numDays = DateTime.DaysInMonth(year, month);

    // Initialize a counter for working days
    int leaveDays = 0;

    // Iterate through each day in the month
    for (int day = 1; day <= numDays; day++)
    {
        DateTime date = new DateTime(year, month, day);
        if (month == startDate.Month && day < startDate.Day)
        {
            continue;
        }

        if (month == endDate.Month && day > endDate.Day)
        {
            continue;
        }

        // Check if the day is a weekday (Monday to Friday)
        if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday && !listHoliday.Any(holiday => holiday.Equals(date)))
        {
            leaveDays++;
        }
    }

    return leaveDays;
}
//
static Dictionary<Tuple<int, int>, int> LeaveDaysInDuration(DateTime startDate, DateTime endDate)
{
    // Initialize a dictionary to store the count of working days for each month
    Dictionary<Tuple<int, int>, int> leaveDaysPerMonth = new Dictionary<Tuple<int, int>, int>();

    // Iterate through each month within the duration
    while (startDate <= endDate)
    {
        int year = startDate.Year;
        int month = startDate.Month;
        // Calculate the number of working days in the current month
        int leaveDays = LeaveDaysInMonth(year, month, startDate, endDate);
        // Store the count in the dictionary
        leaveDaysPerMonth.Add(new Tuple<int, int>(year, month), leaveDays);
        // Move to the next month
        startDate = new DateTime(year, month, 1).AddMonths(1);
    }

    return leaveDaysPerMonth;
}

#region Lấy số ngày OT theo tháng
var excelOT = GetOpeningExcelFile("Simulate", "OT");
int accountColumnInOTSheet = 3;
int OTSummaryColumn = 14;
int MonthColumn = 16;
int startRowOTSheet = 3;

List<Models> GetlstOTModel(int accountColumnInOTSheet, int OTSummaryColumn, int MonthColumn, int startRowOTSheet)
{
    List<Models> lstOverTimePersonal = new List<Models>();
    do
    {
        try
        {
            lstOverTimePersonal.Add(new Models
            {
                Account = excelOT.Item3.Cells[startRowOTSheet, accountColumnInOTSheet].Value,
                OverTimeHoursSummary = excelOT.Item3.Cells[startRowOTSheet, OTSummaryColumn].Value,
                Month = (int)excelOT.Item3.Cells[startRowOTSheet, MonthColumn].Value,
            });
            startRowOTSheet++;
        }
        catch (Exception ex)
        {
            break;
        }
    }
    while (excelOT.Item3.Cells[startRowOTSheet, accountColumnInOTSheet] != null
         && excelOT.Item3.Cells[startRowOTSheet, OTSummaryColumn] != null
         && excelOT.Item3.Cells[startRowOTSheet, MonthColumn] != null);
    return lstOverTimePersonal;
}
#endregion

#region Lấy số ngày nghỉ theo tháng
var excelTMS = GetOpeningExcelFile("Simulate", "TMS");
int startRowLeaveSheet = 2;
int accountColumnLeaveSheet = 3;
int sumDaysColumn = 16;
int leaveFromColumn = 10;
int leaveToColumn = 11;
int leaveColumn = 8;
int leaveTypeColumn = 18;
const int PartialDayLeave = 0;
const int FullDayLeave = 1;

List<PersonalLeaveDay> GetLstLeaveModels(int startRowLeaveSheet, int accountColumnLeaveSheet, int sumDaysColumn, int leaveFromColumn, int leaveToColumn, int leaveColumn, int leaveTypeColumn)
{
    List<PersonalLeaveDay> lstPersonalLeaveDay = new List<PersonalLeaveDay>();
    while (excelTMS.Item3.Cells[startRowLeaveSheet, accountColumnLeaveSheet].Value != null && excelTMS.Item3.Cells[startRowLeaveSheet, sumDaysColumn].Value != null)
    {
        var leaveContent = (excelTMS.Item3.Cells[startRowLeaveSheet, leaveColumn].Value.ToString().ToLower().Trim());
        if ((leaveContent.Contains("nghỉ") || leaveContent.Contains("tạm hoãn")))
        {
            lstPersonalLeaveDay.Add(new PersonalLeaveDay
            {
                Account = excelTMS.Item3.Cells[startRowLeaveSheet, accountColumnLeaveSheet].Value,
                SumsLeaveDays = excelTMS.Item3.Cells[startRowLeaveSheet, sumDaysColumn].Value,
                LeaveFrom = excelTMS.Item3.Cells[startRowLeaveSheet, leaveFromColumn].Value,
                LeaveTo = excelTMS.Item3.Cells[startRowLeaveSheet, leaveToColumn].Value,
                LeaveType = excelTMS.Item3.Cells[startRowLeaveSheet, leaveTypeColumn].Value.ToString().Contains("Buổi")
                                ? PartialDayLeave : FullDayLeave,
                SumLeaveHours = excelTMS.Item3.Cells[startRowLeaveSheet, sumDaysColumn].Value * 8
            });
            startRowLeaveSheet++;
        }
        else
        {
            startRowLeaveSheet++;
        }
    }

    return lstPersonalLeaveDay;
}

Console.WriteLine();
#endregion

#region Lấy số ngày làm việc thực tế
var excelCalendarSheet = GetOpeningExcelFile("Simulate", "Calendar");
int startRow = 3;
int fromDateColumn = 7;
int toDateColumn = 8;
int hoursPerDayColumn = 9;
int accountColumn = 5;
int projectCodeColumn = 3;

List<RangeTimeModel> GetActualWorkingDays(int startRow, int fromDateColumn, int toDateColumn, int hoursPerDayColumn, int accountColumn, int projectCodeColumn)
{
    List<RangeTimeModel> lstRangeModel = new List<RangeTimeModel>();

    do
    {
        try
        {
            lstRangeModel.Add(new RangeTimeModel
            {
                FromDate = excelCalendarSheet.Item3.Cells[startRow, fromDateColumn].Value,
                ToDate = excelCalendarSheet.Item3.Cells[startRow, toDateColumn].Value,
                HoursPerDay = excelCalendarSheet.Item3.Cells[startRow, hoursPerDayColumn].Value,
                Account = excelCalendarSheet.Item3.Cells[startRow, accountColumn].Value,
                ProjectCode = excelCalendarSheet.Item3.Cells[startRow, projectCodeColumn].Value.ToString().ToLower().Trim(),
                Row = startRow
            });
            startRow++;
        }
        catch (Exception ex)
        {
            break;
        }
    }
    while (excelCalendarSheet.Item3.Cells[startRow, fromDateColumn] != null
         && excelCalendarSheet.Item3.Cells[startRow, toDateColumn] != null
         && excelCalendarSheet.Item3.Cells[startRow, hoursPerDayColumn] != null);
    return lstRangeModel;
}
#endregion

var lstOT = GetlstOTModel(accountColumnInOTSheet, OTSummaryColumn, MonthColumn, startRowOTSheet);
var lstLeaveDays = GetLstLeaveModels(startRowLeaveSheet, accountColumnLeaveSheet, sumDaysColumn, leaveFromColumn, leaveToColumn, leaveColumn, leaveTypeColumn);
var lstRangeModel = GetActualWorkingDays(startRow, fromDateColumn, toDateColumn, hoursPerDayColumn, accountColumn, projectCodeColumn);
var workingsDays = WorkingDaysInDuration(new DateTime(2024, 1, 1), new DateTime(2024, 12, 31));

var lastrow = 0;
Excel.Range last = excelCalendarSheet.Item2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
Excel.Range range = excelCalendarSheet.Item2.get_Range("A1", last);
lastrow = last.Row;
Console.WriteLine();

#region Khởi tạo 5 mảng có số cột là số tháng và số hàng là
int accountColumnInArray = 12;
int startRowInExcel = 3;
object[,] expectedWorkingdaysArray = new object[lastrow - startRowInExcel + 1, 13];
object[,] actualWorkingdaysArray = new object[lastrow - startRowInExcel + 1, 13];
object[,] leaveDaysArray = new object[lastrow - startRowInExcel + 1, 13];
object[,] OtDaysArray = new object[lastrow - startRowInExcel + 1, 13];
object[,] manMonthArray = new object[lastrow - startRowInExcel + 1, 13];
#endregion

#region mảng expectedWorkings days, tính số giờ làm việc dự kiến
//lstResults = WorkingDaysInDuration(startDate, endDate);
for (int i = 0; i < expectedWorkingdaysArray.GetLength(0); i++)
{
    foreach (var workingday in workingsDays)
    {
        expectedWorkingdaysArray[i, workingday.Key.Item2 - 1] = workingday.Value.Item2;
    }
    expectedWorkingdaysArray[i, accountColumnInArray] = lstRangeModel.Where(model => (model.Row == i + 3)).FirstOrDefault().Account;
}
#endregion 

#region mảng OtDaysArray, tính số giờ làm việc OT
for (int i = 0; i < OtDaysArray.GetLength(0); i++)
{
    OtDaysArray[i, accountColumnInArray] = lstRangeModel.Where(model => (model.Row == i + 3)).FirstOrDefault().Account;
    if (isDuplicated(i, OtDaysArray, OtDaysArray[i, accountColumnInArray].ToString())) continue ;
 
    var lstotPersonal = lstOT.Where(ot => ot.Account.ToLower().Equals(OtDaysArray[i, accountColumnInArray].ToString().ToLower())).ToList();
    if (lstotPersonal != null)
    {
        foreach(var otPerson in lstotPersonal)
        {
            if (OtDaysArray[i, otPerson.Month - 1] is null)
            {
                OtDaysArray[i, otPerson.Month - 1] = otPerson.OverTimeHoursSummary;
            }
            else
            {
                double otHours = double.Parse(OtDaysArray[i, otPerson.Month - 1].ToString());
                otHours += otPerson.OverTimeHoursSummary;
                OtDaysArray[i, otPerson.Month - 1] = otHours;

            }

        }
        //OtDaysArray[i, lstotPersonal.Month - 1] = sumHoursByAccount.SumHours;
    }

}
#endregion
Console.WriteLine();

#region Tính số ngày nghỉ thực tế
for (int i = 0; i < leaveDaysArray.GetLength(0); i ++)
{
    leaveDaysArray[i, accountColumnInArray] = lstRangeModel.Where(model => (model.Row == i + 3)).FirstOrDefault().Account;
    if (isDuplicated(i, OtDaysArray, OtDaysArray[i, accountColumnInArray].ToString())) continue;


}
#endregion

//tính giờ thực tế
#region Tính số ngày đi làm thực tế
for (int i = 0; i < actualWorkingdaysArray.GetLength(0); i++)
{
    actualWorkingdaysArray[i, accountColumnInArray] = lstRangeModel.Where(model => (model.Row == i + 3)).FirstOrDefault().Account;

}    
foreach (var rangeModel in lstRangeModel)
{

    var lstDuration = WorkingDaysInDuration(rangeModel.FromDate, rangeModel.ToDate);
    foreach (var duration in lstDuration)
    {
        actualWorkingdaysArray[rangeModel.Row - 3, duration.Key.Item2 - 1] = duration.Value.Item1 * rangeModel.HoursPerDay;
    }
}

for (int i = 0; i < actualWorkingdaysArray.GetLength(0); i ++)
{
    for (int j = 0;j < actualWorkingdaysArray.GetLength(1); j ++)
    {
        Console.Write(actualWorkingdaysArray[i, j] + " ");
    }
    Console.Write("\n");
}
#endregion
#region test array
//var excelTemp = GetOpeningExcelFile("Simulate", "TestCase");
//Excel.Range rangeTest = excelTemp.Item2.Range["A1"].Resize[10000, 10000];
//range.Value = actualWorkingdaysArray;
//excelTemp.Item1.Saved = true;
//Console.WriteLine("");
#endregion

#region Get Dictionary of TMS
// Get distinct values by the Name property

var distinctList = lstLeaveDays
                    .Select(p => new {Account = p.Account.ToLower() })
                    .Distinct().ToList();
object[,] tempLeaveArray = new object[distinctList.Count(), 13];
for (int i = 0;  i < tempLeaveArray.GetLength(0); i ++)
{
    tempLeaveArray[i, accountColumnInArray] = distinctList[i].Account;
}

foreach(var leaveDay in lstLeaveDays)
{
    int row = GetRowWithProvidedAccount(tempLeaveArray, leaveDay.Account);
    var lstLeaveDuration = LeaveDaysInDuration(leaveDay.LeaveFrom, leaveDay.LeaveTo);
    if (row == -10) continue;
    foreach(var leaveDuration in lstLeaveDuration)
    {
        double leaveCount = 0.0;
        if (leaveDay.LeaveType == PartialDayLeave)
        {
            leaveCount = (double)leaveDuration.Value * 0.5 * 8;
        }
        else leaveCount = leaveDuration.Value * 8;

        if (tempLeaveArray[row, leaveDuration.Key.Item2 - 1] is null) tempLeaveArray[row, leaveDuration.Key.Item2 - 1] = leaveCount;
        else
        {
            var currentLeaveCount = (double)tempLeaveArray[row, leaveDuration.Key.Item2 - 1];
            currentLeaveCount += leaveCount;
            tempLeaveArray[row, leaveDuration.Key.Item2 - 1] = currentLeaveCount;
        }
    }
}
int GetRowWithProvidedAccount(object[,] tempLeaveArray, string account)
{
    for (int i = 0; i < tempLeaveArray.GetLength(0); i++)
    {
        if (tempLeaveArray[i, accountColumnInArray] is null) continue;
        if (tempLeaveArray[i, accountColumnInArray].ToString().ToLower() == account.ToLower())
        {
            return i;
        }
    }
    return -10;
}
//
Console.WriteLine();
#endregion


for (int i = 0; i < manMonthArray.GetLength(0); i++)
{
    manMonthArray[i, accountColumnInArray] = lstRangeModel.Where(model => (model.Row == i + 3)).FirstOrDefault().Account;

    for (int j = 0; j < manMonthArray.GetLength(1)-1; j++)
    {
        double actual_time = 0.0;
        double expectedTime = 0.0;
        double OT_time = 0.0;
        try
        {
            actual_time = double.Parse(actualWorkingdaysArray[i, j].ToString());
        }
        catch
        {
            actual_time = 0.0;
        }

        try
        {
            expectedTime = double.Parse(expectedWorkingdaysArray[i, j].ToString());
        }
        catch
        {
            expectedTime = 0.0;
        }

        try
        {
            OT_time = double.Parse(OtDaysArray[i, j].ToString());
        }
        catch
        {
            OT_time = 0.0;
        }
        //double OT_time = arr_OT[i, j];
        string IDName = manMonthArray[i, accountColumnInArray].ToString();
        double total_WorkinTime = actual_time + OT_time;
        double TSM_time = Get_TMS_Value(ref tempLeaveArray, IDName, j, total_WorkinTime);

        manMonthArray[i, j] = Math.Round((total_WorkinTime - TSM_time) / expectedTime, 2);
    }
}

double Get_TMS_Value(ref object[,] tempLeaveArray, string IDName, int month, double total_WorkingTime)
{
    double Value = 0.0;

    for (int i = 0; i < tempLeaveArray.GetLength(0); i++)
    {
        if (tempLeaveArray[i, tempLeaveArray.GetLength(1)-1].ToString().ToLower() == IDName.ToLower())
        {
            if (tempLeaveArray[i, month] is null) tempLeaveArray[i, month] = 0.0;
            Value = double.Parse(tempLeaveArray[i, month].ToString());

            if (Value < total_WorkingTime)
            {
                tempLeaveArray[i, month] = 0;
                return Value;
            }
            if (Value > total_WorkingTime)
            {
                //tempLeaveArray[i, month] = Get_TMS_Value(tempLeaveArray, IDName, month, Value - total_WorkingTime);
                //return total_WorkingTime;
                double remainingTime = Value - total_WorkingTime;
                tempLeaveArray[i, month] = remainingTime;
                return total_WorkingTime;

            }
        }
    }

    return Value; // Return Value even if no conditions are met
}


bool isDuplicated(int row, object[,] array2D, string account)
{
    for (int i = 0; i < row; i ++)
    {
        if (array2D[i, accountColumnInArray] is null) continue;
        if (array2D[i, accountColumnInArray].ToString().ToLower().Equals(account))
        {
            return true;
        }
    }
    return false;
}

Marshal.ReleaseComObject(excelOT.Item2);
Marshal.ReleaseComObject(excelTMS.Item2);
Marshal.ReleaseComObject(excelCalendarSheet.Item2);

//Điền 
var excelTemp = GetOpeningExcelFile("Simulate", "TestCase");
Excel.Range rangeTest = excelTemp.Item2.Range["A1"].Resize[lastrow, 13];
rangeTest.Value = manMonthArray;
excelTemp.Item1.Saved = true;
Console.WriteLine();































//for (int j = 0; j < workingsDays.Count; j ++)
//{
//    expectedWorkingdaysArray[i, j] = workingsDays[i]
//}

//var otPersonal = lstOT.Where(ot => ot.Account.ToLower().Equals(OtDaysArray[i, accountColumnInArray].ToString().ToLower())).FirstOrDefault();
//if (otPersonal != null && OtDaysArray[i, otPersonal.Month - 1] == null)
//{
//    var sumHoursByAccount = lstOT.Where(m => m.Account.ToString().ToLower().Equals(OtDaysArray[i, accountColumnInArray].ToString().ToLower()))
//                        .GroupBy(m => new { m.Account, m.Month })
//                        .Select(g => new { Account = g.Key, SumHours = g.Sum(m => m.OverTimeHoursSummary) }).FirstOrDefault();
//    OtDaysArray[i, otPersonal.Month - 1] = sumHoursByAccount.SumHours;
//}