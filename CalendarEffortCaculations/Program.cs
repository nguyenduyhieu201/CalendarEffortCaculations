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
using System.Diagnostics;
using System;

Console.WriteLine("Hello, World!");

var watchAll = Stopwatch.StartNew();
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
    var workWatch = Stopwatch.StartNew();

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
    workWatch.Stop();
    Console.WriteLine($"tinh thoi gian nghi tu ngay {startDate.ToString()} den ngay {endDate.ToString()} het {workWatch.ElapsedMilliseconds}");

    return workingDaysPerMonth;
}
//tính số ngày nghỉ trong tháng
static int LeaveDaysInMonth(int year, int month, DateTime startDate, DateTime endDate)
{
    var leaveDayswatch = Stopwatch.StartNew();
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
    leaveDayswatch.Stop();
    Console.WriteLine($"leaveDayswatch la {leaveDayswatch.ElapsedMilliseconds}");
    return leaveDays;
}
//
static Dictionary<Tuple<int, int>, int> LeaveDaysInDuration(DateTime startDate, DateTime endDate)
{
    var leaveWatch = Stopwatch.StartNew();
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
    leaveWatch.Stop();
    Console.WriteLine($"tinh thoi gian nghi tu ngay {startDate.ToString()} den ngay {endDate.ToString()} het {leaveWatch.ElapsedMilliseconds}");
    return leaveDaysPerMonth;
}

#region Lấy số ngày OT theo tháng
var watch = Stopwatch.StartNew();
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
watch.Stop();
Console.WriteLine($"thoi gian chay la {watch.ElapsedMilliseconds}");
#endregion

#region Lấy số ngày nghỉ theo tháng
var watchLeave = Stopwatch.StartNew();
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

watchLeave.Stop();
Console.WriteLine($"thoi gian chay la {watchLeave.ElapsedMilliseconds}");
#endregion

#region Lấy số ngày làm việc thực tế
var watchActualWorking = Stopwatch.StartNew();

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

watchActualWorking.Stop();
Console.WriteLine($"watchActualWorking la {watchActualWorking.ElapsedMilliseconds}");

#endregion

var lstOT = GetlstOTModel(accountColumnInOTSheet, OTSummaryColumn, MonthColumn, startRowOTSheet);
var lstLeaveDays = GetLstLeaveModels(startRowLeaveSheet, accountColumnLeaveSheet, sumDaysColumn, leaveFromColumn, leaveToColumn, leaveColumn, leaveTypeColumn);
var lstRangeModel = GetActualWorkingDays(startRow, fromDateColumn, toDateColumn, hoursPerDayColumn, accountColumn, projectCodeColumn);
var workingsDays = WorkingDaysInDuration(new DateTime(2024, 1, 1), new DateTime(2024, 12, 31));

var lastrow = 0;
Excel.Range last = excelCalendarSheet.Item2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
Excel.Range range = excelCalendarSheet.Item2.get_Range("A1", last);
lastrow = last.Row;
var lastColumn = 0;
lastColumn = last.Column;

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
    try
    {
        expectedWorkingdaysArray[i, accountColumnInArray] = lstRangeModel.Where(model => (model.Row == i + 3)).FirstOrDefault().Account;
    }
    catch { }
}
#endregion 

#region mảng OtDaysArray, tính số giờ làm việc OT
var watchOTWorkingdays = Stopwatch.StartNew();

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
    }

}

watchOTWorkingdays.Stop();
Console.WriteLine($"watchOTWorkingdays la {watchOTWorkingdays.ElapsedMilliseconds}");

#endregion
Console.WriteLine();

#region Tính số ngày nghỉ thực tế
//for (int i = 0; i < leaveDaysArray.GetLength(0); i ++)
//{
//    leaveDaysArray[i, accountColumnInArray] = lstRangeModel.Where(model => (model.Row == i + 3)).FirstOrDefault().Account;
//    if (isDuplicated(i, OtDaysArray, OtDaysArray[i, accountColumnInArray].ToString())) continue;


//}
#endregion

//tính giờ thực tế
#region Tính số ngày đi làm thực tế
var watchActualWorkingDays = Stopwatch.StartNew();
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
watchActualWorkingDays.Stop();
Console.WriteLine($"watchActualWorkingDays la {watchActualWorkingDays.ElapsedMilliseconds}");

//for (int i = 0; i < actualWorkingdaysArray.GetLength(0); i ++)
//{
//    for (int j = 0;j < actualWorkingdaysArray.GetLength(1); j ++)
//    {
//        Console.Write(actualWorkingdaysArray[i, j] + " ");
//    }
//    Console.Write("\n");
//}
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
var getLeaveWatch = Stopwatch.StartNew();
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

getLeaveWatch.Stop();
Console.WriteLine($"get leave watch la {getLeaveWatch.ElapsedMilliseconds}");

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

//var manmonthWatch = Stopwatch.StartNew();
//for (int i = 0; i < manMonthArray.GetLength(0); i++)
//{
//    manMonthArray[i, accountColumnInArray] = lstRangeModel.Where(model => (model.Row == i + 3)).FirstOrDefault().Account;

//    for (int j = 0; j < manMonthArray.GetLength(1)-1; j++)
//    {
//        double actual_time = 0.0;
//        double expectedTime = 0.0;
//        double OT_time = 0.0;
//        try
//        {
//            actual_time = double.Parse(actualWorkingdaysArray[i, j].ToString());
//        }
//        catch
//        {
//            actual_time = 0.0;
//        }

//        try
//        {
//            expectedTime = double.Parse(expectedWorkingdaysArray[i, j].ToString());
//        }
//        catch
//        {
//            expectedTime = 0.0;
//        }

//        try
//        {
//            OT_time = double.Parse(OtDaysArray[i, j].ToString());
//        }
//        catch
//        {
//            OT_time = 0.0;
//        }
//        //double OT_time = arr_OT[i, j];
//        string IDName = manMonthArray[i, accountColumnInArray].ToString();
//        double total_WorkinTime = actual_time + OT_time;
//        double TSM_time = Get_TMS_Value(ref tempLeaveArray, IDName, j, total_WorkinTime);

//        manMonthArray[i, j] = Math.Round((total_WorkinTime - TSM_time) / expectedTime, 2);
//    }
//}
//manmonthWatch.Stop();
//Console.WriteLine($"man month la {manmonthWatch.ElapsedMilliseconds}");
//double Get_TMS_Value(ref object[,] tempLeaveArray, string IDName, int month, double total_WorkingTime)
//{
//    double Value = 0.0;

//    for (int i = 0; i < tempLeaveArray.GetLength(0); i++)
//    {
//        if (tempLeaveArray[i, tempLeaveArray.GetLength(1)-1].ToString().ToLower() == IDName.ToLower())
//        {
//            if (tempLeaveArray[i, month] is null) tempLeaveArray[i, month] = 0.0;
//            Value = double.Parse(tempLeaveArray[i, month].ToString());

//            if (Value < total_WorkingTime)
//            {
//                tempLeaveArray[i, month] = 0;
//                return Value;
//            }
//            if (Value > total_WorkingTime)
//            {
//                //tempLeaveArray[i, month] = Get_TMS_Value(tempLeaveArray, IDName, month, Value - total_WorkingTime);
//                //return total_WorkingTime;
//                double remainingTime = Value - total_WorkingTime;
//                tempLeaveArray[i, month] = remainingTime;
//                return total_WorkingTime;

//            }
//        }
//    }

//    return Value; // Return Value even if no conditions are met
//}

var manmonthWatch = Stopwatch.StartNew();

#region caculate man month updated
double leaveValue = 0.0;
Dictionary<Tuple<string, int>, double> leaveValues = new Dictionary<Tuple<string, int>, double>();
for (int i = 0; i < tempLeaveArray.GetLength(0); i++)
{
    string leaveID = tempLeaveArray[i, tempLeaveArray.GetLength(1) - 1]?.ToString().ToLower();
    
    for (int month = 0; month < tempLeaveArray.GetLength(1) - 1; month++)
    {
        if (double.TryParse(tempLeaveArray[i, month]?.ToString(), out leaveValue))
        {
            leaveValues[Tuple.Create(leaveID, month)] = leaveValue;
            //leaveValues[leaveID] = leaveValue;
        }
    }
}

for (int i = 0; i < manMonthArray.GetLength(0); i++)
{
    var account = lstRangeModel.FirstOrDefault(model => model.Row == i + 3)?.Account;
    manMonthArray[i, accountColumnInArray] = account?.ToString();

    for (int j = 0; j < manMonthArray.GetLength(1) - 1; j++)
    {
        double actual_time;
        double expectedTime;
        double OT_time;

        double.TryParse(actualWorkingdaysArray[i, j]?.ToString(), out actual_time);
        double.TryParse(expectedWorkingdaysArray[i, j]?.ToString(), out expectedTime);
        double.TryParse(OtDaysArray[i, j]?.ToString(), out OT_time);

        string IDName = manMonthArray[i, accountColumnInArray]?.ToString();
        double total_WorkinTime = actual_time + OT_time;
        //double TSM_time = Get_TMS_Updated_Value(ref tempLeaveArray, IDName, j, total_WorkinTime, leaveValues);
        var TSM_time = Get_TMS_Updated_Value(ref tempLeaveArray, IDName, j, total_WorkinTime, leaveValues);
        if (i == 229)
        {
            Console.WriteLine($"dong thu {i} voi thang {j + 1} co gia tri tms  la {TSM_time.Item1} ;total_WorkinTime la {total_WorkinTime} ");
        }
        manMonthArray[i, j] = Math.Round((total_WorkinTime - TSM_time.Item1) / expectedTime, 2);
    }
}

leaveValue = 0.0;
//Dictionary<string, double> leaveValues = new Dictionary<string, double>();


double Get_TMS_Value(ref object[,] tempLeaveArray, string IDName, int month, double total_WorkingTime)
{
    Dictionary<string, double> leaveValues = new Dictionary<string, double>();

    for (int i = 0; i < tempLeaveArray.GetLength(0); i++)
    {
        string leaveID = tempLeaveArray[i, tempLeaveArray.GetLength(1) - 1]?.ToString().ToLower();
        if (double.TryParse(tempLeaveArray[i, month]?.ToString(), out leaveValue))
        {
            leaveValues[leaveID] = leaveValue;
        }
    }

    if (leaveValues.TryGetValue(IDName?.ToLower(), out leaveValue))
    {
        if (leaveValue < total_WorkingTime)
        {
            leaveValues[IDName?.ToLower()] = 0;
            return leaveValue;
        }
        if (leaveValue > total_WorkingTime)
        {
            double remainingTime = leaveValue - total_WorkingTime;
            leaveValues[IDName?.ToLower()] = remainingTime;
            return total_WorkingTime;
        }
    }

    return 0.0; // Return 0 if no matching condition is met
}


Tuple<double,Dictionary<Tuple<string, int>, double>> Get_TMS_Updated_Value(ref object[,] tempLeaveArray, string IDName, int month, double total_WorkingTime, Dictionary<Tuple<string, int>, double> leaveValues)
{
    Tuple<string, int> leaveKey = Tuple.Create(IDName, month);
    //for (int i = 0; i < tempLeaveArray.GetLength(0); i++)
    //{
    //    string leaveID = tempLeaveArray[i, tempLeaveArray.GetLength(1) - 1]?.ToString().ToLower();
    //    if (double.TryParse(tempLeaveArray[i, month]?.ToString(), out leaveValue))
    //    {
    //        leaveValues[Tuple.Create(leaveID, month)] = leaveValue;
    //        //leaveValues[leaveID] = leaveValue;
    //    }
    //}

    if (leaveValues.TryGetValue(Tuple.Create(IDName?.ToLower(), month), out leaveValue))
    {
        if (leaveValue <= total_WorkingTime)
        {
            leaveValues[Tuple.Create(IDName?.ToLower(), month)] = 0;
            return Tuple.Create(leaveValue, leaveValues);
        }
        if (leaveValue > total_WorkingTime)
        {
            double remainingTime = leaveValue - total_WorkingTime;
            leaveValues[Tuple.Create(IDName?.ToLower(), month)] = remainingTime;
            return Tuple.Create(total_WorkingTime, leaveValues);
        }
    }

    return Tuple.Create(0.0, leaveValues); // Return 0 if no matching condition is met
}
manmonthWatch.Stop();
Console.WriteLine($"man month la {manmonthWatch.ElapsedMilliseconds}");
#endregion


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
//Marshal.ReleaseComObject(excelCalendarSheet.Item2);

//Điền 
//var excelTemp = GetOpeningExcelFile("Simulate", "TestCase");

#region paste value in Calendar sheet
int startPasteColumn = 0;
int startPasteRow = 0;
var currentExcel = excelCalendarSheet.Item2.Range["A1"].Resize[lastrow, 30];
var currentExcelValue = currentExcel.Value;
for(int i = 1; i <= currentExcelValue.GetLength(0); i ++)
{
    if (currentExcelValue[1, i] is null) continue;
    if (currentExcelValue[1, i].ToString() == "1")
    {
        startPasteColumn = i-1;
        break;
    }
}

for (int i = 1; i <= currentExcelValue.GetLength(0); i ++)
{
    for (int j = 1; j <= currentExcelValue.GetLength(1); j ++)
    {
        if (currentExcelValue[i, j] is null) continue;
        if (currentExcelValue[i, j].ToString() == "Job")
        {
            startPasteRow = i + 1;
            break;
        }

    }
}
Excel.Range pastedRange = excelCalendarSheet.Item2.Cells[startPasteRow, startPasteColumn].Resize[lastrow - startRow, 12];
pastedRange.Value = manMonthArray;
excelCalendarSheet.Item1.Saved = true;
#endregion

#region Tính số ngày abnormal case
//var excelAbnormalSheet = 
var numColumnInAbNormalSheet = 3;
var excelAbnormalSheet = GetOpeningExcelFile("Simulate", "AbnormalCase");
currentExcel = excelAbnormalSheet.Item2.Range["A1"].Resize[lastrow, 30];
currentExcelValue = currentExcel.Value;
for (int i = 1; i <= currentExcelValue.GetLength(0); i++)
{
    for (int j = 1; j <= currentExcelValue.GetLength(1); j++) {
        if (currentExcelValue[i, j] is null) continue;
        if (currentExcelValue[i, j].ToString() == "Account")
        {
            startPasteColumn = j;
            startPasteRow = i+1;
            break;
        }
    }
}



var lstAbNormal = leaveValues.Where(leaveValue => leaveValue.Value != 0).ToList();
object[,] abnormalArray = new object[lstAbNormal.Count, 3];
int startRowInAbnormal = 0;
foreach(var abnormal in lstAbNormal)
{
    abnormalArray[startRowInAbnormal, 0] = abnormal.Key.Item1;
    //them thang voi 1 
    abnormalArray[startRowInAbnormal, 1] = abnormal.Key.Item2 + 1;
    abnormalArray[startRowInAbnormal, 2] = abnormal.Value;
    startRowInAbnormal++;
}
Excel.Range pastedAbnormalRange = excelAbnormalSheet.Item2.Cells[startPasteRow, startPasteColumn]
                                                    .Resize[abnormalArray.GetLength(0), abnormalArray.GetLength(1)];
pastedAbnormalRange.Value = abnormalArray;
excelAbnormalSheet.Item1.Saved = true;
watchAll.Stop();
Console.WriteLine("thoi gian chay " + watchAll.ElapsedMilliseconds + "ms");
#endregion
//Console.WriteLine();































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