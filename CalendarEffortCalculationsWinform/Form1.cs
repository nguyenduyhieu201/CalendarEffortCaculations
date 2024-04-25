using OfficeOpenXml;
using System.ComponentModel;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;
using System.IO.Packaging;
using CalendarEffortCalculationsWinform.BusinessLogic;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using Tuple = System.Tuple;
using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeOpenXml.Style;

namespace CalendarEffortCalculationsWinform
{
    public partial class Form1 : Form
    {
        string filePath = "";

        ExcelPackage Ep = null;
        ExcelWorkbook simulateWorkbook = null;
        ExcelWorksheet calendarSheet = null;
        ExcelWorksheet tmsSheet = null;
        ExcelWorksheet otSheet = null;
        ExcelWorksheet abnormalSheet = null;
        int accountColumnInArray = 12;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void OK_Click(object sender, EventArgs e)
        {
            openFilePath(filePath);
            try
            {
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Please choose a file");
                return;
            }
            if (!IsExcelFile(filePath))
            {
                MessageBox.Show("Incorrect Format");
                return;
            }
            openFilePath(filePath);
            int accountColumn = 0;
            var excelBusinessLogic = new ExcelService();
            var workTimes = excelBusinessLogic.GetActualWork(calendarSheet, out accountColumn);
            var leaveTimes = excelBusinessLogic.GetLeave(tmsSheet);
            var otTimes = excelBusinessLogic.GetOTTimes(otSheet);
            var expecteWorkTimes = WorkingDaysService.WorkingDaysInDuration(new DateTime(2024, 1, 1), new DateTime(2024, 12, 31));

            #region lấy start row 
            int startRowInExcel = 0;
            int startPastedColumnInExcel = 0;
            ExcelRangeBase titleRange = calendarSheet.Cells[1, 1, 4, 80];
            object[,] titleExcelValue = titleRange.Value as object[,];
            for (int i = 0; i < titleExcelValue.GetLength(0); i++)
            {
                for (int j = 0; j < titleExcelValue.GetLength(1); j++)
                {
                    if (titleExcelValue[i, j] is null) continue;
                    if (titleExcelValue[i, j].ToString().ToLower() == "project code")
                    {
                        startRowInExcel = i + 3;
                    }
                    if (titleExcelValue[i, j].ToString().ToLower() == "1")
                    {
                        startPastedColumnInExcel = j;
                    }

                }
            }
            #endregion

            #region Khởi tạo 5 mảng có số cột là số tháng và số hàng là
            int accountColumnInArray = 12;
            int lastrow = calendarSheet.Dimension.End.Row;
            if (calendarSheet.Cells[lastrow, accountColumn].Value is null)
            {
                while (calendarSheet.Cells[lastrow, accountColumn].Value is null) lastrow--;
            }
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
                foreach (var workingday in expecteWorkTimes)
                {
                    expectedWorkingdaysArray[i, workingday.Key.Item2 - 1] = workingday.Value.Item2;
                }
                try
                {
                    expectedWorkingdaysArray[i, accountColumnInArray] = workTimes.Where(model => (model.Row == i + startRowInExcel)).FirstOrDefault().Account;
                }
                catch { }
            }
            #endregion

            #region mảng OT Working days, tính số giờ làm việc OT

            for (int i = 0; i < OtDaysArray.GetLength(0); i++)
            {
                OtDaysArray[i, accountColumnInArray] = workTimes.Where(model => (model.Row == i + startRowInExcel)).FirstOrDefault().Account;
                if (isDuplicated(i, OtDaysArray, OtDaysArray[i, accountColumnInArray].ToString())) continue;

                var lstotPersonal = otTimes.Where(ot => ot.Account.ToLower().Equals(OtDaysArray[i, accountColumnInArray].ToString().ToLower())).ToList();
                if (lstotPersonal != null)
                {
                    foreach (var otPerson in lstotPersonal)
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
            #endregion

            #region mảng actual working days, tính số giờ làm việc thực tế
            var watchActualWorkingDays = Stopwatch.StartNew();
            for (int i = 0; i < actualWorkingdaysArray.GetLength(0); i++)
            {
                actualWorkingdaysArray[i, accountColumnInArray] = workTimes.Where(model => (model.Row == i + startRowInExcel)).FirstOrDefault().Account;

            }
            foreach (var rangeModel in workTimes)
            {

                var lstDuration = WorkingDaysService.WorkingDaysInDuration(rangeModel.FromDate, rangeModel.ToDate);
                foreach (var duration in lstDuration)
                {
                    actualWorkingdaysArray[rangeModel.Row - startRowInExcel, duration.Key.Item2 - 1] = duration.Value.Item1 * rangeModel.HoursPerDay;
                }
            }
            #endregion

            #region Get Dictionary of TMS
            // Get distinct values by the Name property
            var getLeaveWatch = Stopwatch.StartNew();
            var distinctList = leaveTimes
                                .Select(p => new { Account = p.Account.ToLower() })
                                .Distinct().ToList();
            object[,] tempLeaveArray = new object[distinctList.Count(), 13];
            for (int i = 0; i < tempLeaveArray.GetLength(0); i++)
            {
                tempLeaveArray[i, accountColumnInArray] = distinctList[i].Account;
            }


            foreach (var leaveDay in leaveTimes)
            {
                int row = GetRowWithProvidedAccount(tempLeaveArray, leaveDay.Account);
                var lstLeaveDuration = WorkingDaysService.WorkingDaysInDuration(leaveDay.LeaveFrom, leaveDay.LeaveTo);
                if (row == -10) continue;
                foreach (var leaveDuration in lstLeaveDuration)
                {
                    double leaveCount = 0.0;
                    if (leaveDay.LeaveType == Constants.PartialDayLeave)
                    {
                        leaveCount = (double)leaveDuration.Value.Item2 * 0.5;
                    }
                    else leaveCount = leaveDuration.Value.Item2;

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
                var account = workTimes.FirstOrDefault(model => model.Row == i + startRowInExcel)?.Account;
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


            Tuple<double, Dictionary<Tuple<string, int>, double>> Get_TMS_Updated_Value(ref object[,] tempLeaveArray, string IDName, int month, double total_WorkingTime, Dictionary<Tuple<string, int>, double> leaveValues)
            {
                Tuple<string, int> leaveKey = Tuple.Create(IDName, month);
                    

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
                #endregion
            #region Paste man month into Excel
            #region Delete calendar row
            int lastCalendarRow = calendarSheet.Dimension.Rows;

            List<object[]> dataEmptyList = new List<object[]>();
            int columnCount = manMonthArray.GetLength(1) - 1;
            for (int row = 0; row < lastCalendarRow; row++)
            {
                object[] rowData = new object[columnCount];
                for (int column = 0; column < columnCount; column++)
                {
                    rowData[column] = "";
                }
                dataEmptyList.Add(rowData);
            }
            var startCell = calendarSheet.Cells[startRowInExcel, startPastedColumnInExcel];
            var endCell = startCell.Offset(lastCalendarRow - 1, manMonthArray.GetLength(1) - 2);
            var range = calendarSheet.Cells[startRowInExcel, startPastedColumnInExcel];
            range.LoadFromArrays(dataEmptyList);

            #endregion
            List<object[]> dataList = new List<object[]>();
            int rowCount = manMonthArray.GetLength(0);
            columnCount = manMonthArray.GetLength(1) - 1;
            for (int row = 0; row < rowCount; row++)
            {
                object[] rowData = new object[columnCount];
                for (int column = 0; column < columnCount; column++)
                {
                    rowData[column] = manMonthArray[row, column];
                }
                dataList.Add(rowData);
            }

            startCell = calendarSheet.Cells[startRowInExcel, startPastedColumnInExcel];
            endCell = startCell.Offset(manMonthArray.GetLength(0) - 1, manMonthArray.GetLength(1) - 2);
            range = calendarSheet.Cells[startRowInExcel, startPastedColumnInExcel];

            range.LoadFromArrays(dataList);
            #endregion

            #region Tính abnormal case
            ExcelRangeBase titleAbnormalRange = abnormalSheet.Cells[1, 1, 4, 80];

            object[,] titleAbnormalValue = titleAbnormalRange.Value as object[,];
            int abnormalStartRow = 0;
            int abnormalStartColumn = 0;
            for (int i = 0; i < titleAbnormalValue.GetLength(0); i++)
            {
                for (int j = 0; j < titleAbnormalValue.GetLength(1); j++)
                {
                    if (titleAbnormalValue[i, j] is null) continue;
                    if (titleAbnormalValue[i, j].ToString().ToLower() == "account")
                    {
                        abnormalStartRow = i + 2;
                        abnormalStartColumn = j + 1;
                    }
                }
            }

            var lstAbNormal = leaveValues.Where(leaveValue => leaveValue.Value != 0).ToList();
            object[,] abnormalArray = new object[lstAbNormal.Count, 3];
            int startRowInAbnormal = 0;
            foreach (var abnormal in lstAbNormal)
            {
                abnormalArray[startRowInAbnormal, 0] = abnormal.Key.Item1;
                //them thang voi 1 
                abnormalArray[startRowInAbnormal, 1] = abnormal.Key.Item2 + 1;
                abnormalArray[startRowInAbnormal, 2] = abnormal.Value;
                startRowInAbnormal++;
            }

            dataList = new List<object[]>();
            rowCount = abnormalArray.GetLength(0);
            columnCount = abnormalArray.GetLength(1);

            #region Empty các dữ liệu
            int lastAbnormalRow = abnormalSheet.Dimension.Rows;
            for (int row = 0; row < lastAbnormalRow; row++)
            {
                object[] rowData = new object[columnCount];
                for (int column = 0; column < columnCount; column++)
                {
                    rowData[column] = "";
                }
                dataList.Add(rowData);
            }
            var startAbnormalCell = abnormalSheet.Cells[abnormalStartRow, abnormalStartColumn];
            var endAbnormalCell = startCell.Offset(abnormalArray.GetLength(0) - 1, abnormalArray.GetLength(1) - 1);
            var abNormalRange = abnormalSheet.Cells[abnormalStartRow, abnormalStartColumn];
            abNormalRange.LoadFromArrays(dataList);
            #endregion

            dataList = new List<object[]>();
            for (int row = 0; row < rowCount; row++)
            {
                object[] rowData = new object[columnCount];
                for (int column = 0; column < columnCount; column++)
                {
                    rowData[column] = abnormalArray[row, column];
                }
                dataList.Add(rowData);
            }


            //var startAbnormalCell = abnormalSheet.Cells[abnormalStartRow, abnormalStartColumn];
            //var endAbnormalCell = startCell.Offset(abnormalArray.GetLength(0) - 1, abnormalArray.GetLength(1) - 1);
            //var abNormalRange = abnormalSheet.Cells[abnormalStartRow, abnormalStartColumn];

            abNormalRange.LoadFromArrays(dataList);
            Ep.Save();
            #endregion
            MessageBox.Show("Thành công");
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error: " + ex.ToString());
        }
        
        }

        private void filePath_DragDrop(object sender, DragEventArgs e)
        {
            string[] fieldList = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            filepath.Text = fieldList[0];
            filePath = fieldList[0];
        }

        private void firePath_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void BrowserButton_Click(object sender, EventArgs e)
        {
            int size = -1;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                try
                {
                    string text = File.ReadAllText(file);
                    size = text.Length;
                    filePath = file;
                    filepath.Text = file;
                }
                catch (IOException)
                {
                    MessageBox.Show("Có thể bạn đang mở file, vui lòng đóng và chọn lại file");
                }
            }
        }

        private bool IsExcelFile(string file)
        {
            var lstExcelExtension = new string[] { ".xlsx", ".xlsm", ".xls" };
            string extension = System.IO.Path.GetExtension(filePath);
            return lstExcelExtension.Contains(extension);
        }

        private void openFilePath(string file)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Ep = new ExcelPackage(file);

            calendarSheet = Ep.Workbook.Worksheets[Constants.CalendarSheetName];
            otSheet = Ep.Workbook.Worksheets[Constants.OTSheetName];
            tmsSheet = Ep.Workbook.Worksheets[Constants.TMSSheetName];
            abnormalSheet = Ep.Workbook.Worksheets[Constants.AbnormalSheetName];
            if (abnormalSheet == null)
            {
                abnormalSheet = Ep.Workbook.Worksheets.Add(Constants.AbnormalSheetName);
                abnormalSheet.Cells["A1"].Value = "Account";
                abnormalSheet.Cells["B1"].Value = "Month";
                abnormalSheet.Cells["C1"].Value = "Ab Hours";
                ExcelRange range = abnormalSheet.Cells[1, 1, 1, 3];
                // Set the cell background color
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);

                // Set the cell border
                range.Style.Font.Bold = true;
                range.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                range.Style.Border.Left.Style = ExcelBorderStyle.Medium;
                range.Style.Border.Right.Style = ExcelBorderStyle.Medium;
                //
                Ep.Save();

            }
            simulateWorkbook = Ep.Workbook;

        }

        bool isDuplicated(int row, object[,] array2D, string account)
        {
            for (int i = 0; i < row; i++)
            {
                if (array2D[i, accountColumnInArray] is null) continue;
                if (array2D[i, accountColumnInArray].ToString().ToLower().Equals(account))
                {
                    return true;
                }
            }
            return false;
        }

        private string GetCellReference(int row, int column)
        {
            char columnLetter = (char)('A' + column - 1);
            return $"{columnLetter}{row}";
        }

        private void filepath_TextChanged(object sender, EventArgs e)
        {

        }
    }
}


//for (int i = 0; i < tempLeaveArray.GetLength(0); i++)
//{
//    string leaveID = tempLeaveArray[i, tempLeaveArray.GetLength(1) - 1]?.ToString().ToLower();
//    if (double.TryParse(tempLeaveArray[i, month]?.ToString(), out leaveValue))
//    {
//        leaveValues[Tuple.Create(leaveID, month)] = leaveValue;
//        //leaveValues[leaveID] = leaveValue;
//    }
//}