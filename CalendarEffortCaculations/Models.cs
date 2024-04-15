using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalendarEffortCaculations
{
    public class RangeTimeModel
    {
        public DateTime FromDate { set; get; }
        public DateTime ToDate { set; get; }
        public int Row { set; get; }
        public double HoursPerDay { set; get; }
        public string Account { set; get; }
        public string ProjectCode { set; get; }
    }

    public class Models
    {
        public string Account { set; get; }
        public double OverTimeHoursSummary { set; get; }
        public int Month { set; get; }
    }

    public class PersonalLeaveDay
    {
        public string Account { set; get; }
        public double LeaveHoursSummary { set; get; }
        public int Month { set; get; }
        public DateTime LeaveFrom {set;get;}
        public DateTime LeaveTo { set; get; }
        public double SumsLeaveDays { set; get; }
        public int LeaveType { set; get; }
        public double SumLeaveHours { set; get; }
    }
}
