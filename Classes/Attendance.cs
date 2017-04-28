using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EnrolmentTool.Classes
{
    class Attendance
    {
        public String StudentNo { get; set; }
        public String Name { get; set; }
        public String Visa { get; set; }
        public String Warning { get; set; }
        public String NewWarning { get; set; }
        public String CourseCode { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public double CurrentAttendace { get; set; }
        public double OverallAttendace { get; set; }
        public const String FIRST = "Attendance - first warning";
        public const String SECOND = "Attendance - second warning";
        public const String THIRD = "Attendance - third warning";
        public const String COUNSEL = "Attendance - counsel";
        public const String INTENT = "Attendance - intent to report";
    }
}
