using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EnrolmentTool.Classes
{
    class Starters
    {
        public String StudentNo { get; set; }
        public String Name { get; set; }
        public String Visa { get; set; }
        public String CourseCode { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public String Agent { get; set; }
        public DateTime CoEStartDate { get; set; }
        public DateTime CoEEndDate { get; set; }
        public String CoEDescription { get; set; }
        public String CoEStatus { get; set; }
        public int CoEDateChecker { get; set; }
        public Starters()
        {

        }
        public bool CheckStartDate()
        {
            if (this.StartDate == this.CoEStartDate)
                return true;
            else return false;
        }
        public bool CheckEndDate(){
            if (this.EndDate == this.CoEStartDate)
                return true;
            else return false;
        }
        public void CheckCourseCode()
        {
            if (this.Visa == "Student") { 
            CoEDateChecker = 0;
            bool aaa = true;
            if (this.CourseCode.Contains("GE") || this.CourseCode.Contains("GRD") || this.CourseCode.Contains("GRN")||this.CourseCode.Contains("GRND"))
            {
                this.CourseCode = "GE";
            }

            if (this.CourseCode.Contains("IELTS"))
            {
                this.CourseCode = "IELTS";
            }
            if(CoEDescription == null || !CourseCode.Contains(CoEDescription)){
                aaa = false;
                CoEDateChecker += 4;
            }
            

            if (this.StartDate != this.CoEStartDate)
            {
                aaa = false;
                CoEDateChecker += 1;                
            }
            if (this.EndDate != this.CoEEndDate)
            {
                CoEDateChecker += 2;
                aaa = false;
            }



            if (aaa)
            {
                this.CoEStatus = "Okay";
            }
            else this.CoEStatus = "!!!";
            }
        }
         
    }
    
    
}
