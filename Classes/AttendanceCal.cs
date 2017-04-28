using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace EnrolmentTool.Classes
{
    class AttendanceCal
    {
           
        //private string sourceFile;
        //private string destFile;
        private Excel.Workbook srcBook;
        //private Excel.Workbook destBook;
        private Microsoft.Office.Interop.Excel.Application xlApp;

        private XmlReader fieldChooser= null;
        private List<Attendance> srcStudentList = null;
        private List<Attendance> resultStudentList = null;
        private String connectionStringSrc = null;
        
      //  private Loading LoadingForm ;
      //  private System.Timers.Timer timer1 = new System.Timers.Timer();

        public AttendanceCal(string src)
        {
            try
            {
                /* field chooser 
                 * Read data location from xml file and let the program know where the data is.
                */
                this.connectionStringSrc = src;

                this.fieldChooser = new XmlReader(XmlReader.ATTENDACE);
                
                this.xlApp = new Excel.Application();
                this.srcBook = this.xlApp.Workbooks.Open(src);
                //this.destBook = this.xlApp.Workbooks.Open(dest);
                
                /*
                 * in order to using LINQ, we connect to Excel file like a database.                  
                 */
                

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Can't Process Properly"+ex.ToString());
            }

        }
        public void Start()
        {
         //   this.connectionStringSrc = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Extended Properties=\"Excel 12.0;\"", this.connectionStringSrc);

            this.srcStudentList = this.CreateAttendanceList();
           
                       
            this.StartCheck();


            this.srcBook.Close(true, Type.Missing, Type.Missing);
            this.xlApp.Quit();
            if (this.resultStudentList.Count != 0)
            {
                this.MakeResults();
            }
        }

        private void MakeResults()
        {
            SavingExcel se = new SavingExcel(this.resultStudentList);
           
            se.WriteResultToExcelAttendance();
        }
       
   
       
     
      


        private void StartCheck()
        {
            DateTime current = DateTime.Today;
            int daysTillNextFriday = (int)current.DayOfWeek - (int)DayOfWeek.Friday;
            DateTime friday = current.AddDays(daysTillNextFriday).AddDays(7);
            
            this.resultStudentList = new List<Attendance>();
            foreach (Attendance student in this.srcStudentList)
            {
                if ((student.CurrentAttendace < 85 && student.EndDate > friday && student.CurrentAttendace != 0.0 ) || student.Warning == Attendance.INTENT )
                {

                    if (student.Warning == Attendance.COUNSEL)
                    {
                        student.NewWarning = Attendance.FIRST;
                    }
                    else if (student.Warning == Attendance.FIRST)
                    {
                        student.NewWarning = Attendance.SECOND;
                    }
                    else if (student.Warning == Attendance.SECOND)
                    {
                        student.NewWarning = Attendance.THIRD;
                    }
                    else if (student.Warning == Attendance.THIRD)
                    {
                        student.NewWarning = Attendance.INTENT;
                    }
                    else if (student.Warning == Attendance.INTENT)
                    {
                        student.NewWarning = Attendance.INTENT;
                    }
                    else
                        student.NewWarning = Attendance.COUNSEL;
                    this.resultStudentList.Add(student);
 
                }
            }

            
           // Console.WriteLine("done comp");
            
          
        }
       



        /*
         * Description this method create a list array for student object.
         * 1st student has to be added, counter is used for counting objects to know whether the student object is first or not.
         *
         * pre-cond : connection String.
         * 
         * returns a student object LIST
         */



        private List<Attendance> CreateAttendanceList()
        {
            
            // Creat OLE connection object

            //OleDbConnection con = new OleDbConnection(connectionStr);
            //con.Open();
            Excel.Worksheet ws1 = this.srcBook.Worksheets.get_Item(1);
            Excel.Range range = ws1.UsedRange;

            List<Attendance> studentList = new List<Attendance>();
            //Excel.Worksheet ws1 = this.srcBook.Open


            
            /*pull the data from excel file and store in dataset
             *pass to data to datatable collection.
            */

            //OleDbDataAdapter dtAdapter = new OleDbDataAdapter("Select * From [Sheet1$]", connectionStr);
            //var adapter = new OleDbDataAdapter("Select * From [Sheet1$]",con);
            //var ds = new DataSet();

            //adapter.Fill(ds, "student");
            //DataTable data = ds.Tables["student"];

            string tempValue = "";
            int rCnt = 0;
            // call row data and store data in student object.



            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {
                Attendance student = new Attendance();
                string value = (string)(range.Cells[rCnt, this.fieldChooser.StNoLoc] as Excel.Range).Value2;
                student.StudentNo = value;


                tempValue = (string)(range.Cells[rCnt, this.fieldChooser.StNameLoc] as Excel.Range).Value2;
                if (!String.IsNullOrEmpty(tempValue))
                    student.Name = tempValue;

                tempValue = (string)(range.Cells[rCnt, this.fieldChooser.StudentVisa] as Excel.Range).Value2;
                if (!String.IsNullOrEmpty(tempValue))
                    student.Visa = tempValue;

                tempValue = (string)(range.Cells[rCnt, this.fieldChooser.CourseNameLoc] as Excel.Range).Value2;
                if (!String.IsNullOrEmpty(tempValue))
                    student.CourseCode = tempValue;
                tempValue = (string)(range.Cells[rCnt, this.fieldChooser.WarnCat] as Excel.Range).Value2;
                if (!String.IsNullOrEmpty(tempValue))
                    student.Warning = tempValue;


                var bbc = (range.Cells[rCnt, this.fieldChooser.StartDateLoc] as Excel.Range).get_Value();
                if (bbc != null)
                    student.StartDate = bbc;

                bbc = (range.Cells[rCnt, this.fieldChooser.EndDateLoc] as Excel.Range).get_Value();
                if (bbc != null)
                    student.EndDate = bbc;
                var attendance = (double)(range.Cells[rCnt, this.fieldChooser.Current] as Excel.Range).Value2;
                
                if (attendance != null) {

                    student.CurrentAttendace = attendance;
                }
                attendance = (double)(range.Cells[rCnt, this.fieldChooser.Overall] as Excel.Range).Value2;
                  
                if (attendance != null) 

                {

                    student.OverallAttendace = attendance;
                }

               

               
               // this.srcStartersStudentList.Add(student);
                
            
                //string value = row[this.fieldChooser.StNoLoc].ToString();
                //student.StudentNo = value;
               
                //tempValue = row[this.fieldChooser.StNameLoc].ToString();
                //if (!String.IsNullOrEmpty(tempValue))
                //    student.Name = tempValue;

                //tempValue = row[this.fieldChooser.StudentVisa].ToString();
                //if (!String.IsNullOrEmpty(tempValue))
                //    student.Visa = tempValue;

                //tempValue = row[this.fieldChooser.WarnCat].ToString();
                //if (!String.IsNullOrEmpty(tempValue))
                //    student.Warning = tempValue;

                //tempValue = row[this.fieldChooser.CourseNameLoc].ToString();
                //if (!String.IsNullOrEmpty(tempValue))
                //    student.CourseCode = tempValue;


                //tempValue = row[this.fieldChooser.StartDateLoc].ToString();
                //if (!String.IsNullOrEmpty(tempValue))
                //    student.StartDate = DateTime.Parse(tempValue);


                //tempValue = row[this.fieldChooser.EndDateLoc].ToString();
                //if (!String.IsNullOrEmpty(tempValue))
                //    student.EndDate = DateTime.Parse(tempValue);

                //tempValue = row[this.fieldChooser.Current].ToString();
                //int temp1;
                //if (int.TryParse(tempValue, out temp1))
                //{
                //    student.CurrentAttendace = temp1;
                //}
                //tempValue = row[this.fieldChooser.Overall].ToString();
                //if (int.TryParse(tempValue, out temp1))
                //{
                //    student.OverallAttendace = temp1;
                //}
              
                           

                studentList.Add(student);
                
                
            }

            // con.Close();

            return studentList;


        }
    }
}
