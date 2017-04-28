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
    class StartersChecker
    {
           
        private string sourceFile;
        //private string destFile;
        private Excel.Workbook srcBook;
        //private Excel.Workbook destBook;
        private Excel.Application xlApp;

        private XmlReader fieldChooser= null;
        private List<Attendance> srcStudentList = null;
        private List<Starters> srcStartersStudentList = null;
        private List<Attendance> resultStudentList = null;
        private String connectionStringSrc = null;
        
      //  private Loading LoadingForm ;
      //  private System.Timers.Timer timer1 = new System.Timers.Timer();

        public StartersChecker(string src)
        {
            try
            {
                /* field chooser 
                 * Read data location from xml file and let the program know where the data is.
                */
                connectionStringSrc = src;

                fieldChooser = new XmlReader(XmlReader.STARTER);
                
                xlApp = new Excel.Application();
                srcBook = xlApp.Workbooks.Open(src);
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
            Excel.Worksheet ws1 = this.srcBook.Worksheets.get_Item(1);
            Excel.Range range = ws1.UsedRange;

            srcStartersStudentList = new List<Starters>();

           // this.connectionStringSrc = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Extended Properties=\"Excel 14.0;\"", this.connectionStringSrc);

           // this.srcStartersStudentList = this.CreateStarters(connectionStringSrc);
           
                       
            //this.StartCheck();


            int rCnt = 0;

            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {
                string tempValue;
                Starters student = new Starters();
                string value = (string) (range.Cells[rCnt, this.fieldChooser.StNoLoc] as Excel.Range).Value2;
                student.StudentNo = value;


                tempValue = (string) (range.Cells[rCnt,this.fieldChooser.StNameLoc] as Excel.Range).Value2;
                if (!String.IsNullOrEmpty(tempValue))
                    student.Name = tempValue;

                tempValue = (string) (range.Cells[rCnt,this.fieldChooser.StudentVisa] as Excel.Range).Value2;
                if (!String.IsNullOrEmpty(tempValue))
                    student.Visa = tempValue;

                tempValue = (string) (range.Cells[rCnt,this.fieldChooser.CourseNameLoc] as Excel.Range).Value2;
                if (!String.IsNullOrEmpty(tempValue))
                    student.CourseCode = tempValue;

                
                var bbc =(range.Cells[rCnt,this.fieldChooser.StartDateLoc] as Excel.Range).get_Value();
                if(bbc !=null)
                student.StartDate = bbc;

                bbc = (range.Cells[rCnt,this.fieldChooser.EndDateLoc] as Excel.Range).get_Value();
                if (bbc != null)
                    student.EndDate = bbc;
                tempValue = (string) (range.Cells[rCnt,this.fieldChooser.AgentLoc] as Excel.Range).Value2;
                if (!String.IsNullOrEmpty(tempValue))
                    student.Agent = tempValue;

                var abc = (range.Cells[rCnt, this.fieldChooser.CoeStartDateLoc] as Excel.Range).get_Value();
                if (abc != null)
                    student.CoEStartDate = abc;


                abc = (range.Cells[rCnt, this.fieldChooser.CoeEndDateLoc] as Excel.Range).get_Value();
                if (abc != null)
                    student.CoEEndDate = abc;

                tempValue = (string) (range.Cells[rCnt,this.fieldChooser.CoeDescriptionLoc] as Excel.Range).get_Value();
                if (!String.IsNullOrEmpty(tempValue))
                    student.CoEDescription = tempValue;

                this.srcStartersStudentList.Add(student);
            }


                //for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                //{

                //Starters student = new Starters();



                //string tempValue;
                //string value = (string) (range.Cells[rCnt,this.fieldChooser.StNoLoc] as Excel.Range).Value2;
                //student.StudentNo = value;

                //tempValue = row[this.fieldChooser.StNameLoc].ToString();
                //if (!String.IsNullOrEmpty(tempValue))
                //    student.Name = tempValue;

                //tempValue = row[this.fieldChooser.StudentVisa].ToString();
                //if (!String.IsNullOrEmpty(tempValue))
                //    student.Visa = tempValue;

                //tempValue = row[this.fieldChooser.CourseNameLoc].ToString();
                //if (!String.IsNullOrEmpty(tempValue))
                //    student.CourseCode = tempValue;


                //tempValue = row[this.fieldChooser.StartDateLoc].ToString();
                //if (!String.IsNullOrEmpty(tempValue))
                //    student.StartDate = DateTime.Parse(tempValue);


                //tempValue = row[this.fieldChooser.EndDateLoc].ToString();
                //if (!String.IsNullOrEmpty(tempValue))
                //    student.EndDate = DateTime.Parse(tempValue);

                //tempValue = row[this.fieldChooser.AgentLoc].ToString();
                //if (!String.IsNullOrEmpty(tempValue))
                //    student.Agent = tempValue;

                //tempValue = row[this.fieldChooser.CoeStartDateLoc].ToString();
                //if (!String.IsNullOrEmpty(tempValue))
                //    student.CoEStartDate = DateTime.Parse(tempValue);


                //tempValue = row[this.fieldChooser.CoeEndDateLoc].ToString();
                //if (!String.IsNullOrEmpty(tempValue))
                //    student.CoEEndDate = DateTime.Parse(tempValue);

                //tempValue = row[this.fieldChooser.CoeDescriptionLoc].ToString();
                //if (!String.IsNullOrEmpty(tempValue))
                //    student.CoEDescription = tempValue;



                //    studentList.Add(student);


                //}

            
                this.srcBook.Close(true, Type.Missing, Type.Missing);

        //private Excel.Workbook destBook;
                this.xlApp.Quit();

                if (this.srcStartersStudentList.Count != 0)
                {
                    foreach (Starters st in this.srcStartersStudentList)
                    {
                        st.CheckCourseCode();
                    }
                    this.MakeResults();
                }
        }

        private void MakeResults()
        {

            SavingExcel se = new SavingExcel(this.srcStartersStudentList);

            se.WriteResultToExcelSterter();
        }
       
   
       
     
      


        private void StartCheck()
        {
            
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



        private List<Starters> CreateStarters(string connectionStr)
        {
            
            // Creat OLE connection object

            OleDbConnection con = new OleDbConnection(connectionStr);
            con.Open();

            List<Starters> studentList = new List<Starters>();
            

            
            /*pull the data from excel file and store in dataset
             *pass to data to datatable collection.
            */

            //OleDbDataAdapter dtAdapter = new OleDbDataAdapter("Select * From [Sheet1$]", connectionStr);
            var adapter = new OleDbDataAdapter("Select * From [Sheet1$]",con);
            var ds = new DataSet();

            adapter.Fill(ds, "student");
            DataTable data = ds.Tables["student"];

            string tempValue = "";
            
            // call row data and store data in student object.
            foreach (DataRow row in data.Rows)
            {
                Starters student = new Starters();




                string value = row[this.fieldChooser.StNoLoc].ToString();
                student.StudentNo = value;

                tempValue = row[this.fieldChooser.StNameLoc].ToString();
                if (!String.IsNullOrEmpty(tempValue))
                    student.Name = tempValue;

                tempValue = row[this.fieldChooser.StudentVisa].ToString();
                if (!String.IsNullOrEmpty(tempValue))
                    student.Visa = tempValue;

                tempValue = row[this.fieldChooser.CourseNameLoc].ToString();
                if (!String.IsNullOrEmpty(tempValue))
                    student.CourseCode = tempValue;


                tempValue = row[this.fieldChooser.StartDateLoc].ToString();
                if (!String.IsNullOrEmpty(tempValue))
                    student.StartDate = DateTime.Parse(tempValue);


                tempValue = row[this.fieldChooser.EndDateLoc].ToString();
                if (!String.IsNullOrEmpty(tempValue))
                    student.EndDate = DateTime.Parse(tempValue);

                tempValue = row[this.fieldChooser.AgentLoc].ToString();
                if (!String.IsNullOrEmpty(tempValue))
                    student.Agent = tempValue;

                tempValue = row[this.fieldChooser.CoeStartDateLoc].ToString();
                if (!String.IsNullOrEmpty(tempValue))
                    student.CoEStartDate = DateTime.Parse(tempValue);


                tempValue = row[this.fieldChooser.CoeEndDateLoc].ToString();
                if (!String.IsNullOrEmpty(tempValue))
                    student.CoEEndDate = DateTime.Parse(tempValue);

                tempValue = row[this.fieldChooser.CoeDescriptionLoc].ToString();
                if (!String.IsNullOrEmpty(tempValue))
                    student.CoEDescription = tempValue;

                

                studentList.Add(student);
                
                
            }

            // con.Close();

            return studentList;


        }
    }
}
