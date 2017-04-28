using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace EnrolmentTool.Classes
{
    class XmlReader
    {
        private XElement xmlFile;
        private int stNoLoc;
        private int stNameLoc;
        private int studentVisa;
        private int agentLoc;
        private int courseNameLoc;
        private int startDateLoc;
        private int endDateLoc;

        private int coeStartDateLoc; 
        private int coeEndDateLoc ;
        private int coeDescriptionLoc;
        private int current; 
        private int overall; 
        private int warnCat;

        //public int CoeStartDateLoc { get { return this.coeStartDateLoc; } set { this.coeStartDateLoc = value; } }
        public int CoeEndDateLoc { get { return this.coeEndDateLoc; } set { this.coeEndDateLoc = value; } }
        public int CoeDescriptionLoc { get { return this.coeDescriptionLoc; } set { this.coeDescriptionLoc = value; } }
        public int CoeStartDateLoc { get { return this.coeStartDateLoc; } set { this.coeStartDateLoc = value; } }

        public int Current
        {
            get { return current;}
            set { this.current = value; }
        }
        public int Overall
        {
            get { return overall;}
            set { this.overall = value; }
        }
        public int WarnCat
        {
            get { return warnCat;}
            set { this.warnCat = value; }
        }

        private int coEFrom;

        public int CoEFrom
        {
            get { return coEFrom; }
            set { this.coEFrom = value; }
        }
        
        

        private int coETo;
        
        public int CoETo
        {
            get { return coETo; }
            set { this.coETo = value; }

        }
        private int coeCode;

        public int CoeCode
        {
            get { return CoeCode; }
            set{ this.CoeCode = value;}
        }

        private int coeDescription;
        public int CoEDescription
        {
            get { return coeDescription; }
            set { this.coeDescription = value; }

        }

        public static int STARTER = 1;
        public static int ATTENDACE = 2;

        public int StudentVisa
        {
            get { return studentVisa; }
            set { this.studentVisa = value; }
        }
        
        

        public int EndDateLoc
        {
            get { return endDateLoc; }
            set { endDateLoc = value; }
        }

        public int StartDateLoc
        {
            get { return startDateLoc; }
            set { startDateLoc = value; }
        }
        public int CourseNameLoc
        {
            get { return courseNameLoc; }
            set { courseNameLoc = value; }
        }


        public int AgentLoc
        {
            get { return agentLoc; }
            set { agentLoc = value; }
        }

        public int StNameLoc
        {
            get { return stNameLoc; }
            set { stNameLoc = value; }
        }

        public int StNoLoc
        {
            get { return stNoLoc; }
            set { stNoLoc = value; }
        }


        public XElement XmlFile
        {
            get { return xmlFile; }
            set { xmlFile = value; }
        }
        public XmlReader(int flag)
        {
            if(flag== STARTER)
            try
            {

                this.ReadDataStarters();
            }
            catch (Exception e)
            {
                //throw (e.Message);
            }
            //xmlFile.Save("courseproperty.xml");
            else if(flag == ATTENDACE)
                try
                {

                    this.ReadDataAttendance();
                }
                catch (Exception e)
                {
                    //throw (e.Message);
                }
        }
        private void ReadDataStarters()
        {
            try
            {
                xmlFile = XElement.Load("datalocation.xml");


            }
            catch (Exception e)
            {

            }
            var query = from c in xmlFile.Elements()
                        select c;

            foreach (XElement el in query)
            {
                //Console.WriteLine(el.Name);

                //Console.WriteLine( el.Attribute("name"));
                if (el.Name == "Starters")
                {

                    int.TryParse(el.Attribute("Number").Value, out this.stNoLoc);
                    int.TryParse(el.Attribute("StudentName").Value, out this.stNameLoc);
                    int.TryParse(el.Attribute("StudentVisa").Value, out this.studentVisa);                   
                    int.TryParse(el.Attribute("CourseName").Value, out this.courseNameLoc);
                    int.TryParse(el.Attribute("CourseStartDate").Value, out this.startDateLoc);
                    int.TryParse(el.Attribute("CourseEndDate").Value, out this.endDateLoc);
                    int.TryParse(el.Attribute("AgentName").Value, out this.agentLoc);
                    int.TryParse(el.Attribute("CoEStartDate").Value, out this.coeStartDateLoc);
                    int.TryParse(el.Attribute("CoEEndDate").Value, out this.coeEndDateLoc);
                    int.TryParse(el.Attribute("CoEDescription").Value, out this.coeDescriptionLoc);

                }


            }
        }
        private void ReadDataAttendance()
        {
            try
            {
                xmlFile = XElement.Load("datalocation.xml");


            }
            catch (Exception e)
            {

            }
            var query = from c in xmlFile.Elements()
                        select c;

            foreach (XElement el in query)
            {
                //Console.WriteLine(el.Name);

                //Console.WriteLine( el.Attribute("name"));
                if (el.Name == "Attendance")
                {

                    int.TryParse(el.Attribute("Number").Value, out this.stNoLoc);
                    int.TryParse(el.Attribute("StudentName").Value, out this.stNameLoc);
                    int.TryParse(el.Attribute("StudentVisa").Value, out this.studentVisa);
                    int.TryParse(el.Attribute("WarnCategory").Value, out this.warnCat);
                    int.TryParse(el.Attribute("CourseCode").Value, out this.courseNameLoc);
                    int.TryParse(el.Attribute("StartDate").Value, out this.startDateLoc);
                    int.TryParse(el.Attribute("EndDate").Value, out this.endDateLoc);
                    int.TryParse(el.Attribute("Current").Value, out this.current);
                    int.TryParse(el.Attribute("Overall").Value, out this.overall);


                }


            }
        }
     
    }
}
