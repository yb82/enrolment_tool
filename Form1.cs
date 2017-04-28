using EnrolmentTool.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EnrolmentTool
{
    public partial class EnrolForm : Form
    {
        private float total;
        private OpenFileDialog fileDialog = null;
        private Processing aa = null;
        private AttendanceCal cal;
        private StartersChecker starters;
        static List<Course> courses = new List<Course>();
        static List<float> payments = new List<float>();
        static int paymentCounter = 0;
        static int courseCounter = 0;
        public EnrolForm()
        {
            InitializeComponent();
        }
        private void initOpenDialog()
        {
            this.fileDialog = new OpenFileDialog();
            this.fileDialog.Title = "Open file";

        }
        private void EnrolForm_Load(object sender, EventArgs e)
        {
            tabControl1.TabPages[0].Text = "Date Calculator";
            tabControl1.TabPages[1].Text = "Check Lists";
        }
        private DateTime changeStringToDate(string s)
        {
            var culture = CultureInfo.GetCultureInfo("en-AU");
            return DateTime.ParseExact(s, "dd/MM/yyyy", culture);
        }

        private void tbxWeeks_TextChanged(object sender, EventArgs e)
        {
            this.weeksCalculator();
        }

        private void tbxSD1_TextChanged(object sender, EventArgs e)
        {
            this.weeksCalculator();
        }
        private void weeksCalculator()
        {
            int duration = 0;
            if (tbxSD1.TextLength == 10 && int.TryParse(tbxWeeks.Text, out duration))
            {
                DateTime date = this.changeStringToDate(tbxSD1.Text);
                DateTime finish = date.AddDays(duration * 7 - 3);

                tbxFD.Text = finish.ToString("dd/MM/yyyy");
                tbxFM.Text = finish.AddDays(3).ToString("dd/MM/yyyy");
                tbxM1.Text = finish.AddDays(10).ToString("dd/MM/yyyy");
                tbxM2.Text = finish.AddDays(17).ToString("dd/MM/yyyy");
                tbxM4BfFriday.Text = finish.AddDays(28).ToString("dd/MM/yyyy");
                tbxM4.Text = finish.AddDays(31).ToString("dd/MM/yyyy");
                tbxM8.Text = finish.AddDays(59).ToString("dd/MM/yyyy");
                txbM12.Text = date.AddDays(87).ToString("dd/MM/yyyy");
                
               // txbW26F.Text = date.AddDays(7 * 26 - 3).ToString("dd/MM/yyyy"); 

            }

        }

        private void tbxSD2_TextChanged(object sender, EventArgs e)
        {
            if (tbxSD2.TextLength == 10)
            {
                this.paymentDateCalculator();
            }
        }
        private void paymentDateCalculator()
        {
            DateTime date = this.changeStringToDate(tbxSD2.Text);
            tbxFBS.Text = date.AddDays(-7).ToString("dd/MM/yyyy");
            tbxF4.Text = date.AddDays(7 * 4 - 3).ToString("dd/MM/yyyy");
            tbxF8.Text = date.AddDays(7 * 8 - 3).ToString("dd/MM/yyyy");
            tbxF12.Text = date.AddDays(7 * 12 - 3).ToString("dd/MM/yyyy");
            tbxF16.Text = date.AddDays(7 * 16 - 3).ToString("dd/MM/yyyy");
            tbxF20.Text = date.AddDays(7 * 20 - 3).ToString("dd/MM/yyyy");
            tbxF24.Text = date.AddDays(7 * 24 - 3).ToString("dd/MM/yyyy");
            tbxF28.Text = date.AddDays(7 * 28 - 3).ToString("dd/MM/yyyy");
            tbxF32.Text = date.AddDays(7 * 32 - 3).ToString("dd/MM/yyyy");
            tbx4F.Text = date.AddDays(-(7 * 4) - 3).ToString("dd/MM/yyyy");
        }


        private void btnRst1_Click(object sender, EventArgs e)
        {
            tbxFD.Text = "";
            tbxFM.Text = "";
            tbxM1.Text = "";
            tbxM2.Text = "";
            tbxM4.Text = "";
            tbxM4BfFriday.Text = "";
            tbxM8.Text = "";

        }

        private void btnRst2_Click(object sender, EventArgs e)
        {
            tbxFBS.Text = "";
            tbxF4.Text = "";
            tbxF8.Text = "";
            tbxF12.Text = "";
            tbxF16.Text = "";
            tbxF20.Text = "";
            tbxF24.Text = "";
            tbxF28.Text = "";
            tbxF32.Text = "";
            tbx4F.Text = "";
        }
        private void refreshTotalTution()
        {

            this.total = 0;
            if (listBox1.Items.Count > 0)
            {
                foreach (string item in listBox1.Items)
                {
                    float temp1;
                    float.TryParse(item, out temp1);
                    total += temp1;

                }
                tbxTotal.Text = "$" + total.ToString();
            }
        }
        private void btnAddCF_Click(object sender, EventArgs e)
        {
            double temp = 0;
            if (tbxCFee.Text != "" && Double.TryParse(tbxCFee.Text, out temp))
            {
                listBox1.Items.Add(tbxCFee.Text);
            }
            this.refreshTotalTution();
        }

        private void btnDelCF_Click(object sender, EventArgs e)
        {
            int index = listBox1.Items.Count - 1;
            if (index > -1)
            {
                listBox1.Items.RemoveAt(index);
            }
            this.refreshTotalTution();
        }

        private void btnAddPmt_Click(object sender, EventArgs e)
        {
            double temp = 0;
            if (tbxPmt.Text != "" && Double.TryParse(tbxPmt.Text, out temp))
            {
                listBox2.Items.Add(tbxPmt.Text);
            }
            this.refreshTotalPayment();
        }

        private void btnDelPmt_Click(object sender, EventArgs e)
        {
            int index = listBox2.Items.Count - 1;
            if (index > -1)
            {
                listBox2.Items.RemoveAt(index);
            }
            this.refreshTotalPayment();
        }
        private void refreshTotalPayment()
        {

            this.total = 0;
            if (listBox1.Items.Count > 0)
            {
                foreach (string item in listBox2.Items)
                {
                    float temp1;
                    float.TryParse(item, out temp1);
                    total += temp1;

                }
                tbxTPayment.Text = "$" + total.ToString();
            }
        }

        private void btnCalPayments_Click(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            courses.Clear();
            payments.Clear();
            float paymentstotal = 0;
            this.total = 0;
            paymentCounter = 0;
            if (listBox1.Items.Count > 0)
            {
                foreach (string item in listBox1.Items)
                {
                    float temp1;
                    float.TryParse(item, out temp1);
                    total += temp1;

                }
               // tbxTPayment.Text = "$" + total.ToString();
            }
            if (listBox2.Items.Count > 0)
            {
                foreach (string item3 in listBox2.Items)
                {
                    float temp1;
                    float.TryParse(item3, out temp1);
                    paymentstotal += temp1;

                }
                //totalbox.Text = "$" + total.ToString();
            }
            if (this.total != paymentstotal)
            {
                MessageBox.Show("Total amount of payment plan does not match!");
                return;
            }

            List<Course> resultCourses = new List<Course>();

            Course a;
            //float remain = 0;
            foreach (string item1 in listBox2.Items)
            {

                float payment;

                float.TryParse(item1, out payment);
                payments.Add(payment);

            }
            foreach (string item in listBox1.Items)
            {
                float courseTuition;
                float.TryParse(item, out courseTuition);
                a = new Course();
                a.Tution = courseTuition;
                courses.Add(a);
            }

            //int totalCourses = courses.Count;
            //int totalPayments = payments.Count;
            float aPayment = getNextPayment();
            // float remain = 0;
            foreach (Course ab in courses)
            {
                float courseTuition = ab.Tution;
              
                if (courseTuition > aPayment)
                {
                    while (courseTuition > 0)
                    {
                        ab.Payments.Add(aPayment);
                        courseTuition -= aPayment;
                        
                        aPayment = getNextPayment();
                        if (aPayment == 0 && courseTuition != 0 )
                        {
                            ab.Payments.Add(courseTuition);
                            break;
                        }
                        if (courseTuition < aPayment && courseTuition != 0)
                        {

                            aPayment-= courseTuition;
                            ab.Payments.Add(courseTuition);
                            break;
                        }
                    }
                } 
                else
                {
                    ab.Payments.Add(courseTuition);
                    aPayment -= courseTuition;
                }


            }



            //foreach (string item in listBox1.Items)
            //{
            //    a = new Course();
            //    float courseTuition;

            //    float.TryParse(item, out courseTuition);
            //    a.Tution = courseTuition;
            //    float remain = 0;
            //    //a.Percent = temp1 / this.total;




            //    foreach (string item1 in listBox2.Items)
            //    {
            //        float payment;

            //        float.TryParse(item1, out payment);
            //        while (courseTuition > 0) {
            //            if (remain > 0 && courseTuition-remain >0 )
            //            {
            //                a.Payments.Add(remain);
            //                courseTuition -= remain;
            //            }
            //            if (courseTuition >= payment)
            //            {
            //                a.Payments.Add(payment); 
            //                courseTuition -= payment;
            //            }


            //            if (courseTuition > 0 && courseTuition <payment)
            //            {
            //                a.Payments.Add(courseTuition); 
            //                remain = courseTuition;
            //                courseTuition = 0;
            //                break;
            //            }

            //        }

            //        //a.Payments.Add(temp2 * a.Percent);
            //    }
            //courses.Add(a);
            //}

            foreach (Course aaaa in courses)
            {
                listBox3.Items.Add("=====================");
                foreach (float payment in aaaa.Payments)
                {
                    listBox3.Items.Add(payment.ToString("c2"));
                }
            }
        }




       

        private float getNextPayment()
        {
            if (payments.Count > paymentCounter)
                return payments[paymentCounter++];
            else return 0;
        }

        private void btnRstPayments_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            listBox3.Items.Clear();
            tbxTotal.Text = "";
            tbxPmt.Text = "";
            tbxCFee.Text = "";
           
        }

        private void btnSRC_Click(object sender, EventArgs e)
        {
            this.initOpenDialog();
            if (this.fileDialog.ShowDialog() == DialogResult.OK)
            {
                tbxSrc.Text = fileDialog.FileName;
            }
        }

        private void btnDest_Click(object sender, EventArgs e)
        {
            //this.initOpenDialog();
            //if (this.fileDialog.ShowDialog() == DialogResult.OK)
            //{
            //   // tbxDest.Text = fileDialog.FileName;
            //}
        }

        private void btnAtt_Click(object sender, EventArgs e)
        {
            this.initOpenDialog();
            if (this.fileDialog.ShowDialog() == DialogResult.OK)
            {
                tbxAtt.Text = fileDialog.FileName;
            }

        }
        private void LoadingWindow()
        {
            this.aa = new Processing();
            this.aa.ShowDialog();
        }
        private void btnCheckStarters_Click(object sender, EventArgs e)
        {
            bool flagSRC= true;

            if (tbxSrc.Text == "")
            {
                flagSRC = false;
                MessageBox.Show("Please select a starters' list");
            }
            
            if (flagSRC )
            {
               // Thread loadingWindow = new Thread(this.LoadingWindow);
                //loadingWindow.Start();

                starters = new StartersChecker(tbxSrc.Text);
                starters.Start();
                //loadingWindow.Abort();
            }
        }

        private void btnCheckAttendance_Click(object sender, EventArgs e)
        {

            if (tbxAtt.Text == "")
            {
                                MessageBox.Show("Please select current student list file");
            }
            else
            {
                cal = new AttendanceCal(tbxAtt.Text);
                cal.Start();

                //Thread loadingWindow = new Thread(this.LoadingWindow);
                //loadingWindow.Start();
                // loadingWindow.Join();
                //cal = new AttendanceCal(tbxAtt.Text);
                //cal.Start();
                //loadingWindow.Abort();
            }
        }

    }
}
