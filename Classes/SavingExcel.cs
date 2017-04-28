using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace EnrolmentTool.Classes
{
    class SavingExcel
    {
        private Microsoft.Office.Interop.Excel.Application xlApp;
        private List<Attendance> attendanceList;
        private List<Starters> statersList;

        public SavingExcel(List<Starters> result)
        {

            this.statersList = result;
           // this.WriteResultToExcelSterter();
        }
        public SavingExcel(List<Attendance> result)
        {
            this.attendanceList = result;
           // this.WriteResultToExcelAttendance();
        }
        public void WriteResultToExcelSterter()
        {
            this.xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (this.xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }
            this.xlApp.Visible = false;
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(1);
            Excel.Sheets xlWorksheet = xlWorkBook.Worksheets;
            var xlNewSheet = (Excel.Worksheet)xlWorksheet.Add(xlWorksheet[1]);
            int row = 1;
            
            foreach (Starters student in statersList)
            {
                xlNewSheet.Cells[row, 1] = student.StudentNo;
                xlNewSheet.Cells[row, 2] = student.Name;
                xlNewSheet.Cells[row, 3] = student.Visa;
                xlNewSheet.Cells[row, 4] = student.CourseCode;
                xlNewSheet.Cells[row, 5] = student.StartDate;
                xlNewSheet.Cells[row, 6] = student.EndDate;
                xlNewSheet.Cells[row, 10] = student.Agent;
                if(student.Visa== "Student"){
                    xlNewSheet.Cells[row, 11] = student.CoEStartDate;
                    xlNewSheet.Cells[row, 12] = student.CoEEndDate;
                    xlNewSheet.Cells[row, 13] = student.CoEDescription;
                    xlNewSheet.Cells[row, 14] = student.CoEStatus;
                    if (student.CoEStatus != "Okay")
                    {
                        switch (student.CoEDateChecker)
                        {
                            case 1:
                                xlNewSheet.Cells[row, 11].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                break;
                            case 2:
                                xlNewSheet.Cells[row, 12].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                break;
                            case 3:
                                xlNewSheet.Cells[row, 11].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                xlNewSheet.Cells[row, 12].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                break;
                            case 4:
                                xlNewSheet.Cells[row, 13].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                break;
                            case 5:
                                xlNewSheet.Cells[row, 11].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                xlNewSheet.Cells[row, 13].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                break;
                            case 6:
                                xlNewSheet.Cells[row, 12].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                xlNewSheet.Cells[row, 13].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                break;
                            case 7:
                                xlNewSheet.Cells[row, 11].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                xlNewSheet.Cells[row, 12].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                xlNewSheet.Cells[row, 13].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                break;

                        }
                    }

                } else xlNewSheet.Cells[row, 14] = "N/A";
                row++;
                
                ////ArrayList payments = student.GetPayments();
                //foreach (Payment payment in payments)
                //{
                //    xlNewSheet.Cells[row, 1] = student.StNo;
                //    if (payment.DueDate < System.DateTime.Today)
                //    {


                //        Excel.Range cell = xlNewSheet.Cells[row, 7];
                //        cell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                //    }
                //    xlNewSheet.Cells[row, 7] = payment.DueDate;
                //    xlNewSheet.Cells[row, 8] = payment.PaymentAmount;
                //    xlNewSheet.Cells[row, 9] = payment.Note;
                //    row++;
                //}
            }


            //Excel.Range range = xlNewSheet.UsedRange;
            //range.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            String filename = xlApp.GetSaveAsFilename("result.xls", "Excel files (*.xls), *.xls");
            xlWorkBook.SaveAs(filename, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            xlWorkBook.Close(true, Type.Missing, Type.Missing);
            xlApp.Quit();
            MessageBox.Show("Done!!!!");

        }
        public void WriteResultToExcelAttendance()
        {
            this.xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (this.xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }
            this.xlApp.Visible = false;
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(1);
            Excel.Sheets xlWorksheet = xlWorkBook.Worksheets;
            var xlNewSheet = (Excel.Worksheet)xlWorksheet.Add(xlWorksheet[1]);
            int row = 1;
            
            foreach (Attendance student in attendanceList)
            {
                xlNewSheet.Cells[row, 1] = student.StudentNo;
                xlNewSheet.Cells[row, 2] = student.Name;
                xlNewSheet.Cells[row, 3] = student.Visa;
                xlNewSheet.Cells[row, 4] = student.NewWarning;
                xlNewSheet.Cells[row, 5] = student.CourseCode;
                xlNewSheet.Cells[row, 6] = student.StartDate;
                xlNewSheet.Cells[row, 7] = student.EndDate;
                xlNewSheet.Cells[row, 8] = student.CurrentAttendace;
                xlNewSheet.Cells[row, 9] = student.OverallAttendace;
                
                
                //foreach (Payment payment in payments)
                //{
                //    xlNewSheet.Cells[row, 1] = student.StNo;
                //    if (payment.DueDate < System.DateTime.Today)
                //    {


                //        Excel.Range cell = xlNewSheet.Cells[row, 7];
                //        cell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                //    }
                //    xlNewSheet.Cells[row, 7] = payment.DueDate;
                //    xlNewSheet.Cells[row, 8] = payment.PaymentAmount;
                //    xlNewSheet.Cells[row, 9] = payment.Note;
                row++;
                //}
            }


            //Excel.Range range = xlNewSheet.UsedRange;
            //range.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            
            String filename = xlApp.GetSaveAsFilename("result.xls", "Excel files (*.xls), *.xls");
            if (string.IsNullOrEmpty(filename)) {
                filename = "result.xls";
            }
            xlWorkBook.SaveAs(filename, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            xlWorkBook.Close(true, Type.Missing, Type.Missing);
            xlApp.Quit();
            MessageBox.Show("Done!!!!");

        }
    }
}
