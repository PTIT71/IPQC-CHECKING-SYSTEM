using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using IPQC_CHECKING_SYSTEM;
using IPQC_CHECKING_SYSTEM.Common;
using IPQC_CHECKING_SYSTEM.Model;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace IPQC_CHECKING_SYSTEM
{
    public partial class MainWindow : Form
    {
        
        List<IPQC> lstSource = null;
        List<IPQC> lstView = new List<IPQC>();
        
        int index = 0;
        int rowCount = 15;

        private void ChangeDisplayData()
        {
            if (lstSource.Count <= rowCount)
            {
               lstView = lstSource;
               index = 0;
               return;
            }
            lstView.Clear();
            int i = index;
            while (lstView.Count< lstSource.Count && lstView.Count<=rowCount)
            {
                if (i >= lstSource.Count - 1)
                    i = 0;
                lstView.Add(lstSource[i]);
                i++;                
            }
            index++;
            if (index >= lstSource.Count - 1)
                index = 0;
            dgvData.Refresh();            
        }
        public MainWindow()
        {
            InitializeComponent();
            lstSource = new List<IPQC>();
            readExcelFile();
            lbl_CurrentDate.Text = "Date: " + DateTime.Now.Day + "/" + DateTime.Now.Month + "/" +DateTime.Now.Year;
        }

        public bool DesDisplay()
        {
            bool flat = false;
            try
            {
                for (int i = 0; i < lstSource.Count; i++)
                {
                    if (lstSource[i].Result != null && lstSource[i].Result.Equals("OK"))
                    {
                        DateTime date1 = DateTime.Parse(lstSource[i].ReleaseTime);
                        string dateTime = DateTime.Now.Hour + ":" + DateTime.Now.Minute;
                        DateTime date2 = DateTime.Parse(dateTime);

                        DateTime compete = DateTime.Parse(date2.Subtract(date1).ToString());
                        int total = compete.Hour * 60 + compete.Minute;
                        if (total >= 2)
                        {
                            lstSource[i].isNOTShown = "F";
                            flat = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return flat;
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            //Danh sach data chua cac POINT trong driver
            List<DataSending> lstData = new List<DataSending>();
            for (int z = 0; z < Driver.lstDataSending.Count; z++)
            {
                lstData.Add(Driver.lstDataSending[z]);
            }
            Driver.lstDataSending.Clear();

            //Neu lstData co gia tri (Trong khoang thoi gian nay nguoi ta co check)
            if (lstData.Count > 0)
            {
                for (int r = 0; r < lstData.Count; r++)
                {
                    DataSending dataSend = lstData[r];
                    if (lstSource.Count == 0)
                    {
                        IPQC ipItem = new IPQC();
                        ipItem.PartNumber = dataSend.partNumber;
                        ipItem.SubmitPIC = dataSend.employee;
                        if (dataSend.value.Trim().Equals(BTN.PRESS_1) || dataSend.value.Trim().Equals(BTN.PRESS_2))
                            ipItem.Type = TYPE.BUY_OFF_SAMPLE;
                        ipItem.TimeSubmit = getCurrentHourAndMinute();
                        ipItem.Status = Common.STATUS.WAITING;
                        // update online time list source
                        lstSource.Add(ipItem);
                    }
                    else
                    {
                        if (searchData(lstData[r].partNumber) != -1)
                        {
                            int z = searchData(lstData[r].partNumber);
                            // update online time list source
                            switch (dataSend.value)
                            {
                                //case 1 and 2: 
                                case BTN.PRESS_1:
                                case BTN.PRESS_2:
                                    //is repaired
                                    if (STATUS.REPAIR.Equals(lstSource[z].Status) == true)
                                    {
                                        lstSource[z].isRepaired = "T";
                                        IPQC ipItem = new IPQC(); 
                                        ipItem.SubmitPIC = dataSend.employee;
                                        ipItem.PartNumber = dataSend.partNumber;
                                        ipItem.Type = TYPE.REPAIR;
                                        ipItem.TimeSubmit = getCurrentHourAndMinute();
                                        ipItem.Status = Common.STATUS.WAITING;
                                        lstSource.Add(ipItem);
                                    }
                                    else
                                    {
                                        MessageBox.Show("Thông tin sản phẩm tồn tại");
                                    }
                                    break;
                                case BTN.PRESS_3:
                                    lstSource[z].TimeRecive = getCurrentHourAndMinute();
                                    lstSource[z].Status = STATUS.CHECKING;
                                    lstSource[z].IPQC1 = dataSend.employee;
                                    break;
                                case BTN.PRESS_OK:
                                    {
                                        lstSource[z].ReleaseTime = getCurrentHourAndMinute();
                                        lstSource[z].Result = RESULT.OK;
                                        DateTime date1 = DateTime.Parse(lstSource[z].ReleaseTime);
                                        DateTime date2 = DateTime.Parse(lstSource[z].TimeSubmit);
                                        DateTime date3 = DateTime.Parse(lstSource[z].TimeRecive);

                                        DateTime compete = DateTime.Parse(date1.Subtract(date3).ToString());
                                        lstSource[z].CheckingTime = (compete.Hour * 60 + compete.Minute).ToString();
                                        compete = DateTime.Parse(date1.Subtract(date2).ToString());
                                        lstSource[z].TotalTime = (compete.Hour * 60 + compete.Minute).ToString();

                                        if (lstSource[z].Type.Equals(TYPE.REPAIR))
                                        {
                                            for (int j = 0; j < lstSource.Count; j++)
                                            {
                                                if (lstSource[j].PartNumber.Equals(lstSource[z].PartNumber) && lstSource[j].Result.Equals(RESULT.NG))
                                                {
                                                    lstSource[j].isNOTShown = "True";
                                                }
                                            }
                                        }

                                    }
                                    break;
                                case BTN.PRESS_NG:
                                    {
                                        lstSource[z].ReleaseTime = getCurrentHourAndMinute();
                                        lstSource[z].Result = RESULT.NG;
                                        DateTime date1 = DateTime.Parse(lstSource[z].ReleaseTime);
                                        DateTime date2 = DateTime.Parse(lstSource[z].TimeSubmit);
                                        DateTime date3 = DateTime.Parse(lstSource[z].TimeRecive);

                                        DateTime compete = DateTime.Parse(date1.Subtract(date3).ToString());
                                        lstSource[z].CheckingTime = (compete.Hour * 60 + compete.Minute).ToString();
                                        compete = DateTime.Parse(date1.Subtract(date2).ToString());
                                        lstSource[z].TotalTime = (compete.Hour * 60 + compete.Minute).ToString();

                                    }
                                    break;
                                //case 4 and 5
                                case "4":
                                case "5":
                                    if (lstSource[z].Result == RESULT.NG)
                                    {
                                        lstSource[z].Status = STATUS.REPAIR;
                                        lstSource[z].TimeRepair = getCurrentHourAndMinute();
                                    }
                                    break;
                                default:
                                    MessageBox.Show("Value code Error! Code value" + dataSend.value);
                                    break;
                            }
                        }

                        else
                        {
                            IPQC ipItem = new IPQC();
                            ipItem.PartNumber = dataSend.partNumber;
                            ipItem.SubmitPIC = dataSend.employee;
                            if (dataSend.value.Trim().Equals(BTN.PRESS_1) || dataSend.value.Trim().Equals(BTN.PRESS_2))
                                ipItem.Type = TYPE.BUY_OFF_SAMPLE;
                            ipItem.TimeSubmit = getCurrentHourAndMinute();
                            ipItem.Status = Common.STATUS.WAITING;
                            // update online time list source
                            lstSource.Add(ipItem);
                        }
                    }
                }
                WriteExcelFile();
            }
            else
            {
                if (DesDisplay())
                {
                    WriteExcelFile();
                }
            }
           // Thread th = new Thread(readExcelFile);
          //  th.Start();
            ChangeDisplayData();

            dgvData.Refresh();
        }

        public int searchData(string partNumber)
        {
            int z = 0;
            for (z = 0; z < lstSource.Count; z++)
            {
                try
                {
                    if (lstSource[z].PartNumber.Equals(partNumber) && lstSource[z].isRepaired.Trim().Length == 0)
                    {
                        return z;
                    }
                }
                catch
                {

                }
            }

            return -1;
        }

        public string getCurrentHourAndMinute()
        {
            DateTime date =  DateTime.Now;
            return date.Hour + " : " + date.Minute;
        }

        public string StandarTime(string time)
        {
            if (time != null && time.Trim().Length>0)
            {
                DateTime time1 = DateTime.Parse(time);
                return time1.Hour + " : " + time1.Minute;
            }
            return time;
        }

        private void MainWindow_Shown(object sender, EventArgs e)
        {
            dgvData.AutoGenerateColumns = false;                              
            lbl_CurrentDate.Focus();
            dgvData.EnableHeadersVisualStyles = false;
            dgvData.DataSource = lstView;
           
            t_Reload_Data.Start();
            t_Reload_View.Start();
        }

        private void dgvData_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            e.Paint(e.CellBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.Border);
            using (Pen p = new Pen(Color.Black, 2))
            {
                System.Drawing.Rectangle rect = e.CellBounds;
                rect.Width -= 2;
                rect.Height -= 2;
                e.Graphics.DrawRectangle(p, rect);
            }
            e.Handled = true;
        }

        private void dgvData_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            e.CellStyle.Font = new System.Drawing.Font("Arial", 25);
            if (e.Value != null)
            {
                e.Value = e.Value.ToString().TrimStart().TrimEnd();
                if (("RESULT").Equals(dgvData.Columns[e.ColumnIndex].Name.ToUpper()))
                {
                    string result = e.Value as string;
                    if (result.Contains("OK"))
                        e.CellStyle.BackColor = Color.Green;
                    if (result.Contains("NG"))
                        e.CellStyle.BackColor = Color.Red;
                }
            }
        }

        private void t_Reload_Data_Tick(object sender, EventArgs e)
        {
            MainWindow_Load(this, null);
        }

        private void t_Reload_View_Tick(object sender, EventArgs e)
        {
            ChangeDisplayData();
            dgvData.DataSource = lstView;
            dgvData.Refresh();
        }

        private void SendData_Click(object sender, EventArgs e)
        {
            if (txtValue.Text.Trim().Length > 0 && txtPartNumber.Text.Trim().Length > 0 && txtEmployee.Text.Trim().Length >0)
            {
                DataSending dataSend = new DataSending();
                dataSend.partNumber = txtPartNumber.Text.ToString();
                dataSend.value = txtValue.Text.ToString();
                dataSend.employee = txtEmployee.Text.ToString();
                Driver.lstDataSending.Add(dataSend);
            }
        }

        public void readExcelFile()
        {
            List<IPQC> lstIPQC = new List<IPQC>();
            try
            {
                Workbook MyBook = null;
                Application MyApp = null;
                Worksheet MySheet = null;
                MyApp = new Application();
                MyApp.Visible = false;
                string urlRead = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DATABASE.xlsx");
                MyBook = MyApp.Workbooks.Open(urlRead);
                MySheet = (Worksheet)MyBook.Sheets[1];
                int lastRow = MySheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                int lastColumn = MySheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column;
                //MessageBox.Show("LastRow= " + lastRow + " LastCol= " + lastColumn);
                int step = 2;
                for (int i = step; i <= lastRow; i++)
                {
                    Microsoft.Office.Interop.Excel.Range rowContent_cellLeft = MySheet.Cells[i, 1];
                    Range rowContent_cellRight = MySheet.Cells[i, 13];
                    System.Array rowContent = (System.Array)MySheet.get_Range(rowContent_cellLeft, rowContent_cellRight).Cells.Value;
                    if (rowContent.GetValue(1, 1) == null)
                        break;
                    string partCode     = rowContent.GetValue(1, 1) + "";
                    string type         = rowContent.GetValue(1, 2) + "";
                    string submitPIC    = rowContent.GetValue(1, 3) + "";
                    string IPQC         = rowContent.GetValue(1, 4) + "";
                    string timeSubmit   = rowContent.GetValue(1, 5) + "";
                    string timeStart    = rowContent.GetValue(1, 6) + "";
                    string releaseTime  = rowContent.GetValue(1, 7) + "";
                    string repairTime   = rowContent.GetValue(1, 8) + "";
                    string checkingTime = rowContent.GetValue(1, 9) + "";
                    string totalTime    = rowContent.GetValue(1, 10) + "";
                    string result       = rowContent.GetValue(1, 11) + "";
                    string isRepaired   = rowContent.GetValue(1, 12) + "";
                    string isNotShown   = rowContent.GetValue(1, 13) + "";
                    
                    
                        IPQC ipqc = new IPQC();
                        ipqc.PartNumber = partCode;
                        ipqc.Type = type;
                        ipqc.SubmitPIC = submitPIC;
                        ipqc.IPQC1 = IPQC;
                        ipqc.TimeSubmit = timeSubmit.Replace("H", ":");
                        ipqc.TimeRecive = timeStart.Replace("H", ":"); ;
                        ipqc.ReleaseTime = releaseTime.Replace("H", ":"); ;
                        ipqc.TimeRepair = repairTime.Replace("H", ":"); ;
                        ipqc.CheckingTime = checkingTime.Replace("H", ":"); ;
                        ipqc.TotalTime = totalTime.Replace("H", ":"); ;
                        ipqc.Result = result;
                        ipqc.isRepaired = isRepaired;
                    ipqc.isRepaired = isNotShown;


                        if (ipqc.Result.Trim().Length > 0)
                        {
                            if ((ipqc.TimeRepair != null && ipqc.TimeRepair.Trim().Length > 0))
                            {
                                ipqc.Status = STATUS.REPAIR;
                            }
                            else
                            {
                                ipqc.Status = STATUS.EMPTY;
                            }
                        }
                        else
                        {
                            if (ipqc.TimeRecive.Trim().Length > 0)
                            {
                                ipqc.Status = STATUS.CHECKING;
                            }
                            else
                            {
                                ipqc.Status = STATUS.WAITING;
                            }
                        }

                        lstIPQC.Add(ipqc);
                    
                }

                MyBook.Close(true);
                MyApp.Quit();
                releaseObject(MySheet);
                releaseObject(MyBook);
                releaseObject(MyApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể load dữ liệu");
            }
            lstSource = lstIPQC;
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                obj = null;
            }
            finally
            { GC.Collect(); }

        }

        public void WriteExcelFile()
        {
            try
            {
                Workbook MyBook = null;
                Application MyApp = null;
                Worksheet MySheet = null;
                Range RangeCell = null;

               

                MyApp = new Application();
                MyApp.Visible = false;
                string urlRead = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DATABASE.xlsx");

                MyBook = MyApp.Workbooks.Open(urlRead, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
                MySheet = (Worksheet)MyBook.Sheets[1];

                MySheet.Cells[1, 1] = "PART CODE";
                MySheet.Cells[1, 2] = "TYPE";
                MySheet.Cells[1, 3] = "SUBMIT P.I.C";
                MySheet.Cells[1, 4] = "IPQC";
                MySheet.Cells[1, 5] = "TIMESUBIT";
                MySheet.Cells[1, 6] = "TIME START";
                MySheet.Cells[1, 7] = "RELEASE TIME";
                MySheet.Cells[1, 8] = "REPAIR TIME";
                MySheet.Cells[1, 9] = "CHECKING TIME";
                MySheet.Cells[1, 10] = "TOTAL TIME";
                MySheet.Cells[1, 11] = "RESULT";
                MySheet.Cells[1, 12] = "IS REPAIRED";
                MySheet.Cells[1, 13] = "IS NOT SHOWN";

                MySheet.get_Range("A1", "K1").Font.Bold = true;
                MySheet.get_Range("A1", "K1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                MySheet.get_Range("A1", "K1").EntireColumn.AutoFit();
                
                RangeCell = MySheet.get_Range("A1", "K1");

                for (int i = 0; i < lstSource.Count; i++)
                {
                    try
                    {
                        int currentIndex = i + 2;
                        IPQC ipqc = lstSource[i];
                        MySheet.Cells[currentIndex, 1] = ipqc.PartNumber;
                        MySheet.Cells[currentIndex, 2] = ipqc.Type;
                        MySheet.Cells[currentIndex, 3] = ipqc.SubmitPIC;
                        MySheet.Cells[currentIndex, 4] = ipqc.IPQC1;
                        MySheet.Cells[currentIndex, 5] = ipqc.TimeSubmit.Replace(":", "H");
                        MySheet.Cells[currentIndex, 6] = ipqc.TimeRecive.Replace(":", "H"); 
                        MySheet.Cells[currentIndex, 7] = ipqc.ReleaseTime.Replace(":", "H"); 
                        MySheet.Cells[currentIndex, 8] = ipqc.TimeRepair.Replace(":", "H"); ;
                        MySheet.Cells[currentIndex, 9] = ipqc.CheckingTime.Replace(":", "H"); 
                        MySheet.Cells[currentIndex, 10] = ipqc.TotalTime.Replace(":", "H"); ;
                        MySheet.Cells[currentIndex, 11] = ipqc.Result;
                        MySheet.Cells[currentIndex, 12] = ipqc.isRepaired;
                        MySheet.Cells[currentIndex, 13] = ipqc.isNOTShown;
                    }
                    catch
                    {
                        
                    }
                   
                }
                MyApp.DisplayAlerts = false;
                MyBook.SaveAs(urlRead, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault );
                MyBook.Close(true);
                MyApp.Quit();
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể lưu dữ liệu");
            }
        }

        private void MainWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            StoredExcelFile(lstSource);
        }

        public void StoredExcelFile(List<IPQC> lstIPQC)
        {
            try
            {
                Workbook MyBook = null;
                Application MyApp = null;
                Worksheet MySheet = null;
                Range RangeCell = null;

                MyApp = new Application();
                MyApp.Visible = false;
                DateTime current = DateTime.Now;
                string FileName = current.Year + "_" + current.Month + "_" + current.Day + "_IPQC_CHECKING";
                string url = "D:\\" + FileName;
                if (!System.IO.File.Exists(url))
                {
                    MyBook = MyApp.Workbooks.Add(System.Reflection.Missing.Value);
                    MySheet = MyBook.ActiveSheet;
                }
                else
                {
                    MyBook = MyApp.Workbooks.Open(url, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
                    MySheet = (Worksheet)MyBook.Sheets[1];
                }
               
                MySheet.Cells[1, 1] = "PART CODE";
                MySheet.Cells[1, 2] = "TYPE";
                MySheet.Cells[1, 3] = "SUBMIT P.I.C";
                MySheet.Cells[1, 4] = "IPQC";
                MySheet.Cells[1, 5] = "TIMESUBIT";
                MySheet.Cells[1, 6] = "TIME START";
                MySheet.Cells[1, 7] = "RELEASE TIME";
                MySheet.Cells[1, 8] = "REPAIR TIME";
                MySheet.Cells[1, 9] = "CHECKING TIME";
                MySheet.Cells[1, 10] = "TOTAL TIME";
                MySheet.Cells[1, 11] = "RESULT";

                MySheet.get_Range("A1", "K1").Font.Bold = true;
                MySheet.get_Range("A1", "K1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                MySheet.get_Range("A1", "K1").EntireColumn.AutoFit();

                RangeCell = MySheet.get_Range("A1", "K1");
                RangeCell.Interior.ColorIndex = 36;

                for (int i = 0; i < lstSource.Count; i++)
                {
                    try
                    {
                        int currentIndex = i + 2;
                        IPQC ipqc = lstSource[i];
                        MySheet.Cells[currentIndex, 1] = ipqc.PartNumber;
                        MySheet.Cells[currentIndex, 2] = ipqc.Type;
                        MySheet.Cells[currentIndex, 3] = ipqc.SubmitPIC;
                        MySheet.Cells[currentIndex, 4] = ipqc.IPQC1;
                        MySheet.Cells[currentIndex, 5] = ipqc.TimeSubmit.Replace(":", "H");
                        MySheet.Cells[currentIndex, 6] = ipqc.TimeRecive.Replace(":", "H");
                        MySheet.Cells[currentIndex, 7] = ipqc.ReleaseTime.Replace(":", "H");
                        MySheet.Cells[currentIndex, 8] = ipqc.TimeRepair.Replace(":", "H"); ;
                        MySheet.Cells[currentIndex, 9] = ipqc.CheckingTime.Replace(":", "H");
                        MySheet.Cells[currentIndex, 9].Interior.Color = 16758272;
                        MySheet.Cells[currentIndex, 10] = ipqc.TotalTime.Replace(":", "H");
                        MySheet.Cells[currentIndex, 10].Interior.Color = 16758272;
                        MySheet.Cells[currentIndex, 11] = ipqc.Result;
                        if(ipqc.Result.Trim().Equals(RESULT.NG))
                        {
                            MySheet.Cells[currentIndex, 11].Characters[1, 3].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        Range itemRang = MySheet.get_Range("A" + currentIndex, "K" + currentIndex);
                        itemRang.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                    }
                    catch
                    {

                    }

                }
               
                MySheet.Columns.ColumnWidth = 20;
                RangeCell.Font.Size = 14;
                RangeCell.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                RangeCell.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                RangeCell.Interior.ColorIndex = 36;
                MyApp.DisplayAlerts = false;
                MyBook.SaveAs(url, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault);
                MyBook.Close(true);
                MyApp.Quit();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể Export data dữ liệu");
            }
        }
    }

}
   