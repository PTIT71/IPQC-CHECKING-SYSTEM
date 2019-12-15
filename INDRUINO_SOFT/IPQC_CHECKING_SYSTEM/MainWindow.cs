using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using IPQC_CHECKING_SYSTEM;
using IPQC_CHECKING_SYSTEM.Common;
using IPQC_CHECKING_SYSTEM.Model;
using IronXL;

namespace IPQC_CHECKING_SYSTEM
{
    public partial class MainWindow : Form
    {
        
        List<IPQC> lstSource = null;
        List<IPQC> lstView = new List<IPQC>();
        
        int index = 0;
        int rowCount = 10;

        private void ChangeDisplayData()
        {
           
            if (lstSource.Count <= rowCount)
            {
                lstView = lstSource;
                index = 0;
                Console.WriteLine("--------");
                for(int iss=0;iss<lstView.Count;iss++)
                {
                    Console.WriteLine(lstView[iss].PartNumber);
                }
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
            lbl_CurrentDate.Text = "Date: " + DateTime.Now.Day + "/" + DateTime.Now.Month + "/" +DateTime.Now.Year;
        }

        private bool FillterCondition(string time)
        {
            if (time == null)
                return true;
            bool result = true;            
            if (TimeSpan.Parse(time) <
                TimeSpan.FromHours(DateTime.Now.Hour) +
                TimeSpan.FromMinutes(DateTime.Now.Minute) +
                TimeSpan.FromMinutes(30))
                result = false;
            return result;
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            List<DataSending> lstData = new List<DataSending>();
            for(int z =0;z <Driver.lstDataSending.Count;z++)
            {
                lstData.Add(Driver.lstDataSending[z]);
            }
            Driver.lstDataSending.Clear();

            if (lstData.Count > 0)
            {
                for (int r = 0; r < lstData.Count; r++)
                {
                    DataSending dataSend = lstData[r];
                    if (lstSource.Count == 0)
                    {
                        IPQC ipItem = new IPQC();
                        ipItem.PartNumber = dataSend.partNumber;
                        ipItem.Type = "Buy off sample";
                        ipItem.TimeSubmit = getCurrentHourAndMinute();
                        ipItem.Status = Common.STATUS.WAITING;
                        // update online time list source
                        lstSource.Add(ipItem);
                    }
                    else
                    {
                        if(searchData(lstData[r].partNumber)>0)
                        {
                            int z = searchData(lstData[r].partNumber);
                            // update online time list source
                            switch (dataSend.value)
                            {
                                case "3":
                                    lstSource[z].TimeRecive = getCurrentHourAndMinute();
                                    lstSource[z].Status = STATUS.CHECKING;
                                    break;
                                case "OK":
                                    {
                                        lstSource[z].ReleaseTime = getCurrentHourAndMinute();
                                        lstSource[z].Result = RESULT.OK;
                                        DateTime date1 = DateTime.Parse(lstSource[z].ReleaseTime);
                                        DateTime date2 = DateTime.Parse(lstSource[z].TimeSubmit);
                                        DateTime date3 = DateTime.Parse(lstSource[z].TimeRecive);
                                        lstSource[z].CheckingTime = (date1.Subtract(date3)).ToString();
                                        lstSource[z].TotalTime = (date1.Subtract(date2)).ToString();
                                    }
                                   
                                    break;
                                case "NG":
                                    {
                                        DateTime date1 = DateTime.Parse(lstSource[z].ReleaseTime);
                                        DateTime date2 = DateTime.Parse(lstSource[z].TimeSubmit);
                                        DateTime date3 = DateTime.Parse(lstSource[z].TimeRecive);
                                        lstSource[z].ReleaseTime = getCurrentHourAndMinute();
                                        lstSource[z].Result = RESULT.NG;
                                        lstSource[z].CheckingTime = (date1.Subtract(date3)).ToString();
                                        lstSource[z].TotalTime = (date1.Subtract(date2)).ToString();
                                    }
                                    
                                    break;
                                case "5":
                                    break;
                                case "4":
                                    break;
                                default:
                                     MessageBox.Show("Value code Error!" + dataSend.value);
                                    break;
                            }
                        }
                        
                        else
                        {
                            IPQC ipItem = new IPQC();
                            ipItem.PartNumber = dataSend.partNumber;
                            ipItem.Type = "Buy off sample";
                            ipItem.TimeSubmit = getCurrentHourAndMinute();
                            ipItem.Status = STATUS.WAITING;
                            lstSource.Add(ipItem);
                            break;
                        }
                        
                    }
                  
                }
                writeFile();
            }
            lstSource = readExcel().ToList();
             ChangeDisplayData();
           
             dgvData.Refresh();
        }

        public int searchData(string partNumber)
        {
            int z = 0;
            for (z = 0; z < lstSource.Count; z++)
            {
                if (lstSource[z].PartNumber.Equals(partNumber))
                {
                    return z;
                }
            }

            return -1;
        }

        public string getCurrentHourAndMinute()
        {
            DateTime date =  DateTime.Now;
            return date.Hour + " : " + date.Minute;
        }

        public void writeFile()
        {
            WorkBook xlsxWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
            xlsxWorkbook.Metadata.Author = "IronXL";

            WorkSheet xlsSheet = xlsxWorkbook.CreateWorkSheet("main_sheet");
            xlsSheet["B2"].Value = "PartNumber";
            xlsSheet["C2"].Value = "Type";
            xlsSheet["D2"].Value = "Submit P.I.C";
            xlsSheet["E2"].Value = "IPQC";
            xlsSheet["F2"].Value = "Time Submit";
            xlsSheet["G2"].Value = "Time Recive";
            xlsSheet["H2"].Value = "Release Time";
            xlsSheet["I2"].Value = "Checking Time";
            xlsSheet["J2"].Value = "Total Time";
            xlsSheet["K2"].Value = "Result";
            List<string> Y = new List<string>();
            Y.Add("B");
            Y.Add("C");
            Y.Add("D");
            Y.Add("E");
            Y.Add("F");
            Y.Add("G");
            Y.Add("H");
            Y.Add("I");
            Y.Add("J");
            Y.Add("K");
            int indexs = 3;
            for (int i = 0; i < lstSource.Count; i++)
            {
                for (int j = 0; j < Y.Count; j++)
                {
                    switch (j)
                    {
                        case 0:
                            xlsSheet[Y[j] + indexs].Value = lstSource[i].PartNumber;
                            break;
                        case 1:
                            xlsSheet[Y[j] + indexs].Value = lstSource[i].Type;
                            break;
                        case 2:
                            xlsSheet[Y[j] + indexs].Value = lstSource[i].SubmitPIC;
                            break;
                        case 3:
                            xlsSheet[Y[j] + indexs].Value = lstSource[i].IPQC1;
                            break;
                        case 4:
                            xlsSheet[Y[j] + indexs].Value = StandarTime(lstSource[i].TimeSubmit);
                            break;
                        case 5:
                            xlsSheet[Y[j] + indexs].Value = StandarTime(lstSource[i].TimeRecive);
                            break;
                        case 6:
                            xlsSheet[Y[j] + indexs].Value = StandarTime(lstSource[i].ReleaseTime);
                            break;
                        case 7:
                            xlsSheet[Y[j] + indexs].Value = StandarTime(lstSource[i].CheckingTime);
                            break;
                        case 8:
                            xlsSheet[Y[j] + indexs].Value = StandarTime(lstSource[i].TotalTime);
                            break;
                        case 9:
                            xlsSheet[Y[j] + indexs].Value = lstSource[i].Result;
                            break;
                    }
                }
                indexs++;
            }
            xlsxWorkbook.SaveAs("DBB.xlsx");
        }

        public List<IPQC> readExcel()
        {
            List<string> Y = new List<string>();
            Y.Add("B");
            Y.Add("C");
            Y.Add("D");
            Y.Add("E");
            Y.Add("F");
            Y.Add("G");
            Y.Add("H");
            Y.Add("I");
            Y.Add("J");
            Y.Add("K");
            WorkBook workbook = WorkBook.Load("DBB.xlsx");
            WorkSheet sheet = workbook.WorkSheets.First();
            List<IPQC> lstIPQC = new List<IPQC>();

            int i = 3;
            while(sheet["B" + i].StringValue.Trim().Count()>0)
            {
                IPQC TempIPQC = new IPQC();
                for (int j = 0; j < Y.Count; j++)
                {
                    switch(j)
                    {
                        case 0:
                            TempIPQC.PartNumber = sheet[Y[j] + i].StringValue.Trim();
                            break;
                        case 1:
                            TempIPQC.Type = sheet[Y[j] + i].StringValue.Trim();
                            break;
                        case 2:
                            TempIPQC.SubmitPIC = sheet[Y[j] + i].StringValue.Trim();
                            break;
                        case 3:
                            TempIPQC.IPQC1 = sheet[Y[j] + i].StringValue.Trim();
                            break;
                        case 4:
                            TempIPQC.TimeSubmit = StandarTime(sheet[Y[j] + i].StringValue.Trim());
                            break;
                        case 5:
                            TempIPQC.TimeRecive = StandarTime(sheet[Y[j] + i].StringValue.Trim());
                            break;
                        case 6:
                            TempIPQC.ReleaseTime = StandarTime(sheet[Y[j] + i].StringValue.Trim());
                            break;
                        case 7:
                            TempIPQC.CheckingTime = StandarTime(sheet[Y[j] + i].StringValue.Trim());
                            break;
                        case 8:
                            TempIPQC.TotalTime = StandarTime(sheet[Y[j] + i].StringValue.Trim());
                            break;
                        case 9:
                            TempIPQC.Result = sheet[Y[j] + i].StringValue.Trim();
                            break;
                    }
                }
                if(TempIPQC.Result.Trim().Length>0)
                {
                    TempIPQC.Status = STATUS.CHECKED;
                }
                else
                {
                    if(TempIPQC.TimeRecive.Trim().Length>0)
                    {
                        TempIPQC.Status = STATUS.CHECKING;
                    }
                    else
                    {
                        TempIPQC.Status = STATUS.WAITING;
                    }
                }
                lstIPQC.Add(TempIPQC);
                i++;
            }
            workbook.Close();
            return lstIPQC;
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
                Rectangle rect = e.CellBounds;
                rect.Width -= 2;
                rect.Height -= 2;
                e.Graphics.DrawRectangle(p, rect);
            }
            e.Handled = true;
        }

        private void dgvData_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            e.CellStyle.Font = new Font("Arial", 25);
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
            if (txtValue.Text.Trim().Length > 0 && txtPartNumber.Text.Trim().Length > 0)
            {
                DataSending dataSend = new DataSending();
                dataSend.partNumber = txtPartNumber.Text.ToString();
                dataSend.value = txtValue.Text.ToString();

                Driver.lstDataSending.Add(dataSend);
            }
        }
    }
}
