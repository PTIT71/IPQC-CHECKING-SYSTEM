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
using IPQC_CHECKING_SYSTEM.Model;

namespace IPQC_CHECKING_SYSTEM
{
    public partial class MainWindow : Form
    {
        //nguonof
        List<IPQC> lstSource = null;
        /// <summary>
        /// 
        /// </summary>
        List<IPQC> lstView = new List<IPQC>();
        int index = 0;
        int rowCount = 10;
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
        }
        /// <summary>
        /// Item is removed after 30 minutes. from have result
        /// </summary>
        /// <param name="time"></param>
        /// <returns></returns>
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
            var ipqcEntites = new DB_Entities();
            lstSource = ipqcEntites.IPQCs.ToList();
            for (int i = 0; i < lstSource.Count - 1; i++)
                if (FillterCondition(lstSource[i].ReleaseTime))
                    lstSource.RemoveAt(i);
            ChangeDisplayData();
            foreach (DataGridViewRow row in dgvData.Rows)
                row.Height = (dgvData.Height - 50) / rowCount;            
            Schedule.Schedule.Start(00,34,15,@"d:\test.xlsx");
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
        }
    }
}
