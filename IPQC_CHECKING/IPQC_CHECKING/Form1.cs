using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IPQC_CHECKING
{
    public partial class Checking_IPQC : Form
    {
        public Checking_IPQC()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void Checking_IPQC_AutoSizeChanged(object sender, EventArgs e)
        {
            lbl_TitleSyestem.Left = (this.ClientSize.Width - lbl_TitleSyestem.Size.Width) / 2;
        }

        private void Checking_IPQC_Resize(object sender, EventArgs e)
        {
            lbl_TitleSyestem.Left = (pnl_Header.Width - lbl_TitleSyestem.Size.Width) / 2;
        }

        private void btnCheckingIPQC_Click(object sender, EventArgs e)
        {
            var DisplayForm = new DisplayChecking();
            DisplayForm.Show();
        }
    }
}
