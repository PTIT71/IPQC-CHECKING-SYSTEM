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
    public partial class DisplayChecking : Form
    {
        public DisplayChecking()
        {
            InitializeComponent();
        }

        private void DisplayChecking_Resize(object sender, EventArgs e)
        {
            lbl_TitleSyestem.Left = (pnl_Header.Width - lbl_TitleSyestem.Size.Width) / 2;
        }
    }
}
