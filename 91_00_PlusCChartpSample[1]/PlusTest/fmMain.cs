using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlusTest
{
    public partial class fmMain : Form
    {
        public fmMain()
        {
            InitializeComponent();
        }

        private void m_btnForm1_Click(object sender, EventArgs e)
        {
            var form1 = new fmStock();
            form1.Show();
        }

        private void m_btnForm2_Click(object sender, EventArgs e)
        {
            var form1 = new fmFuture();
            form1.Show();
        }

        private void m_btnForm3_Click(object sender, EventArgs e)
        {
            var form1 = new fmOption();
            form1.Show();
        }
        private void m_btnForm4_Click(object sender, EventArgs e)
        {
            var form1 = new fmSFuture();
            form1.Show();
        }
        private void m_btnForm5_Click(object sender, EventArgs e)
        {
            var form1 = new fmVolTop();
            form1.Show();
        }
    }
}
