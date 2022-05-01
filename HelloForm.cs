using System;
using System.Windows.Forms;

namespace ArchiveApp
{
    public partial class HelloForm : Form
    {
        public HelloForm()
        {
            InitializeComponent();
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            Close();
        }
        private void HelloForm_Load(object sender, EventArgs e)
        {
            timer1.Interval = 2000;
            timer1.Enabled = true;
        }
    }
}
