using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ArchiveApp
{
  
    public partial class ViewDocument : Form
    {

        public ViewDocument()
        {
            InitializeComponent();
        }

        private void ViewDocument_Load(object sender, EventArgs e)
        {
            richTextBox1.Font = new Font(DocInf.docFont, DocInf.docFontsize);
            richTextBox1.Text = DocInf.docText.ToString();
        }
    }
    class DocInf
    {
        static public string docFont { get; set; }
        static public float docFontsize { get; set; }
        static public string docText { get; set; }
    }
}
