using System;
using System.Drawing;
using System.IO;
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
            pictureBox1.Image = Image.FromStream(new MemoryStream(Convert.FromBase64String(DocInf.DocBase64)));
            this.Text = "Просмотр документа - " + DocInf.DocTitle;
         }
    }
    class DocInf
    {
        static public string DocBase64 { get; set; }
        static public string DocTitle { get; set; }
    }
}
