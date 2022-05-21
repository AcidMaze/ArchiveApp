using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;

namespace ArchiveApp
{
    public partial class HelloForm : Form
    {
        public HelloForm()
        {
            InitializeComponent();
            //-------- Читаем реестре при запуске программы -------
            RegistryKey myKey = Registry.CurrentUser; //Открываем рестр текущего пользователя для записи
            RegistryKey wKey = myKey.OpenSubKey(@"Software\FileTronic");
            if (wKey != null)
            {
                int theme = Convert.ToInt32(wKey.GetValue("theme"));
                Form1.themeColor = theme;
                if (theme == 0)
                {
                    this.BackColor = Color.FromName("Window");
                    pictureBox1.Image = Properties.Resources.load_anim_2;
                }
                else
                {
                    pictureBox1.Image = Properties.Resources.load_anim_1;
                }
                myKey.Close();
            }
            else
            {
                RegistryKey heyKey = myKey.OpenSubKey("Software", true);
                RegistryKey newKey = heyKey.CreateSubKey("FileTronic");
                newKey.SetValue("theme", 0);
                myKey.Close();
            }
            //------------------------------------------------------------------
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            Close();
        }
        private void HelloForm_Load(object sender, EventArgs e)
        {
            timer1.Interval = 2000;
            timer1.Enabled = true;
        }

    }
}
