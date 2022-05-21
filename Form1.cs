using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using MySql.Conn;
using MySql.Data.MySqlClient;

namespace ArchiveApp
{
    public partial class Form1 : Form
    {
        private bool db_state;
        public static int themeColor = 0;
        private MySqlConnection conn = DBUtils.GetDBConnection();
        public Form1()
        {
            InitializeComponent();
            try
            {
                conn.Open();
                db_state = true;
            }
            catch (MySqlException e) // Если возникают проблемы с БД
            {
                MessageBox.Show(e.Message);
                db_state = false;
                conn.Close();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            User.AuthUser = true;
            if (User.AuthUser == false) // Если пользователь не авторизован
            {
                Form authForm = new Form2();
                authForm.ShowDialog(); //Отобразить окно авторизации
            }
            switch (themeColor)
            {
                case 0:
                {
                    светлаяТемаToolStripMenuItem.Checked = true;
                    this.BackColor = Color.FromName("Window");
                    this.ForeColor = Color.FromName("ControlText");
                    menuStrip1.BackColor = Color.FromName("Window");
                    dataGridView1.BackgroundColor = Color.FromName("Window");
                    dataGridView1.DefaultCellStyle.ForeColor = Color.FromName("ControlText");
                    dataGridView1.DefaultCellStyle.SelectionForeColor = Color.FromName("ControlText");
                    dataGridView1.DefaultCellStyle.BackColor = Color.FromName("Window");
                    dataGridView1.DefaultCellStyle.SelectionBackColor = Color.FromName("GradientActiveCaption");
                    dataGridView1.RowHeadersDefaultCellStyle.ForeColor = Color.FromName("ControlText");
                    dataGridView1.RowHeadersDefaultCellStyle.SelectionForeColor = Color.FromName("ControlLightLight");
                    dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.FromName("Window");
                    dataGridView1.RowHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(15, 15, 15);
                    break;
                }
                case 1:
                {
                    тёмнаяТемаToolStripMenuItem.Checked = true;
                    menuStrip1.BackColor = Color.FromArgb(41, 41, 41);
                    this.BackColor = Color.FromArgb(41, 41, 41);
                    this.ForeColor = Color.FromName("ControlLightLight");
                    dataGridView1.BackgroundColor = Color.FromArgb(41, 41, 41);
                    dataGridView1.DefaultCellStyle.ForeColor = Color.FromName("ControlLightLight");
                    dataGridView1.DefaultCellStyle.SelectionForeColor = Color.FromName("ControlLightLight");
                    dataGridView1.DefaultCellStyle.BackColor = Color.FromArgb(41, 41, 41);
                    dataGridView1.DefaultCellStyle.SelectionBackColor = Color.FromArgb(15, 15, 15);
                    dataGridView1.RowHeadersDefaultCellStyle.ForeColor = Color.FromName("ControlLightLight");
                    dataGridView1.RowHeadersDefaultCellStyle.SelectionForeColor = Color.FromName("ControlLightLight");
                    dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(41, 41, 41);
                    dataGridView1.RowHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(15, 15, 15);
                    break;
                }
                default:
                {
                    светлаяТемаToolStripMenuItem.Checked = true;
                    this.BackColor = Color.FromName("Window");
                    this.ForeColor = Color.FromName("ControlText");
                    break;
                }
            }
            label1.Text = User.Name;
        }
        private void Form1_Shown(object sender, EventArgs e)
        {
            LoadDocs();
        }

        private void ToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Form editForm = new Form3();
            editForm.ShowDialog(); // Отобразить форму редактирования
        }

        private void ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ViewDocument();
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
            DocInf.DocID = Int32.Parse(row.Cells["id"].Value.ToString());
            DocInf.DocTitle = row.Cells["title"].Value.ToString();
        }

        private void добавитьДокументToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddFileDialog.ShowDialog();
        }

        private void AddFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            string fName = AddFileDialog.SafeFileName;
            string directoryPath = Path.GetDirectoryName(AddFileDialog.FileName);
            ConvertDocToPNG(directoryPath, fName);
            
        }

        private void ConvertDocToPNG(string startupPath, string filename1)
        {
            string docPath = Path.Combine(startupPath, filename1);
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Document doc = new Document();
            doc = app.Documents.Open(docPath);
            app.Visible = true;
            doc.ShowGrammaticalErrors = false;
            doc.ShowRevisions = false;
            doc.ShowSpellingErrors = false;
            string query = "INSERT INTO `docs` (id_user, id_city, title, filename) VALUES (@id_user, @id_city, @title, @filename);";
            MySqlCommand cmd = new MySqlCommand(query, conn);// Обращение к БД
            cmd.Parameters.AddWithValue("@id_user", User.IdUser);
            cmd.Parameters.AddWithValue("@id_city", User.IdUserCity);
            cmd.Parameters.AddWithValue("@title", filename1);
            cmd.Parameters.AddWithValue("@filename", filename1);
            cmd.ExecuteNonQuery(); // Отправка запроса
            cmd.CommandText = "SELECT LAST_INSERT_ID();";
            int lastId = Convert.ToInt32(cmd.ExecuteScalar());
            foreach (Window window in doc.Windows)
            {
                foreach (Pane pane in window.Panes)
                {
                    for (var i = 1; i <= pane.Pages.Count; i++)
                    {
                        Page page = null;
                        bool populated = false;
                        while (!populated)
                        {
                            try
                            {
                                page = pane.Pages[i];
                                populated = true;
                            }
                            catch (COMException ex)
                            {
                                Thread.Sleep(1);
                            }
                        }
                        string base64ImageRepresentation = Convert.ToBase64String((byte[])(page.EnhMetaFileBits));
                        //var target = Path.Combine(startupPath + Path.DirectorySeparatorChar, string.Format("{1}_page_{0}", i, filename1.Split('.')[0]));
                        try
                        {
                            //using (var ms = new MemoryStream((byte[])(bits)))
                            //{
                            string qry = "INSERT INTO `doc_pages` (id_doc, page_num, page_file) VALUES (@id_doc, @page_num, @page_file);";
                            MySqlCommand command = new MySqlCommand(qry, conn);// Обращение к БД
                            command.Parameters.AddWithValue("@id_doc", lastId);
                            command.Parameters.AddWithValue("@page_num", i);
                            command.Parameters.AddWithValue("@page_file", base64ImageRepresentation);
                            command.ExecuteNonQuery(); // Отправка запроса
                            //var image = Image.FromStream(ms);
                            //var pngTarget = Path.ChangeExtension(target, "png");
                            ////image.Save(pngTarget, ImageFormat.Png);
                            //}
                        }
                        catch (Exception ex)
                        {
                            doc.Close(false, Type.Missing, Type.Missing);
                            Marshal.ReleaseComObject(doc);
                            doc = null;
                            app.Quit(false, Type.Missing, Type.Missing);
                            Marshal.ReleaseComObject(app);
                            app = null;
                            throw ex;
                        }
                    }
                }
            }
            doc.Close(false, Type.Missing, Type.Missing);
            Marshal.ReleaseComObject(doc);
            doc = null;
            app.Quit(false, Type.Missing, Type.Missing);
            Marshal.ReleaseComObject(app);
            app = null;
            GC.Collect();
        }

        private void ViewDocument()
        {
            if (db_state)
            {
                Form ViewDocument = new ViewDocument();
                ViewDocument.ShowDialog();
            }
            else
            {
                MessageBox.Show("Ошибка! База данных не доступна. Обратитесть к системному администратору.", "Закрыть");
            }
        }

        private void PrintDocument()
        {
            //    
        }

        private void печатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintDocument();
        }

        private void тёмнаяТемаToolStripMenuItem_Click(object sender, EventArgs e)
        {

            RegistryKey myKey = Registry.CurrentUser;
            RegistryKey wKey = myKey.OpenSubKey(@"Software\FileTronic",true);
            if (wKey != null)
            {
                wKey.SetValue("theme", 1);// Устанавливаем переменную темной темы
                themeColor = 1;
                светлаяТемаToolStripMenuItem.Checked = false;
                тёмнаяТемаToolStripMenuItem.Checked = true;
                Form1_Load(null, null);
            }
            myKey.Close();

        }
        private void светлаяТемаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RegistryKey myKey = Registry.CurrentUser;
            RegistryKey wKey = myKey.OpenSubKey(@"Software\FileTronic", true);
            if (wKey != null)
            {
                wKey.SetValue("theme", 0);// Устанавливаем переменную светлой темы
                themeColor = 0;
                тёмнаяТемаToolStripMenuItem.Checked = false;
                светлаяТемаToolStripMenuItem.Checked = true;
                Form1_Load(null, null);
            }
            myKey.Close();
        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadDocs();
        }

        public void LoadDocs()
        {
            if (db_state == true)
            {
                MySqlDataReader dataReader;
                string query = "SELECT * FROM `docs`";
                MySqlCommand cmd = new MySqlCommand(query, conn);// Обращение к БД
                dataReader = cmd.ExecuteReader(); // Отправка запроса
                if (dataReader.HasRows)
                {
                    dataGridView1.Rows.Clear();
                    while (dataReader.Read())
                    {
                        dataGridView1.Rows.Add(
                            dataReader["id"].ToString(),
                            dataReader["title"].ToString(),
                            dataReader["id_city"].ToString(),
                            dataReader["id_user"].ToString(),
                            dataReader["dateReg"].ToString()
                        );
                    }
                }
                dataReader.Close();
            }
            else
            {
                MessageBox.Show("Ошибка! База данных не доступна. Обратитесть к системному администратору.", "Закрыть");
            }
        }



        private void backgroundWorker1_DoWork_1(object sender, DoWorkEventArgs e)
        {
            //
        }

    }
}
