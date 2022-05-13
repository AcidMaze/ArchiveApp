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
//using Aspose.Words;

namespace ArchiveApp
{
    public partial class Form1 : Form
    {
        private bool db_state;
        public static int IdAarch = -1; //Ид строки в БД
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
            //User.AuthUser = true; //ПОЗЖЕ УДАЛИТЬ
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
                    break;
                }
                case 1:
                {
                    тёмнаяТемаToolStripMenuItem.Checked = true;
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

                    toolStrip1.BackColor = Color.FromArgb(41, 41, 41);
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
        }
        private void Form1_Shown(object sender, EventArgs e)
        {
            if (db_state == true)
            {
                MySqlDataReader dataReader;
                string query = "SELECT * FROM `achrive`";
                MySqlCommand cmd = new MySqlCommand(query, conn);// Обращение к БД
                dataReader = cmd.ExecuteReader(); // Отправка запроса
                if (dataReader.HasRows)
                {
                    while (dataReader.Read())
                    {
                        dataGridView1.Rows.Add(
                            dataReader["id"].ToString(),
                            dataReader["title"].ToString(),
                            dataReader["dateReg"].ToString(),
                            dataReader["dateChange"].ToString()
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

        private void SelectStatus(int id_status)
        {
            MySqlDataReader dataReader;
            string query;
            string[] statusText = { "Зарегестрированные", "На утверждении", "На исполнении" };
            if (id_status == 0) query = "SELECT * FROM `achrive`"; // Выбран пункт все документы
            else query = "SELECT * FROM `achrive` WHERE `status` = '" + id_status + "'"; //Отображение по статусу
            MySqlCommand cmdAuth = new MySqlCommand(query, conn);
            dataReader = cmdAuth.ExecuteReader(); // Отправка запроса
            dataGridView1.Rows.Clear();
            if (dataReader.HasRows)
            {
                while (dataReader.Read())
                {
                    dataGridView1.Rows.Add(
                        dataReader["id"].ToString(),
                        dataReader["title"].ToString(),
                        dataReader["dateReg"].ToString(),
                        dataReader["dateChange"].ToString(),
                        dataReader["status"].ToString()
                    );
                }
            }
            dataReader.Close();
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

        private void ToolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            SelectStatus(Convert.ToInt32(e.ClickedItem.Tag));
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
            IdAarch = Int32.Parse(row.Cells["id"].Value.ToString());
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
            var docPath = Path.Combine(startupPath, filename1);
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Document doc = new Document();
            doc = app.Documents.Open(docPath);
            app.Visible = true;
            doc.ShowGrammaticalErrors = false;
            doc.ShowRevisions = false;
            doc.ShowSpellingErrors = false;

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
                        var bits = page.EnhMetaFileBits;
                        var target = Path.Combine(startupPath + Path.DirectorySeparatorChar, string.Format("{1}_page_{0}", i, filename1.Split('.')[0]));
                        try
                        {
                            //using (var ms = new MemoryStream((byte[])(bits)))
                            //{
                                string base64ImageRepresentation = Convert.ToBase64String((byte[])(bits));
                                string query = "INSERT INTO `achrive` (id_user, id_city, title, filename, file) VALUES (@id_user, @id_city, @title, @filename, @file);";
                                MySqlCommand cmd = new MySqlCommand(query, conn);// Обращение к БД
                                cmd.Parameters.AddWithValue("@id_user", User.IdUser);
                                cmd.Parameters.AddWithValue("@id_city", User.IdUserCity);
                                cmd.Parameters.AddWithValue("@title", filename1);
                                cmd.Parameters.AddWithValue("@filename", filename1);
                                cmd.Parameters.AddWithValue("@file", base64ImageRepresentation);
                                cmd.ExecuteNonQuery(); // Отправка запроса
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
                MySqlDataReader dataReader;
                string query = "SELECT * FROM `achrive` WHERE `id` = '" + IdAarch + "'";
                MySqlCommand cmd = new MySqlCommand(query, conn);// Обращение к БД
                dataReader = cmd.ExecuteReader(); // Отправка запроса
                if (dataReader.HasRows)
                {
                    dataReader.Read();
                    DocInf.DocTitle = dataReader.GetString(3);
                    DocInf.DocBase64 = dataReader.GetString(5);
                    Form ViewDocument = new ViewDocument();
                    ViewDocument.ShowDialog();

                }
                else
                {
                    MessageBox.Show("Запись не найдена. Обратитесть к системному администратору.", "Закрыть");
                    dataReader.Close();
                }
                dataReader.Close();

            }
            else
            {
                MessageBox.Show("Ошибка! База данных не доступна. Обратитесть к системному администратору.", "Закрыть");
            }
        }

        private void PrintDocument()
        {
        //    if (db_state)
        //    {
        //        MySqlDataReader dataReader;
        //        string query = "SELECT * FROM `achrive` WHERE `id` = '" + IdAarch + "'";
        //        MySqlCommand cmd = new MySqlCommand(query, conn);// Обращение к БД
        //        dataReader = cmd.ExecuteReader(); // Отправка запроса
        //        if (dataReader.HasRows)
        //        {
        //            dataReader.Read();
        //            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
        //            Document wordDoc = wordApp.Documents.Add();
        //            Range content = wordDoc.Content;
        //            content.Text = Base64Decode(dataReader.GetString(5));
        //            content.Font.Name = dataReader.GetString(6);
        //            content.Font.Size = dataReader.GetFloat(7);
        //            wordApp.PrintPreview = true;
        //            dataReader.Close();
        //            //wordDoc.Close();
        //            //wordApp.Quit();
        //        }
        //        else
        //        {
        //            MessageBox.Show("Запись не найдена. Обратитесть к системному администратору.", "Закрыть");
        //            dataReader.Close();
        //        }
        //    }
        //    else
        //    {
        //        MessageBox.Show("Ошибка! База данных не доступна. Обратитесть к системному администратору.", "Закрыть");
        //    }
        }

        public static string Base64Encode(string plainText)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(plainText));
        }
        public static string Base64Decode(string base64EncodedData)
        {
            return Encoding.UTF8.GetString(Convert.FromBase64String(base64EncodedData));
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
            }
            myKey.Close();
        }
    }
}
