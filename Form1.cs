using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.ServiceModel.Dispatcher;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using MySql.Conn;
using MySql.Data.MySqlClient;

namespace ArchiveApp
{
    public partial class Form1 : Form
    {
        private bool db_state;
        public static int IdAarch = -1;
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
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Object fileName = AddFileDialog.FileName;
            Document wordDoc = wordApp.Documents.Open(ref fileName);
            Range content = wordDoc.Content;
            string fileText = content.Text;
            string font = content.Font.Name;
            float fontSize = content.Font.Size;
            if (fontSize > 200 || fontSize == 0) fontSize = 14;
            if (font == "" || font == null) font = "Times New Roman";
            string name = AddFileDialog.SafeFileName;//Имя файла без пути
            string query = "INSERT INTO `achrive` (id_user, title, filename, font, font_size, file) VALUES (@id_user, @id_city, @title, @name, @font, @fontsize, @file);";
            MySqlCommand cmd = new MySqlCommand(query, conn);// Обращение к БД
            cmd.Parameters.AddWithValue("@id_user", User.IdUser);
            cmd.Parameters.AddWithValue("@id_city", User.IdUserCity);
            cmd.Parameters.AddWithValue("@title", name);
            cmd.Parameters.AddWithValue("@name", name);
            cmd.Parameters.AddWithValue("@font", font);
            cmd.Parameters.AddWithValue("@fontsize", fontSize);
            cmd.Parameters.AddWithValue("@file", Base64Encode(fileText));
            cmd.ExecuteNonQuery(); // Отправка запроса
            wordDoc.Close();
            wordApp.Quit();
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
                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                    //Document wordDoc = wordApp.Documents.Open(TemplateFilename);
                    Document wordDoc = wordApp.Documents.Add();
                    Range content = wordDoc.Content;
                    content.Text = Base64Decode(dataReader.GetString(5));
                    content.Font.Name = dataReader.GetString(6);
                    content.Font.Size = dataReader.GetFloat(7);
                    wordApp.Visible = true;
                    //wordDoc.Close();
                    //wordApp.Quit();

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
            if (db_state)
            {
                MySqlDataReader dataReader;
                string query = "SELECT * FROM `achrive` WHERE `id` = '" + IdAarch + "'";
                MySqlCommand cmd = new MySqlCommand(query, conn);// Обращение к БД
                dataReader = cmd.ExecuteReader(); // Отправка запроса
                if (dataReader.HasRows)
                {
                    dataReader.Read();
                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                    Document wordDoc = wordApp.Documents.Add();
                    Range content = wordDoc.Content;
                    content.Text = Base64Decode(dataReader.GetString(5));
                    content.Font.Name = dataReader.GetString(6);
                    content.Font.Size = dataReader.GetFloat(7);
                    wordApp.PrintPreview = true;

                    dataReader.Close();
                    //wordDoc.Close();
                    //wordApp.Quit();
                }
                else
                {
                    MessageBox.Show("Запись не найдена. Обратитесть к системному администратору.", "Закрыть");
                    dataReader.Close();
                }
            }
            else
            {
                MessageBox.Show("Ошибка! База данных не доступна. Обратитесть к системному администратору.", "Закрыть");
            }
        }

        public static string Base64Encode(string plainText)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(plainText));
        }
        public static string Base64Decode(string base64EncodedData)
        {
            return Encoding.UTF8.GetString(Convert.FromBase64String(base64EncodedData));
        }

        public void ConvertWordToImg(Object filename1, Microsoft.Office.Interop.Word.Application app)
        {
            
        }

        private void печатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintDocument();
        }
    }
}
