using System;
using System.Drawing;
using System.IO;
using System.Text;
//using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using MySql.Conn;
using MySql.Data.MySqlClient;


namespace ArchiveApp
{
    public partial class Form3 : Form
    {
        private bool db_state;
        private MySqlConnection conn = DBUtils.GetDBConnection();
        public Form3()
        {
            InitializeComponent();
            try
            {
                conn.Open();
                db_state = true;
            }
            catch (MySqlException) // Если возникают проблемы с БД
            {
                db_state = false;
                conn.Close();
            }
        }
        private void Form3_Load(object sender, EventArgs e)
        {
            if (db_state)
            {
                MySqlDataReader dataReader;
                string query = "SELECT * FROM `achrive` WHERE `id` = '" + Form1.IdAarch + "'";
                MySqlCommand cmd = new MySqlCommand(query, conn);// Обращение к БД
                dataReader = cmd.ExecuteReader(); // Отправка запроса
                if (dataReader.HasRows)
                {
                    dataReader.Read();
                    this.Text = "Редактирование: " + dataReader.GetString(3);
                    cueTextBox1.Text = dataReader.GetString(3);
                    dataReader.Close();
                }
                else
                {
                    MessageBox.Show("Запись не найдена. Обратитесть к системному администратору.");
                    dataReader.Close();
                }
            }
            else
            {
                MessageBox.Show("Ошибка! База данных не доступна. Обратитесть к системному администратору.");
            }

        }

        private void ComboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None; // Делаем список не доступным для редактирования
        }
  

























        //private void ReplaceWord(string replace, string text, Document doc)//замена
        //{
        //    var d = doc.Content;
        //    d.Find.ClearFormatting();
        //    d.Find.Execute(FindText: replace, ReplaceWith: text);
        //}

        //public Image Base64ToImage(string base64String)
        //{
        //    // Convert base 64 string to byte[]
        //    byte[] imageBytes = Convert.FromBase64String(base64String);
        //    // Convert byte[] to Image
        //    using (var ms = new MemoryStream(imageBytes, 0, imageBytes.Length))
        //    {
        //        Image image = Image.FromStream(ms, true);
        //        return image;
        //    }
        //}

    }

}
