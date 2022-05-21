using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Conn;
using MySql.Data.MySqlClient;

namespace ArchiveApp
{
    public partial class Form2 : Form
    {
        private bool db_state;
        private MySqlConnection conn = DBUtils.GetDBConnection();
        public Form2()
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
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (db_state)
            {
                MySqlDataReader dataReader;
                string cmdText = "SELECT * FROM `users` WHERE login = '" + cueTextBox1.Text + "' AND password = '" + cueTextBox2.Text + "' LIMIT 1";
                MySqlCommand cmdAuth = new MySqlCommand(cmdText, conn);
                dataReader = cmdAuth.ExecuteReader(); // Отправка запроса
                if (dataReader.HasRows)
                {
                    dataReader.Read();
                    User.IdUser = dataReader.GetInt32(0);
                    User.Name = dataReader.GetString(1);
                    User.IdUserCity = dataReader.GetInt32(4);
                    User.AuthUser = true; // Пользователь авторизован
                    dataReader.Close();
                    this.Close();
                    conn.Close();
                }
                else
                {
                    MessageBox.Show("Вы ввели неверный логин или пароль!");
                    dataReader.Close();
                }
            }
            else
            {
                MessageBox.Show("Ошибка! База данных не доступна. Обратитесть к системному администратору.");
            }
        }

        private void cueTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form2_Load(object sender, EventArgs e)
        {
            switch (Form1.themeColor)
            {
                case 0:
                {
                    this.BackColor = Color.FromName("Window");
                    button1.BackColor = Color.FromName("Window");
                    button1.ForeColor = Color.FromName("ControlText");
                    break;
                }
                case 1:
                {
                    this.BackColor = Color.FromArgb(41, 41, 41);
                    button1.BackColor = Color.FromArgb(41, 41, 41);
                    button1.ForeColor = Color.FromName("ControlLightLight");
                    break;
                }
                default:
                {
                    this.BackColor = Color.FromName("Window");
                    button1.BackColor = Color.FromName("Window");
                    button1.ForeColor = Color.FromName("ControlText");
                    break;
                }
            }
        }
    }
}
