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
    }
}
