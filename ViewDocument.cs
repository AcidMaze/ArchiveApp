﻿using MySql.Conn;
using MySql.Data.MySqlClient;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace ArchiveApp
{

    public partial class ViewDocument : Form
    {

        private MySqlConnection conn = DBUtils.GetDBConnection();
        public ViewDocument()
        {
            InitializeComponent();
        }

        private void ViewDocument_Load(object sender, EventArgs e)
        {
            this.Text = "Просмотр документа - " + DocInf.DocTitle;
            ShowPage(DocInf.DocID, 1);// Показываем первую страницу при загрузке
         }

        private void pageNum_TextChanged(object sender, EventArgs e)
        {
   
            if (!String.IsNullOrWhiteSpace(pageNum.Text)) // Проверка на пустую строку
            {
                int page = Convert.ToInt32(pageNum.Text);
                ShowPage(DocInf.DocID, page);
            }
        }

        private void ShowPage(int doc_id, int page)
        {
            
            conn.Open();
            MySqlDataReader dataReader;
            string query = "SELECT * FROM `doc_pages` WHERE `id_doc` = '" + doc_id + "' AND `page_num` = '" + page + "' LIMIT 1";
            MySqlCommand cmd = new MySqlCommand(query, conn);// Обращение к БД
            dataReader = cmd.ExecuteReader(); // Отправка запроса
            if (dataReader.HasRows)
            {
                dataReader.Read();
                pictureBox1.Image = Image.FromStream(new MemoryStream(Convert.FromBase64String(dataReader.GetString(3))));
                conn.Close();
            }
            else
            {
                MessageBox.Show("Такой страницы не существует.", "Закрыть");
                dataReader.Close();
                conn.Close();
            }
            conn.Close();
            dataReader.Close();
        }

        private void btn_next_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrWhiteSpace(pageNum.Text)) // Проверка на пустую строку
            {
                pageNum.Text = (Convert.ToInt32(pageNum.Text) + 1).ToString();
            }
        }
        private void btn_back_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrWhiteSpace(pageNum.Text)) // Проверка на пустую строку
            {
                pageNum.Text = (Convert.ToInt32(pageNum.Text) - 1).ToString();
            }
        }

        private void pageNum_KeyPress(object sender, KeyPressEventArgs e)
        {
            string str = e.KeyChar.ToString();
            //Регулярное выражение дял проверки на запращенные символы
            if (Regex.IsMatch(str, @"^[\%\/\\\&\?\,\'\`\;\:\.\!\-\@\№\<\>\+\=\*\~\#,а-я,a-z,А-Я,A-Z]+$"))
            {
                e.Handled = true;
            }
        }
    }
    class DocInf
    {
        static public string DocTitle { get; set; }
        static public int DocID { get; set; }
    }
}
