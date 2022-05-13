using MySql.Conn;
using MySql.Data.MySqlClient;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ArchiveApp
{

    public partial class ViewDocument : Form
    {

        private bool db_state;
        private MySqlConnection conn = DBUtils.GetDBConnection();
        public ViewDocument()
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

        private void ViewDocument_Load(object sender, EventArgs e)
        {
            this.Text = "Просмотр документа - " + DocInf.DocTitle;
            ShowPage(DocInf.DocID, 1);// Показываем первую страницу при загрузке
         }

        private void pageNum_TextChanged(object sender, EventArgs e)
        {
            int page = Convert.ToInt32(pageNum.Text);
            if (page <= 0) pageNum.Text = "1";
            ShowPage(DocInf.DocID, page);
        }

        private void ShowPage(int doc_id, int page)
        {
            MySqlDataReader dataReader;
            string query = "SELECT * FROM `doc_pages` WHERE `id_doc` = '" + doc_id + "' AND `page_num` = '" + page + "' LIMIT 1";
            MySqlCommand cmd = new MySqlCommand(query, conn);// Обращение к БД
            dataReader = cmd.ExecuteReader(); // Отправка запроса
            if (dataReader.HasRows)
            {
                dataReader.Read();
                pictureBox1.Image = Image.FromStream(new MemoryStream(Convert.FromBase64String(dataReader.GetString(3))));
            }
            else
            {
                MessageBox.Show("Такой страницы не существует.", "Закрыть");
                dataReader.Close();
            }
            dataReader.Close();
        }

        private void btn_next_Click(object sender, EventArgs e)
        {
            pageNum.Text = (Convert.ToInt32(pageNum.Text) + 1).ToString();
        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            pageNum.Text = (Convert.ToInt32(pageNum.Text) - 1).ToString();
        }
    }
    class DocInf
    {
        static public string DocTitle { get; set; }
        static public int DocID { get; set; }
    }
}
