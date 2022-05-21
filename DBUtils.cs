using MySql.Data.MySqlClient;

namespace MySql.Conn
{
    class DBUtils
    {
        public static MySqlConnection GetDBConnection()
        {
            return DBMySQLUtils.GetDBConnection("localhost", 3306, "arch_db", "root", "");
        }
    }
}
