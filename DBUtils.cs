using MySql.Data.MySqlClient;

namespace MySql.Conn
{
    class DBUtils
    {
        public static MySqlConnection GetDBConnection()
        {
            return DBMySQLUtils.GetDBConnection("localhost", 3306, "arch_db", "root", "");
            //return DBMySQLUtils.GetDBConnection("141.8.194.203", 3306, "a0684658_archive_db", "a0684658_archive_db", "WZXPzdPX");
        }
    }
}
