﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace MySql.Conn
{
    class DBUtils
    {
        public static MySqlConnection GetDBConnection()
        {
            string host = "localhost";
            int port = 3306;
            string database = "arch_db";
            string username = "root";
            string password = "";
            return DBMySQLUtils.GetDBConnection(host, port, database, username, password);
        }

    }
}
