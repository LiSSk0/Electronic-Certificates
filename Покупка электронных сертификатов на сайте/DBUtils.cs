using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace Покупка_электронных_сертификатов_на_сайте
{
    class DBUtils
    {

        public static string database = "";

        public static SqlConnection GetDBConnection()
        {
            string datasource = @"apteka-server";

            string username = "";
            string password = "";

            return DBSQLServerUtils.GetDBConnection(datasource, database, username, password);
        }
    }
}
