using System;
using System.Collections.Generic;
using System.Text;
using MySql.Data.MySqlClient;

namespace Test
{
    class CommunicationTest
    {
        private MySqlConnection connection;

        public string query(string server, string UID, string databasename, string password)
        {
            string path = "SERVER=" + server + ";DATABASE=" + databasename + ";UID=" + UID + ";PASSWORD=" + password + ";";
            try
            {
                connection = new MySqlConnection(path);
                connection.Open();
                connection.Close();
                return "OK";
            }
            catch (Exception ex)
            {
#pragma warning disable
                string read = string.Empty;
#pragma warning restore
                read = ex.Message;
                return read;
            }
        }
    }
}