using System;
using System.Collections.Generic;
using System.Text;
using MySql.Data.MySqlClient;

namespace GeneralDatabaseAccess
{
    class Remove
    {
        private MySqlConnection connection;

        public string remove(string server, string UID, string databasename, string password, string RemoveMsg)
        {
            string path = "SERVER=" + server + ";DATABASE=" + databasename + ";UID=" + UID + ";PASSWORD=" + password + ";";
            try
            {
                connection = new MySqlConnection(path);
                connection.Open();
                string insert = RemoveMsg;
                MySqlCommand comandos = new MySqlCommand(insert, connection);
                comandos.ExecuteNonQuery();
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

    class Query
    {
        private MySqlConnection connection;
        public List<string[]> query(string server, string UID, string databasename, string password, string queryMsg)
        {
            string path = "SERVER=" + server + ";DATABASE=" + databasename + ";UID=" + UID + ";PASSWORD=" + password + ";";
            try
            {
                connection = new MySqlConnection(path);
                connection.Open();
                MySqlCommand Query = new MySqlCommand();
                Query.Connection = connection;
                Query.CommandText = queryMsg;
                MySqlDataReader dtreader = Query.ExecuteReader();
                List<string[]> matrix = new List<string[]>();
                List<string> columns = new List<string>();
                int kj = 0;
                while (dtreader.Read())
                {
                    columns.Clear();
                    for (kj = 0; kj < dtreader.FieldCount; kj++)
                    {
                        columns.Add(dtreader[kj].ToString());
                    }
                    matrix.Add(columns.ToArray());
                }
                connection.Close();
                return matrix;
            }
            catch (Exception ex)
            {
                List<string[]> read2 = new List<string[]>();
                string[] read = new string[0];
                read[0] = ex.Message;
                read2.Add(read);
                return read2;
            }
        }
    }
}
