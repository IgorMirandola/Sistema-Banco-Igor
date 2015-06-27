using System;
using System.Collections.Generic;
using System.Text;
using MySql.Data.MySqlClient;

namespace CaseTable
{
    class Insert_Case
    {
        private MySqlConnection connection;

        public string insert(Class_Case model, string server, string UID, string databasename, string password)
        {
            string CurrentDatabaseTable = "power_system_case";
            string path = "SERVER=" + server + ";DATABASE=" + databasename + ";UID=" + UID + ";PASSWORD=" + password + ";";
            try
            {
                connection = new MySqlConnection(path);
                connection.Open();

                string insert = "INSERT INTO `" + databasename + "`.`" + CurrentDatabaseTable + "` (`Title`, `Author`, `Description`, `Power Base`, `Case Date`, `Publication Date`, `System Type`) VALUES ('" + model.Title + "','" + model.Author + "','" + model.Description + "','" + model.PowerBase + "','" + model.CaseDate + "','" + model.PublicationDate + "','" + model.SystemType + "');";
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

    class Remove_Case
    {
        private MySqlConnection connection;

        public string remove(int ID, string server, string UID, string databasename, string password)
        {
            string CurrentDatabaseTable = "power_system_case";
            string path = "SERVER=" + server + ";DATABASE=" + databasename + ";UID=" + UID + ";PASSWORD=" + password + ";";
            try
            {
                connection = new MySqlConnection(path);
                connection.Open();
                //  `power_system_database`.`power_system_case` WHERE `ID`='3';
                string insert = "DELETE FROM `" + databasename + "`.`" + CurrentDatabaseTable + "` WHERE `" + "ID" + "` = '"+ID+"';";
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

    class Update_Case
    {
        private MySqlConnection connection;

        public string update(Class_Case model, string server, string UID, string databasename, string password)
        {
            string CurrentDatabaseTable = "power_system_case";
            string path = "SERVER=" + server + ";DATABASE=" + databasename + ";UID=" + UID + ";PASSWORD=" + password + ";";
            try
            {
                connection = new MySqlConnection(path);
                connection.Open();

                string update = "UPDATE `" + databasename + "`.`" + CurrentDatabaseTable + "` SET `Title`='" + model.Title + "', `Author`='" + model.Author + "', `Description`='" + model.Description + "', `Power Base`='" + model.PowerBase + "', `Case Date`='" + model.CaseDate + "', `Publication Date`='" + model.PublicationDate + "' WHERE `ID`='" + model.ID + "';";
                MySqlCommand comandos = new MySqlCommand(update, connection);
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

    class Query_Case
    {
        private MySqlConnection connection;

        public List<string[]> query(List<bool> selectedItems, string server, string UID, string databasename, string password)
        {
            string CurrentDatabaseTable = "power_system_case";
            string path = "SERVER=" + server + ";DATABASE=" + databasename + ";UID=" + UID + ";PASSWORD=" + password + ";";

            if (selectedItems == null) // It means that no filter is necessary
            {
                try
                {
                    connection = new MySqlConnection(path);
                    connection.Open();

                    // Query

                    MySqlCommand Query = new MySqlCommand();
                    Query.Connection = connection;
                    Query.CommandText = "SELECT * FROM `" + databasename + "`.`" + CurrentDatabaseTable + "`;";
                    MySqlDataReader dtreader = Query.ExecuteReader();

                    List<string[]> matrix = new List<string[]>();
                    List<string> columns = new List<string>();
                    int kj = 0;

                    while (dtreader.Read())//Enquanto existir dados no select
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
            else
            {
                try
                {
                    connection = new MySqlConnection(path);
                    connection.Open();

                    // Query

                    MySqlCommand Query = new MySqlCommand();
                    Query.Connection = connection;
                    string filter = string.Empty;
                    bool isFirst = true;
                    if (selectedItems[0] == true)
                    {
                        if (isFirst == true)
                        {
                            filter = filter + "`Title`";
                            isFirst = false;
                        }
                        else
                        {
                            filter = filter + ", `Title`";
                        }
                    }
                    if (selectedItems[1] == true)
                    {
                        if (isFirst == true)
                        {
                            filter = filter + "`Author`";
                            isFirst = false;
                        }
                        else
                        {
                            filter = filter + ", `Author`";
                        }
                    }
                    if (selectedItems[2] == true)
                    {
                        if (isFirst == true)
                        {
                            filter = filter + "`Description`";
                            isFirst = false;
                        }
                        else
                        {
                            filter = filter + ", `Description`";
                        }
                    }
                    if (selectedItems[3] == true)
                    {
                        if (isFirst == true)
                        {
                            filter = filter + "`Power Base`";
                            isFirst = false;
                        }
                        else
                        {
                            filter = filter + ", `Power Base`";
                        }
                    }
                    if (selectedItems[4] == true)
                    {
                        if (isFirst == true)
                        {
                            filter = filter + "`Case Date`";
                            isFirst = false;
                        }
                        else
                        {
                            filter = filter + ", `Case Date`";
                        }
                    }
                    if (selectedItems[5] == true)
                    {
                        if (isFirst == true)
                        {
                            filter = filter + "`Publication Date`";
                            isFirst = false;
                        }
                        else
                        {
                            filter = filter + ", `Publication Date`";
                        }
                    }
                    if (isFirst == true)
                    {
                        filter = filter + "`System Type`";
                    }
                    else
                    {
                        filter = filter + ", `System Type`";
                    }
                    Query.CommandText = "SELECT " + filter + " FROM `" + databasename + "`.`" + CurrentDatabaseTable + "`;";
                    
                    MySqlDataReader dtreader = Query.ExecuteReader();

                    List<string[]> matrix = new List<string[]>();
                    List<string> columns = new List<string>();
                    int kj = 0;

                    while (dtreader.Read())//Enquanto existir dados no select
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

    

    class Class_Case
    {
        private int nID;
        private string nTitle;
        private string nAuthor;
        private string nDescription;
        private int nPowerBase;
        private string nCaseDate;
        private string nPublicationDate;
        private int nSystemType;

        public int ID
        {
            get { return nID; }
            set { nID = value; }
        }

        public string Title
        {
            get { return nTitle; }
            set { nTitle = value; }
        }

        public string Author
        {
            get { return nAuthor; }
            set { nAuthor = value; }
        }

        public string Description
        {
            get { return nDescription; }
            set { nDescription = value; }
        }

        public int PowerBase
        {
            get { return nPowerBase; }
            set { nPowerBase = value; }
        }

        public string CaseDate
        {
            get { return nCaseDate; }
            set { nCaseDate = value; }
        }

        public string PublicationDate
        {
            get { return nPublicationDate; }
            set { nPublicationDate = value; }
        }

        public int SystemType
        {
            get { return nSystemType; }
            set { nSystemType = value; }
        }
    }
}
