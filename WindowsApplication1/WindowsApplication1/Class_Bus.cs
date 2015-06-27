using System;
using System.Collections.Generic;
using System.Text;
using MySql.Data.MySqlClient;

namespace BusTable
{
    class Insert_Bus
    {
        private MySqlConnection connection;

        public string insert(Class_Bus model, string server, string UID, string databasename, string password)
        {
            string CurrentDatabaseTable = "bus";
            string path = "SERVER=" + server + ";DATABASE=" + databasename + ";UID=" + UID + ";PASSWORD=" + password + ";";
            try
            {
                connection = new MySqlConnection(path);
                connection.Open();

                string insert = "INSERT INTO `" + databasename + "`.`" + CurrentDatabaseTable + "` (`Bus Number`, `Case ID`, `Sequencial Number`, `Bus Name`, `Voltage`, `Phase`, `Voltage Base`, `Desired Voltage`, `Max Power Voltage`, `Min Power Voltage`) VALUES ('" + model.BusNumber + "','" + model.CaseID + "','" + model.SequencialBusNumber.ToString() + "','" + model.BusName + "','" + model.Voltage.ToString() + model.Phase.ToString() + "','" + model.VoltageBase + "','" + "','" + model.DesiredVoltage + "','" + model.MaxPowerVoltage + "','" + model.MinPowerVoltage + "');";
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

    class Class_Bus
    {
        private int nCaseID;
        private int nBusNumber;
        private int nBusSequencialNumber;
        private string nBusName;
        private double nVoltage;
        private double nPhase;
        private double nVoltageBase;
        private double nDesiredVoltage;
        private double nMaxPowerVoltage;
        private double nMinPowerVoltage;

        public int CaseID
        {
            get { return nCaseID; }
            set { nCaseID = value; }
        }

        public int BusNumber
        {
            get { return nBusNumber; }
            set { nBusNumber = value; }
        }

        public int SequencialBusNumber
        {
            get { return nBusSequencialNumber; }
            set { nBusSequencialNumber = value; }
        }

        public string BusName
        {
            get { return nBusName; }
            set { nBusName = value; }
        }

        public double Voltage
        {
            get { return nVoltage; }
            set { nVoltage = value; }
        }

        public double Phase
        {
            get { return nPhase; }
            set { nPhase = value; }
        }

        public double VoltageBase
        {
            get { return nVoltageBase; }
            set { nVoltageBase = value; }
        }

        public double DesiredVoltage
        {
            get { return nDesiredVoltage; }
            set { nDesiredVoltage = value; }
        }

        public double MaxPowerVoltage
        {
            get { return nMaxPowerVoltage; }
            set { nMaxPowerVoltage = value; }
        }

        public double MinPowerVoltage
        {
            get { return nMinPowerVoltage; }
            set { nMinPowerVoltage = value; }
        }
    }
}
