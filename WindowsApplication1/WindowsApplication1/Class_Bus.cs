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
                string insert = string.Empty;
                string busNumber = model.BusNumber.ToString();
                string caseID = model.CaseID.ToString();
                string busSequencialNumber = model.SequencialBusNumber.ToString();
                string busName = model.BusName;
                string voltage = string.Empty;
                try
                {
                    voltage = model.Voltage.ToString().Split(',')[0] + "." + model.Voltage.ToString().Split(',')[1];
                }
                catch
                {
                    voltage = model.Voltage.ToString();
                }
                string phase = string.Empty;
                try
                {
                    phase = model.Phase.ToString().Split(',')[0] + "." + model.Phase.ToString().Split(',')[1];
                }
                catch
                {
                    phase = model.Phase.ToString();
                }

                string voltageBase = string.Empty;
                try
                {
                    voltageBase = model.VoltageBase.ToString().Split(',')[0] + "." + model.VoltageBase.ToString().Split(',')[1];
                }
                catch
                {
                    voltageBase = model.VoltageBase.ToString();
                }

                string desiredVoltage = string.Empty;
                try
                {
                    desiredVoltage = model.DesiredVoltage.ToString().Split(',')[0] + "." + model.DesiredVoltage.ToString().Split(',')[1];
                }
                catch
                {
                    desiredVoltage = model.DesiredVoltage.ToString();
                }

                string maxPower = string.Empty;
                try
                {
                    maxPower = model.MaxPowerVoltage.ToString().Split(',')[0] + "." + model.MaxPowerVoltage.ToString().Split(',')[1];
                }
                catch
                {
                    maxPower = model.MaxPowerVoltage.ToString();
                }

                string minPower = string.Empty;
                try
                {
                    minPower = model.MinPowerVoltage.ToString().Split(',')[0] + "." + model.MinPowerVoltage.ToString().Split(',')[1];
                }
                catch
                {
                    minPower = model.MinPowerVoltage.ToString();
                }
                string buffer = "'" + busNumber + "','" + caseID + "','" + busSequencialNumber + "','" + busName + "','" + voltage + "','" + phase + "','" + voltageBase + "','" + desiredVoltage + "','" + maxPower + "','" + minPower + "'";
                insert = "INSERT INTO `" + databasename + "`.`" + CurrentDatabaseTable + "` (`Bus Number`, `Case ID`, `Sequencial Number`, `Bus Name`, `Voltage`, `Phase`, `Voltage Base`, `Desired Voltage`, `Max Power Voltage`, `Min Power Voltage`) VALUES ("+buffer+");";
                
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
