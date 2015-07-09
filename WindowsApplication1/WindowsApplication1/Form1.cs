using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel; 

namespace WindowsApplication1
{
    public partial class Form1 : Form
    {
        public string fileName = "dictionary.dat";

        public string convencionalDateNull = "1900-01-01";
        public System.Drawing.Color GeneralHeaderColor = System.Drawing.Color.LightBlue;

        public List<string> CategoryDropDownList = new List<string>();
        public List<string> DataDropDownList = new List<string>();

        public enum category { Distribution, Transmission, Unknown };
        
        public enum data { Case, Bus, BusType, Line, Unknown };

        public void GenerateNewForm(string host, string UserID, string databaseName, string password, int categorystring, int datastring, string Operation)
        {
            category Category = TranslateCategoryID(categorystring);
            data Data = TranslateDataID(datastring);

            if (Category == category.Transmission && Data == data.BusType && Operation == "Insert")
            {
                label60.Text = GetGenericInfoLabel(fileName, "BusType.CaseID");
                label61.Text = GetGenericInfoLabel(fileName, "BusType.ID");
                label62.Text = GetGenericInfoLabel(fileName, "BusType.Description");
                button31.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                button32.Text = GetGenericInfoLabel(fileName, "FormClear");
                panel21.Visible = true;
            }
            if (Category == category.Distribution && Data == data.Bus && Operation == "Query")
            {
                label59.Text = GetGenericInfoLabel(fileName, "GenericItem.Select");
                button29.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                button30.Text = GetGenericInfoLabel(fileName, "FormClear");
                SetDistribuctionItemList(comboBox16);
                panel21.Visible = true;
            }

            if (Category == category.Distribution && Data == data.Bus && Operation == "Update")
            {
                label58.Text = GetGenericInfoLabel(fileName, "FormNotUsed");
                panel20.Visible = true;
            }

            if (Category == category.Distribution && Data == data.Bus && Operation == "Remove")
            {
                SetDistribuctionItemList(comboBox14);
                label56.Text = GetGenericInfoLabel(fileName, "Bus.Case");
                label57.Text = GetGenericInfoLabel(fileName, "Bus.Number");
                button27.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                button28.Text = GetGenericInfoLabel(fileName, "FormClear");
                panel19.Visible = true;
            }

            if (Category == category.Distribution && Data == data.Bus && Operation == "Insert")
            {
                label54.Text = GetGenericInfoLabel(fileName, "GenericItem.Select");
                label55.Text = GetGenericInfoLabel(fileName, "Bus.Number");
                button25.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                button26.Text = GetGenericInfoLabel(fileName, "FormClear");
                SetDistribuctionItemList(comboBox13);
                panel18.Visible = true;
            }

            if (Category == category.Transmission && Data == data.Bus && Operation == "Query")
            {
                label51.Text = GetGenericInfoLabel(fileName, "GenericItem.Select");
                SetTransmissionItemList(comboBox12);
                button23.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                button24.Text = GetGenericInfoLabel(fileName, "FormClear");
                panel17.Visible = true;
            }

            if (Category == category.Transmission && Data == data.Bus && Operation == "Update")
            {
                label43.Text = GetGenericInfoLabel(fileName, "Bus.MinReactivePowerOrVoltage");
                label44.Text = GetGenericInfoLabel(fileName, "Bus.MaxReactivePowerOrVoltage");
                label45.Text = GetGenericInfoLabel(fileName, "Bus.DesiredVoltage");
                label46.Text = GetGenericInfoLabel(fileName, "Bus.VoltageBase");
                label47.Text = GetGenericInfoLabel(fileName, "Bus.Phase");
                label48.Text = GetGenericInfoLabel(fileName, "Bus.Voltage");
                label49.Text = GetGenericInfoLabel(fileName, "Bus.Name");
                label50.Text = GetGenericInfoLabel(fileName, "Bus.SequencialNumber");
                label52.Text = GetGenericInfoLabel(fileName, "Bus.Number");
                label53.Text = GetGenericInfoLabel(fileName, "Bus.Case");
                button21.Text = GetGenericInfoLabel(fileName, "FormClear");
                button22.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                SetTransmissionItemList(comboBox11);
                panel16.Visible = true;
            }

            if (Category == category.Transmission && Data == data.Bus && Operation == "Remove")
            {
                label42.Text = GetGenericInfoLabel(fileName, "GenericItem.Select");
                label41.Text = GetGenericInfoLabel(fileName, "GenericItem.Select");
                button19.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                button20.Text = GetGenericInfoLabel(fileName, "FormClear");
                SetTransmissionItemList(comboBox8);
                panel15.Visible = true;
            }

            if (Category == category.Transmission && Data == data.Bus && Operation == "Insert")
            {
                label31.Text = GetGenericInfoLabel(fileName, "Bus.Case");
                label32.Text = GetGenericInfoLabel(fileName, "Bus.Number");
                label33.Text = GetGenericInfoLabel(fileName, "Bus.SequencialNumber");
                label34.Text = GetGenericInfoLabel(fileName, "Bus.Name");
                label35.Text = GetGenericInfoLabel(fileName, "Bus.Voltage");
                label36.Text = GetGenericInfoLabel(fileName, "Bus.Phase");
                label37.Text = GetGenericInfoLabel(fileName, "Bus.VoltageBase");
                label38.Text = GetGenericInfoLabel(fileName, "Bus.DesiredVoltage");
                label39.Text = GetGenericInfoLabel(fileName, "Bus.MaxReactivePowerOrVoltage");
                label40.Text = GetGenericInfoLabel(fileName, "Bus.MinReactivePowerOrVoltage");
                button17.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                button18.Text = GetGenericInfoLabel(fileName, "FormClear");
                SetTransmissionItemList(comboBox7);
                panel14.Visible = true;
            }

            if (Category == category.Distribution && Data == data.Case && Operation == "Query")
            {
                button16.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                panel13.Visible = true;
            }

            if (Category == category.Distribution && Data == data.Case && Operation == "Update")
            {
                label28.Text = GetGenericInfoLabel(fileName, "GenericItem.Select");
                label29.Text = GetGenericInfoLabel(fileName, "Case.Title");
                label30.Text = GetGenericInfoLabel(fileName, "Case.Description");
                button14.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                button15.Text = GetGenericInfoLabel(fileName, "FormClear");
                SetDistribuctionItemList(comboBox6);
                panel12.Visible = true;
            }
            if (Category == category.Transmission && Data == data.Case && Operation == "Insert")
            {
                label11.Text = GetGenericInfoLabel(fileName, "Case.Title") + ":";
                label12.Text = GetGenericInfoLabel(fileName, "Case.Author") + ":";
                label13.Text = GetGenericInfoLabel(fileName, "Case.Description") + ":";
                label14.Text = GetGenericInfoLabel(fileName, "Case.PowerBase") + ":";
                label15.Text = GetGenericInfoLabel(fileName, "Case.CaseDate") + ":";
                label16.Text = GetGenericInfoLabel(fileName, "Case.PublicationDate") + ":";
                button3.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                button4.Text = GetGenericInfoLabel(fileName, "FormClear");
                panel6.Visible = true;
            }
            if (Category == category.Transmission && Data == data.Case && Operation == "Remove")
            {
                comboBox3.Items.Clear();
                label22.Text = GetGenericInfoLabel(fileName, "GenericItem.Select") + ":";
                CaseTable.Query_Case DatabaseAccess = new CaseTable.Query_Case();
                List<string[]> queryresult = new List<string[]>();
                queryresult = DatabaseAccess.query(null, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                for (int j = 0; j < queryresult.Count; j++)
                {
                    if (queryresult[j][7].Equals("1")) // 1 means transmission, 0 means distribuction
                    {
                        comboBox3.Items.Add(queryresult[j][1] + ", " + queryresult[j][2] + ". " + queryresult[j][3]);
                    }
                }
                button5.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                button2.Text = GetGenericInfoLabel(fileName, "FormClear");
                panel7.Visible = true;
            }

            if (Category == category.Transmission && Data == data.Case && Operation == "Update")
            {
                label24.Text = GetGenericInfoLabel(fileName, "GenericItem.Select") + ":";
                label23.Text = GetGenericInfoLabel(fileName, "Case.Title") + ":";
                label21.Text = GetGenericInfoLabel(fileName, "Case.Author") + ":";
                label20.Text = GetGenericInfoLabel(fileName, "Case.Description") + ":";
                label19.Text = GetGenericInfoLabel(fileName, "Case.PowerBase") + ":";
                label18.Text = GetGenericInfoLabel(fileName, "Case.CaseDate") + ":";
                label17.Text = GetGenericInfoLabel(fileName, "Case.PublicationDate") + ":";
                comboBox4.Items.Clear();
                CaseTable.Query_Case DatabaseAccess = new CaseTable.Query_Case();
                List<string[]> queryresult = new List<string[]>();
                queryresult = DatabaseAccess.query(null, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                for (int j = 0; j < queryresult.Count; j++)
                {
                    if (queryresult[j][7].Equals("1")) // 1 means transmission, 0 means distribuction
                    {
                        comboBox4.Items.Add(queryresult[j][1] + ", " + queryresult[j][2] + ". " + queryresult[j][3]);
                    }
                }
                button7.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                button6.Text = GetGenericInfoLabel(fileName, "FormClear");
                panel8.Visible = true;
            }

            if (Category == category.Transmission && Data == data.Case && Operation == "Query")
            {
                checkBox1.Text = GetGenericInfoLabel(fileName, "Case.Title");
                checkBox2.Text = GetGenericInfoLabel(fileName, "Case.Author");
                checkBox3.Text = GetGenericInfoLabel(fileName, "Case.Description");
                checkBox4.Text = GetGenericInfoLabel(fileName, "Case.PowerBase");
                checkBox5.Text = GetGenericInfoLabel(fileName, "Case.CaseDate");
                checkBox6.Text = GetGenericInfoLabel(fileName, "Case.PublicationDate");
                checkBox1.Checked = true;
                checkBox2.Checked = true;
                checkBox3.Checked = true;
                checkBox4.Checked = true;
                checkBox5.Checked = true;
                checkBox6.Checked = true;
                button8.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                button9.Text = GetGenericInfoLabel(fileName, "FormClear");
                panel9.Visible = true;
            }
            if (Category == category.Distribution && Data == data.Case && Operation == "Insert")
            {
                label25.Text = GetGenericInfoLabel(fileName, "Case.Title");
                label26.Text = GetGenericInfoLabel(fileName, "Case.Description");
                button10.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                button11.Text = GetGenericInfoLabel(fileName, "FormClear");
                panel10.Visible = true;
            }
            if (Category == category.Distribution && Data == data.Case && Operation == "Remove")
            {
                comboBox5.Items.Clear();
                CaseTable.Query_Case DatabaseAccess = new CaseTable.Query_Case();
                List<string[]> matrix = new List<string[]>();
                matrix = DatabaseAccess.query(null, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);

                int MaxRows = matrix.Count;
                int MaxCol = matrix[0].Length;

                for (int i = 0; i < MaxRows; i++)
                {
                    if (matrix[i][7].Equals("0"))
                    {
                        comboBox5.Items.Add(matrix[i][1] + " / " + matrix[i][3]);
                    }
                }
                label27.Text = GetGenericInfoLabel(fileName, "GenericItem.Select");
                button12.Text = GetGenericInfoLabel(fileName, "FormSubmit");
                button13.Text = GetGenericInfoLabel(fileName, "FormClear");
                panel11.Visible = true;
            }
        }

        public category TranslateCategoryID(int categoryID)
        {
            switch (categoryID)
            {
                case 0:
                    return category.Distribution;;
                case 1:
                    return category.Transmission;
            }
            return category.Unknown;
        }

        public data TranslateDataID(int dataID)
        {
            switch (dataID)
            {
                case 0:
                    return data.Case; ;
                case 1:
                    return data.Bus;
                case 2:
                    return data.BusType;
            }
            return data.Unknown;
        }

        private void SetDistribuctionItemList(ComboBox comboBox)
        {
            comboBox.Items.Clear();
            List<string[]> matrix = new List<string[]>();
            matrix = GetDistributionMatrix();
            for (int i = 0; i < matrix.Count; i++)
            {
                comboBox.Items.Add(matrix[i][1] + " / " + matrix[i][3]);
            }
        }

        private List<string[]> GetDistributionMatrix()
        {
            CaseTable.Query_Case databaseAccess = new CaseTable.Query_Case();
            List<string[]> matrix = new List<string[]>();
            matrix = databaseAccess.query(null, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
            matrix = CaseDistribuctionFiltering(matrix);
            return matrix;
        }

        private List<string[]> GetTransmissionMatrix()
        {
            CaseTable.Query_Case databaseAccess = new CaseTable.Query_Case();
            List<string[]> matrix = new List<string[]>();
            matrix = databaseAccess.query(null, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
            matrix = CaseTransmissionFiltering(matrix);
            return matrix;
        }

        private void SetTransmissionItemList(ComboBox comboBox)
        {
            comboBox.Items.Clear();
            List<string[]> matrix = new List<string[]>();
            matrix = GetTransmissionMatrix();
            for (int i = 0; i < matrix.Count; i++)
            {
                comboBox.Items.Add(matrix[i][1] + " / " + matrix[i][3]);
            }
        }

        private List<string[]> CaseDistribuctionFiltering(List<string[]> matrix)
        {
            List<string[]> filteredMatrix = new List<string[]>();
            for (int i = 0; i < matrix.Count; i++)
            {
                if (matrix[i][matrix[0].Length-1].Equals("0"))
                {
                    filteredMatrix.Add(matrix[i]);
                }
            }
            return filteredMatrix;
        }

        private List<string[]> CaseTransmissionFiltering(List<string[]> matrix)
        {
            List<string[]> filteredMatrix = new List<string[]>();
            for (int i = 0; i < matrix.Count; i++)
            {
                if (matrix[i][matrix[0].Length-1].Equals("1"))
                {
                    filteredMatrix.Add(matrix[i]);
                }
            }
            return filteredMatrix;
        }

        public string GetErrorLabel(string fileName, int ErrorCode)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "Error."+ErrorCode;
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetGenericInfoLabel(string fileName, string Key)
        {
            string ErrorMsg = "ERROR.998";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetVerifyInfo(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "VerifyInfo";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public void ShowError(int ErrorCode, string ExtraInfo)
        {
            string Error = GetErrorLabel(fileName, ErrorCode);
            label5.Text = Error + GetVerifyInfo(fileName) + ": " + ExtraInfo;
        }

        public string GetSystemLabel(string fileName, string ErrorMsg, string Key)
        {
            string filePath = Application.StartupPath;
            List<string> Buffer = new List<string>();
            System.IO.StreamReader file = new System.IO.StreamReader(@"" + filePath + "//" + fileName, System.Text.Encoding.UTF7);
            string line;
            while ((line = file.ReadLine()) != null)
            {
                Buffer.Add(line);
            }
            file.Close();
            string[] ReadText = Buffer.ToArray();
            string ans = ErrorMsg;
            string[] buffer;
            int j = 0;
            for (j = 0; j < ReadText.Length; j++)
            {
                if (ReadText[j].StartsWith(Key + " <"))
                {
                    buffer = ReadText[j].Split('<');
                    buffer = buffer[1].Split('>');
                    ans = (string)buffer[0];
                }
            }
            if (ans == ErrorMsg)
            {
                ShowError(998, Key); 
            }
            return ans;
        }

        public string GetSystemTitleLabel()
        {
            string ErrorMsg = "Error: No system title found. Check dictonary.dat file.";
            string Key = "SystemName";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public Form1()
        {
            InitializeComponent();
            this.Text = GetSystemTitleLabel();
            this.Size = new Size(900, 543);
            panel1.Location = new Point(12, 14);
            panel1.Size = new Size(383, 215);
            panel2.Location = new Point(12, 427);
            panel2.Size = new Size(857, 25);
            panel3.Location = new Point(12, 449);
            panel3.Size = new Size(857, 39);
            panel4.Location = new Point(12, 244);
            panel4.Size = new Size(857, 166);
        }


        public string GetCategoryLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "CategoryLabel";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetDistributionLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "CategoryInformation.Distribution";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            CategoryDropDownList.Add(label);
            return label;
        }

        public string GetTransmissionLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "CategoryInformation.Transmission";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            CategoryDropDownList.Add(label);
            return label;
        }

        public string GetDataLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "DataLabel";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetCaseLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "DataInformation.Case";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            DataDropDownList.Add(label);
            return label;
        }

        public string GetBusLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "DataInformation.Bus";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            DataDropDownList.Add(label);
            return label;
        }

        public string GetLineLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "DataInformation.Line";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            DataDropDownList.Add(label);
            return label;
        }

        public string GetOperationLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "OperationLabel";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetInsertLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "OperationInformation.Insert";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetRemoveLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "OperationInformation.Remove";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetUpdateLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "OperationInformation.Update";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetQueryLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "OperationInformation.Query";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetUserMsgLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "UserMsgLabel";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetDBConnectionLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "ConnectionLabel";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetHostLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "Connection.Host";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetUserIDLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "Connection.UserID";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetDatabaseLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "Connection.DatabaseName";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetDatabasePasswordLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "Connection.DatabasePassword";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetRunNoErrorMsg(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "Error.RunMsg";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetRunLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "RunButton";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetStopLabel(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "StopButton";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetHostDefaut(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "Host";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetUserIDDefaut(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "UserID";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetDababaseNameDefaut(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "DatabaseName";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        public string GetCorrectlySet(string fileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "NoError.CorrectlySet";
            string label = GetSystemLabel(fileName, ErrorMsg, Key);
            return label;
        }

        private void SetPanelLocation(int PanelLocationX, int PanelLocationY, int PanelLocationH, int PanelLocationW)
        {
            panel5.Location = new Point(PanelLocationX, PanelLocationY);
            panel5.Size = new Size(PanelLocationH, PanelLocationW);
            panel5.Visible = true;
            panel6.Location = new Point(PanelLocationX, PanelLocationY);
            panel6.Size = new Size(PanelLocationH, PanelLocationW);
            panel7.Location = new Point(PanelLocationX, PanelLocationY);
            panel7.Size = new Size(PanelLocationH, PanelLocationW);
            panel8.Location = new Point(PanelLocationX, PanelLocationY);
            panel8.Size = new Size(PanelLocationH, PanelLocationW);
            panel9.Location = new Point(PanelLocationX, PanelLocationY);
            panel9.Size = new Size(PanelLocationH, PanelLocationW);
            panel10.Location = new Point(PanelLocationX, PanelLocationY);
            panel10.Size = new Size(PanelLocationH, PanelLocationW);
            panel11.Location = new Point(PanelLocationX, PanelLocationY);
            panel11.Size = new Size(PanelLocationH, PanelLocationW);
            panel12.Location = new Point(PanelLocationX, PanelLocationY);
            panel12.Size = new Size(PanelLocationH, PanelLocationW);
            panel13.Location = new Point(PanelLocationX, PanelLocationY);
            panel13.Size = new Size(PanelLocationH, PanelLocationW);
            panel14.Location = new Point(PanelLocationX, PanelLocationY);
            panel14.Size = new Size(PanelLocationH, PanelLocationW);
            panel15.Location = new Point(PanelLocationX, PanelLocationY);
            panel15.Size = new Size(PanelLocationH, PanelLocationW);
            panel16.Location = new Point(PanelLocationX, PanelLocationY);
            panel16.Size = new Size(PanelLocationH, PanelLocationW);
            panel17.Location = new Point(PanelLocationX, PanelLocationY);
            panel17.Size = new Size(PanelLocationH, PanelLocationW);
            panel18.Location = new Point(PanelLocationX, PanelLocationY);
            panel18.Size = new Size(PanelLocationH, PanelLocationW);
            panel19.Location = new Point(PanelLocationX, PanelLocationY);
            panel19.Size = new Size(PanelLocationH, PanelLocationW);
            panel20.Location = new Point(PanelLocationX, PanelLocationY);
            panel20.Size = new Size(PanelLocationH, PanelLocationW);
            panel21.Location = new Point(PanelLocationX, PanelLocationY);
            panel21.Size = new Size(PanelLocationH, PanelLocationW);
            panel22.Location = new Point(PanelLocationX, PanelLocationY);
            panel22.Size = new Size(PanelLocationH, PanelLocationW);
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            // Correct the place of painels. 
            int PanelLocationX = 412;
            int PanelLocationY = 15;
            int PanelLocationH = 457;
            int PanelLocationW = 395;

            // Location of panels with forms.
            SetPanelLocation(PanelLocationX, PanelLocationY, PanelLocationH, PanelLocationW);

            // The first label must be the msg ok for user. 
            // User msg No error
            label4.Text = GetUserMsgLabel(fileName);
            label5.Text = GetRunNoErrorMsg(fileName);

            // Category label
            label1.Text = GetCategoryLabel(fileName) + ":";
            comboBox1.Items.Add(GetDistributionLabel(fileName));
            comboBox1.Items.Add(GetTransmissionLabel(fileName));
            
            // Operation label
            label2.Text = GetOperationLabel(fileName) + ":";
            radioButton1.Text = GetInsertLabel(fileName);
            radioButton2.Text = GetRemoveLabel(fileName);
            radioButton3.Text = GetUpdateLabel(fileName);
            radioButton4.Text = GetQueryLabel(fileName);

            // Data label
            label3.Text = GetDataLabel(fileName) + ":";
            comboBox2.Items.Add(GetCaseLabel(fileName));
            comboBox2.Items.Add(GetBusLabel(fileName));
            comboBox2.Items.Add(GetLineLabel(fileName));

            // Database Connection Information
            label10.Text = GetDBConnectionLabel(fileName) + ":";
            label6.Text = GetHostLabel(fileName) + ":";
            label7.Text = GetUserIDLabel(fileName) + ":";
            label8.Text = GetDatabaseLabel(fileName) + ":";
            label9.Text = GetDatabasePasswordLabel(fileName) + ":";
            
            // Run button
            button1.Text = GetRunLabel(fileName);

            // Defaut DB Configurations  
            string ConfigfileName = "config.ini";
            textBox1.Text = GetHostDefaut(ConfigfileName);
            textBox2.Text = GetUserIDDefaut(ConfigfileName);
            textBox3.Text = GetDababaseNameDefaut(ConfigfileName);


        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (button1.Text != GetRunLabel(fileName))
            {
                // Clear all important panels
                panel5.Visible = true;
                panel6.Visible = false;
                panel7.Visible = false;
                panel8.Visible = false;
                panel9.Visible = false;
                panel10.Visible = false;
                panel11.Visible = false;
                panel12.Visible = false;
                panel13.Visible = false;
                panel14.Visible = false;
                panel15.Visible = false;
                panel16.Visible = false;
                panel17.Visible = false;
                panel18.Visible = false;
                panel19.Visible = false;
                panel20.Visible = false;
                panel21.Visible = false;
                panel22.Visible = false;

                // Clear all important forms.
                textBox1.Enabled = true;
                textBox2.Enabled = true;
                textBox3.Enabled = true;
                maskedTextBox1.Enabled = true;
                comboBox1.Enabled = true;
                //comboBox1.Text = string.Empty;
                comboBox2.Enabled = true;
                //comboBox2.Text = string.Empty;
                radioButton1.Enabled = true;
                //radioButton1.Checked = false;
                radioButton2.Enabled = true;
                //radioButton2.Checked = false;
                radioButton3.Enabled = true;
                //radioButton3.Checked = false;
                radioButton4.Enabled = true;
                //radioButton4.Checked = false;
                button1.Text = GetRunLabel(fileName);
                label5.Text = GetGenericInfoLabel(fileName, "ReturnedSuccess");
            }
            else
            {
                button1.Visible = false;
                bool ErrorFound = false;
                bool validated = false;

                string Category = string.Empty;
                if (ErrorFound == false)
                {
                    for (int j = 0; j < CategoryDropDownList.Count; j++)
                    {
                        if (comboBox1.Text.Equals(CategoryDropDownList[j]) == true)
                        {
                            validated = true;
                        }
                    }
                    if (validated == false)
                    {
                        ErrorFound = true;
                        ShowError(999, label1.Text.Split(':')[0]);
                    }
                    else
                    {
                        Category = comboBox1.Text;
                    }
                }
                validated = false;

                string Data = string.Empty;
                if (ErrorFound == false)
                {
                    for (int j = 0; j < DataDropDownList.Count; j++)
                    {
                        if (comboBox2.Text.Equals(DataDropDownList[j]) == true)
                        {
                            validated = true;
                        }
                    }
                    if (validated == false)
                    {
                        ErrorFound = true;
                        ShowError(999, label3.Text.Split(':')[0]);
                    }
                    else
                    {
                        Data = comboBox2.Text;
                    }
                }
                validated = false;

                string Operation = string.Empty;
                if (ErrorFound == false)
                {
                    if (radioButton1.Checked == false && radioButton2.Checked == false && radioButton3.Checked == false && radioButton4.Checked == false)
                    {
                        ErrorFound = true;
                        ShowError(999, label2.Text);
                    }
                    else
                    {
                        if (radioButton1.Checked == true)
                            Operation = "Insert";
                        if (radioButton2.Checked == true)
                            Operation = "Remove";
                        if (radioButton3.Checked == true)
                            Operation = "Update";
                        if (radioButton4.Checked == true)
                            Operation = "Query";
                    }
                }

                if (ErrorFound == false)
                {
                    Test.CommunicationTest NewTest = new Test.CommunicationTest();
                    string read = NewTest.query(textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                    if (read == "OK")
                    {
                        // 
                        // All correct data:
                        //string host = textBox1.Text;
                        //string UserID = textBox2.Text;
                        //string databasename = textBox3.Text;
                        //string password = maskedTextBox1.Text;
                        //string operation = Operation;
                        //string data = Data;
                        //string category = Category;
                        textBox1.Enabled = false;
                        textBox2.Enabled = false;
                        textBox3.Enabled = false;
                        maskedTextBox1.Enabled = false;
                        comboBox1.Enabled = false;
                        comboBox2.Enabled = false;
                        radioButton1.Enabled = false;
                        radioButton2.Enabled = false;
                        radioButton3.Enabled = false;
                        radioButton4.Enabled = false;
                        button1.Text = GetStopLabel(fileName);
                        panel5.Visible = false;
                        GenerateNewForm(textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text, comboBox1.SelectedIndex, comboBox2.SelectedIndex, Operation);
                        label5.Text = GetGenericInfoLabel(fileName,"NoError.OK");
                    }
                    else if (read.Contains("Unknown database"))
                    {
                        ShowError(994, label8.Text.Split(':')[0]);
                    }
                    else if (read.Contains("Unable to connect to any of the specified MySQL hosts."))
                    {
                        ShowError(995, "MySQL server or " + label6.Text.Split(':')[0]);
                    }
                    else if (read.Contains("Access denied for user"))
                    {
                        ShowError(997, label9.Text.Split(':')[0] + "/" + label7.Text.Split(':')[0]);
                    }
                    else
                    {
                        ShowError(996, "MySQL server");
                    }
                }
                button1.Visible = true;
            }
        }

        private void panel5_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void panel6_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
        }

        private void button3_Click(object sender, EventArgs e)
        {
            bool error = false;
            //label11.Text; a label16.Text tem informação. 
            CaseTable.Class_Case model = new CaseTable.Class_Case();
            model.SystemType = comboBox1.SelectedIndex;
            model.Title = textBox4.Text;
            if (model.Title == "" && error == false) 
            {
                error = true;
                ShowError(993, label11.Text.Split(':')[0]);
            }
            model.Author = textBox5.Text;
            if (model.Author == "" && error == false)
            {
                error = true;
                ShowError(993, label12.Text.Split(':')[0]);
            }
            model.Description = textBox6.Text;
            if (model.Description == "" && error == false)
            {
                error = true;
                ShowError(993, label13.Text.Split(':')[0]);
            }
            try
            {
                model.PowerBase = Convert.ToInt32(textBox7.Text);
            }
            catch
            {
                model.PowerBase = 0;
            }
            if (model.PowerBase == 0 && error == false)
            {
                error = true;
                ShowError(993, label14.Text.Split(':')[0]);
            }
            model.CaseDate = dateTimePicker1.Value.Year + "-" + dateTimePicker1.Value.Month + "-" + dateTimePicker1.Value.Day;
            model.PublicationDate = dateTimePicker2.Value.Year + "-" + dateTimePicker2.Value.Month + "-" + dateTimePicker2.Value.Day;

            if (error == false)
            {
                label5.Text = GetGenericInfoLabel(fileName, "NoError.OK");
                CaseTable.Insert_Case AccessDatabase = new CaseTable.Insert_Case();
                string read = AccessDatabase.insert(model, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                if (read == "OK")
                {
                    label5.Text = GetGenericInfoLabel(fileName, "InsertSuccess");
                    textBox4.Text = string.Empty;
                    textBox5.Text = string.Empty;
                    textBox6.Text = string.Empty;
                    textBox7.Text = string.Empty;
                }
                else
                {
                    ShowError(992, read);
                }
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox4.Text = string.Empty;
            textBox5.Text = string.Empty;
            textBox6.Text = string.Empty;
            textBox7.Text = string.Empty;
        }

        private void panel7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel7_Paint_1(object sender, PaintEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            int selectedvalue = comboBox3.SelectedIndex;
            if (selectedvalue >= 0)
            {
                CaseTable.Query_Case DatabaseCommunicaton01 = new CaseTable.Query_Case();
                List<string[]> matrix = new List<string[]>();
                matrix = DatabaseCommunicaton01.query(null, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);

                int ID = Convert.ToInt32(matrix[selectedvalue][0]);

                CaseTable.Remove_Case DatabaseCommunication02 = new CaseTable.Remove_Case();
                string msgBack = DatabaseCommunication02.remove(ID, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                if (msgBack == "OK")
                {
                    label5.Text = GetGenericInfoLabel(fileName, "RemoveSuccess");
                    comboBox3.Text = string.Empty;
                    comboBox3.Items.Clear();
                    CaseTable.Query_Case DatabaseAccess = new CaseTable.Query_Case();
                    List<string[]> queryresult = new List<string[]>();
                    queryresult = DatabaseAccess.query(null, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                    for (int j = 0; j < queryresult.Count; j++)
                    {
                        if (queryresult[j][7].Equals("1")) // 1 means transmission, 0 means distribuction
                        {
                            comboBox3.Items.Add(queryresult[j][1] + ", " + queryresult[j][2] + ". " + queryresult[j][3]);
                        }
                    }
                }
                else
                {
                    ShowError(992, msgBack);
                }
            }
            else
            {
                ShowError(999, label22.Text.Split(':')[0]);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            comboBox3.Text = string.Empty;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void panel9_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel8_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel8_Paint_1(object sender, PaintEventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Get index of selected index
            int selectedIndex = comboBox4.SelectedIndex;
            if (selectedIndex >= 0)
            {
                CaseTable.Query_Case DatabaseAccess = new CaseTable.Query_Case();
                List<string[]> queryresult = new List<string[]>();
                queryresult = DatabaseAccess.query(null, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                string ID_of_selected_item = queryresult[selectedIndex][0];
                textBox11.Text = queryresult[selectedIndex][1];
                textBox10.Text = queryresult[selectedIndex][2];
                textBox9.Text = queryresult[selectedIndex][3];
                textBox8.Text = queryresult[selectedIndex][4];
                dateTimePicker4.Value = new DateTime(Convert.ToInt32(queryresult[selectedIndex][5].Split('/')[2].Split(null)[0]), Convert.ToInt32(queryresult[selectedIndex][5].Split('/')[1]), Convert.ToInt32(queryresult[selectedIndex][5].Split('/')[0]));
                dateTimePicker3.Value = new DateTime(Convert.ToInt32(queryresult[selectedIndex][6].Split('/')[2].Split(null)[0]), Convert.ToInt32(queryresult[selectedIndex][6].Split('/')[1]), Convert.ToInt32(queryresult[selectedIndex][6].Split('/')[0]));
            }
            else
            {
                ShowError(999, label22.Text.Split(':')[0]);
            }
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker4_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            bool nullerror = false;
            string errormsg = string.Empty;
            // Get ID of selected index
            int selectedIndex = comboBox4.SelectedIndex; 
            if (selectedIndex < 0)
            {
                errormsg = label24.Text.Split(':')[0];
            }
            if (string.IsNullOrEmpty(errormsg))
            {
                
                CaseTable.Query_Case DatabaseAccess = new CaseTable.Query_Case();
                List<string[]> queryresult = new List<string[]>();
                queryresult = DatabaseAccess.query(null, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                string buffer = queryresult[selectedIndex][0];
                int ID_of_selected_item = Convert.ToInt32(buffer);

                CaseTable.Class_Case class_case = new CaseTable.Class_Case();
                class_case.SystemType = comboBox1.SelectedIndex;
                class_case.ID = ID_of_selected_item;
                
                if (string.IsNullOrEmpty(textBox11.Text))
                {
                    errormsg = label23.Text.Split(':')[0];
                }
                class_case.Title = textBox11.Text;

                if (string.IsNullOrEmpty(textBox10.Text))
                {
                    errormsg = label21.Text.Split(':')[0];
                }
                class_case.Author = textBox10.Text;

                if (string.IsNullOrEmpty(textBox9.Text))
                {
                    errormsg = label20.Text.Split(':')[0];
                }
                class_case.Description = textBox9.Text;

                try
                {
                    if (Convert.ToInt32(textBox8.Text) == 0)
                    {
                        nullerror = true;
                        errormsg = label19.Text.Split(':')[0];
                    }
                        class_case.PowerBase = Convert.ToInt32(textBox8.Text);
                }
                catch
                {
                    class_case.PowerBase = 0;
                    errormsg = label19.Text.Split(':')[0];
                }
                class_case.CaseDate = dateTimePicker4.Value.Year + "-" + dateTimePicker4.Value.Month + "-" + dateTimePicker4.Value.Day;
                class_case.PublicationDate = dateTimePicker3.Value.Year + "-" + dateTimePicker3.Value.Month + "-" + dateTimePicker3.Value.Day;

                if (string.IsNullOrEmpty(errormsg))
                {
                    CaseTable.Update_Case DatabaseAccess2 = new CaseTable.Update_Case();
                    string error = DatabaseAccess2.update(class_case, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                    if (error == "OK")
                    {
                        label5.Text = GetGenericInfoLabel(fileName, "UpdateSuccess");
                        textBox11.Text = string.Empty;
                        textBox10.Text = string.Empty;
                        textBox9.Text = string.Empty;
                        textBox8.Text = string.Empty;
                        comboBox4.Text = string.Empty;
                        comboBox4.Items.Clear();
                        CaseTable.Query_Case DatabaseAccess3 = new CaseTable.Query_Case();
                        List<string[]> queryresult2 = new List<string[]>();
                        queryresult2 = DatabaseAccess3.query(null, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                        for (int j = 0; j < queryresult2.Count; j++)
                        {
                            comboBox4.Items.Add(queryresult2[j][1] + ", " + queryresult2[j][2] + ". " + queryresult2[j][3]);
                        }
                    }
                    else
                    {
                        ShowError(992, error);
                    }
                }
                else
                {
                    if (nullerror == true)
                    {
                        ShowError(992, errormsg);
                    }
                    else
                    {
                        ShowError(999, errormsg);
                    }
                }
            }
            else
            {
                if (nullerror == true)
                {
                    ShowError(992, errormsg);
                }
                else
                {
                    ShowError(999, errormsg);
                }
            }
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            comboBox4.Text = string.Empty;
            textBox11.Text = string.Empty;
            textBox10.Text = string.Empty;
            textBox9.Text = string.Empty;
            textBox8.Text = string.Empty;
        }

        private void panel9_Paint_1(object sender, PaintEventArgs e)
        {

        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                ShowError(990, ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            List<bool> filter = new List<bool>();
            filter.Add(checkBox1.Checked);
            filter.Add(checkBox2.Checked);
            filter.Add(checkBox3.Checked);
            filter.Add(checkBox4.Checked);
            filter.Add(checkBox5.Checked);
            filter.Add(checkBox6.Checked);

            bool atLeastOneSelected = false;
            for (int i = 0; i < filter.Count; i++)
            {
                if (filter[i] == true)
                {
                    atLeastOneSelected = true;
                    break;
                }
            }

            if (atLeastOneSelected == true)
            {
                CaseTable.Query_Case DatabaseAccess = new CaseTable.Query_Case();
                List<string[]> matrix = new List<string[]>();
                matrix = DatabaseAccess.query(filter, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);

                int MaxRow = matrix.Count;
                int MaxColumn = matrix[0].Length;

                int kj = 0;
                List<string[]> FilteredMatrix = new List<string[]>();
                for (int k = 0; k < MaxRow; k++)
                {
                    if (matrix[k][MaxColumn - 1].Equals("1")) // it means distribuction
                    {
                        FilteredMatrix.Add(matrix[k]);
                        kj = kj + 1;
                    }
                }

                // try to create excel spreadsheet
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    ShowError(991, "Microsoft Excel");
                }
                else
                {
                    Excel.Workbook xlWorkBook;
                    Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;

                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
                    string filePath = Application.StartupPath;
                    System.Drawing.Color HeaderColor = System.Drawing.Color.LightBlue;

                    int k = 1;
                    if (filter[0] == true)
                    {
                        xlWorkSheet.Cells[1, k] = GetGenericInfoLabel(fileName, "Case.Title");
                        xlWorkSheet.get_Range(xlWorkSheet.Cells[1, k], xlWorkSheet.Cells[1, k]).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.get_Range(xlWorkSheet.Cells[1, k], xlWorkSheet.Cells[1, k]).Interior.Color = System.Drawing.ColorTranslator.ToOle(HeaderColor);
                        k = k + 1;
                    }
                    if (filter[1] == true)
                    {
                        xlWorkSheet.Cells[1, k] = GetGenericInfoLabel(fileName, "Transmission.Case.Author");
                        xlWorkSheet.get_Range(xlWorkSheet.Cells[1, k], xlWorkSheet.Cells[1, k]).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.get_Range(xlWorkSheet.Cells[1, k], xlWorkSheet.Cells[1, k]).Interior.Color = System.Drawing.ColorTranslator.ToOle(HeaderColor);
                        k = k + 1;
                    }
                    if (filter[2] == true)
                    {
                        xlWorkSheet.Cells[1, k] = GetGenericInfoLabel(fileName, "Transmission.Case.Description");
                        xlWorkSheet.get_Range(xlWorkSheet.Cells[1, k], xlWorkSheet.Cells[1, k]).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.get_Range(xlWorkSheet.Cells[1, k], xlWorkSheet.Cells[1, k]).Interior.Color = System.Drawing.ColorTranslator.ToOle(HeaderColor);
                        k = k + 1;
                    }
                    if (filter[3] == true)
                    {
                        xlWorkSheet.Cells[1, k] = GetGenericInfoLabel(fileName, "Transmission.Case.PowerBase");
                        xlWorkSheet.get_Range(xlWorkSheet.Cells[1, k], xlWorkSheet.Cells[1, k]).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.get_Range(xlWorkSheet.Cells[1, k], xlWorkSheet.Cells[1, k]).Interior.Color = System.Drawing.ColorTranslator.ToOle(HeaderColor);
                        k = k + 1;
                    }
                    int posDate1 = 0;
                    if (filter[4] == true)
                    {
                        xlWorkSheet.Cells[1, k] = GetGenericInfoLabel(fileName, "Transmission.Case.CaseDate");
                        xlWorkSheet.get_Range(xlWorkSheet.Cells[1, k], xlWorkSheet.Cells[1, k]).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.get_Range(xlWorkSheet.Cells[1, k], xlWorkSheet.Cells[1, k]).Interior.Color = System.Drawing.ColorTranslator.ToOle(HeaderColor);
                        posDate1 = k;
                        k = k + 1;
                    }
                    int posDate2 = 0;
                    if (filter[5] == true)
                    {
                        xlWorkSheet.Cells[1, k] = GetGenericInfoLabel(fileName, "Transmission.Case.PublicationDate");
                        xlWorkSheet.get_Range(xlWorkSheet.Cells[1, k], xlWorkSheet.Cells[1, k]).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.get_Range(xlWorkSheet.Cells[1, k], xlWorkSheet.Cells[1, k]).Interior.Color = System.Drawing.ColorTranslator.ToOle(HeaderColor);
                        posDate2 = k;
                        k = k + 1;
                    }

                    MaxRow = FilteredMatrix.Count;
                    MaxColumn = FilteredMatrix[0].Length;

                    for (int i = 0; i < MaxRow; i++)
                    {
                        for (int j = 0; j < MaxColumn - 1; j++)
                        {
                            if (posDate1 - 1 == j || posDate2 - 1 == j)
                            {
                                xlWorkSheet.get_Range(xlWorkSheet.Cells[i + 2, j + 1], xlWorkSheet.Cells[i + 2, j + 1]).NumberFormat = "dd/mm/aaaa";
                            }
                            xlWorkSheet.get_Range(xlWorkSheet.Cells[i + 2, j + 1], xlWorkSheet.Cells[i + 2, j + 1]).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[i + 2, j + 1] = FilteredMatrix[i][j];
                        }
                    }
                    xlWorkSheet.Columns.AutoFit();


                    xlApp.Visible = true;
                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);
                }
            }
            else
            {
                ShowError(989, GetGenericInfoLabel(fileName, "GenericItem.Select"));
            }
        }

        private void panel10_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {
            int ErrorCode = 0;
            string ErrorMsg = string.Empty;

            string description = textBox13.Text;
            if (description.Equals(string.Empty))
            {
                ErrorCode = 993;
                ErrorMsg = GetGenericInfoLabel(fileName, "Case.Description");
            }

            string title = textBox12.Text;
            if (title.Equals(string.Empty))
            {
                ErrorCode = 993;
                ErrorMsg = GetGenericInfoLabel(fileName, "Case.Title");
            }

            if (ErrorCode != 0)
            {
                ShowError(ErrorCode, ErrorMsg);
            }
            else
            {
                string insertresult = InsertDistribuctionCaseData(title, description);
                if (!insertresult.Equals("OK"))
                {
                    ShowError(992, insertresult);
                }
                else
                {
                    textBox12.Text = string.Empty;
                    textBox13.Text = string.Empty;
                    label5.Text = GetGenericInfoLabel(fileName, "InsertSuccess");
                }
            }

        }

        private string InsertDistribuctionCaseData(string title, string description)
        {
            CaseTable.Insert_Case DatabaseAccess = new CaseTable.Insert_Case();
            CaseTable.Class_Case classCase = new CaseTable.Class_Case();

            classCase.Title = title;
            classCase.Description = description;
            classCase.Author = string.Empty;
            classCase.PowerBase = 0;
            classCase.CaseDate = convencionalDateNull;
            classCase.PublicationDate = convencionalDateNull;
            classCase.SystemType = 0;

            return DatabaseAccess.insert(classCase, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
        }

        private void textBox12_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {
            textBox12.Text = string.Empty;
            textBox13.Text = string.Empty;
        }

        private void panel11_Paint(object sender, PaintEventArgs e)
        {

        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            comboBox5.Text = string.Empty;
        }

        private string RemoveCaseByID(int SelectedID)
        {
            CaseTable.Remove_Case DatabaseCommunication = new CaseTable.Remove_Case();
            return DatabaseCommunication.remove(SelectedID, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            int selectedValue = comboBox5.SelectedIndex;

            if (selectedValue < 0)
            {
                ShowError(988, label27.Text);
            }
            else
            {
                int SelectedID = 0;
                CaseTable.Query_Case DatabaseAccess = new CaseTable.Query_Case();
                List<string[]> queryresult = new List<string[]>();
                queryresult = DatabaseAccess.query(null, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                int ij = 0;
                for (int j = 0; j < queryresult.Count; j++)
                {
                    if (queryresult[j][7].Equals("0")) // 1 means transmission, 0 means distribuction
                    {
                        if (ij == selectedValue)
                        {
                            SelectedID = Convert.ToInt32(queryresult[j][0]);
                        }
                        ij = ij + 1;
                    }
                }
                
                // Use selectedID for removal event
                string msgReturned = RemoveCaseByID(SelectedID);

                if (!msgReturned.ToLower().Equals("ok"))
                {
                    ShowError(992, "MySQL server");
                }
                else
                {
                    label5.Text = GetGenericInfoLabel(fileName, "RemoveSuccess");
                    comboBox5.Items.Clear();
                    CaseTable.Query_Case DatabaseAccess1 = new CaseTable.Query_Case();
                    List<string[]> matrix1 = new List<string[]>();
                    matrix1 = DatabaseAccess1.query(null, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                    int MaxRows1 = matrix1.Count;
                    int MaxCol1 = matrix1[0].Length;
                    for (int i = 0; i < MaxRows1; i++)
                    {
                        if (matrix1[i][7].Equals("0"))
                        {
                            comboBox5.Items.Add(matrix1[i][1] + " / " + matrix1[i][3]);
                        }
                    }
                    comboBox5.Text = string.Empty;
                }
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            int selectedIndex = comboBox6.SelectedIndex;
            if (selectedIndex < 0)
            {
                ShowError(988, label28.Text);
            }
            else
            {
                CaseTable.Query_Case databaseAccess = new CaseTable.Query_Case();
                List<string[]> matrix = new List<string[]>();
                matrix = databaseAccess.query(null, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                matrix = CaseDistribuctionFiltering(matrix);
                int selectedID = Convert.ToInt32(matrix[selectedIndex][0]);
                CaseTable.Update_Case databaseAccess1 = new CaseTable.Update_Case();
                CaseTable.Class_Case classCase = new CaseTable.Class_Case();
                classCase.ID = selectedID;
                classCase.Title = textBox14.Text;
                classCase.Description = textBox15.Text;
                classCase.Author = matrix[selectedIndex][2];
                classCase.PowerBase = Convert.ToInt32(matrix[selectedIndex][4]);
                classCase.CaseDate = convencionalDateNull;
                classCase.PublicationDate = convencionalDateNull;
                classCase.SystemType = Convert.ToInt32(matrix[selectedIndex][7]);
                string returned = databaseAccess1.update(classCase, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                comboBox6.Text = string.Empty;
                SetDistribuctionItemList(comboBox6);
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedIndex = comboBox6.SelectedIndex;
            if (selectedIndex < 0)
            {
                ShowError(988, label28.Text);
            }
            else
            {
                CaseTable.Query_Case databaseAccess = new CaseTable.Query_Case();
                List<string[]> matrix = new List<string[]>();
                matrix = databaseAccess.query(null, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                matrix = CaseDistribuctionFiltering(matrix);
                textBox14.Text = matrix[selectedIndex][1];
                textBox15.Text = matrix[selectedIndex][3];
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            comboBox6.Text = string.Empty;
            textBox14.Text = string.Empty;
            textBox15.Text = string.Empty;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            CaseTable.Query_Case DatabaseAccess = new CaseTable.Query_Case();
            List<bool> selectedItems = new List<bool>();
            selectedItems.Add(true); // Title
            selectedItems.Add(false); // Author
            selectedItems.Add(true); // Descript. 
            selectedItems.Add(false);
            selectedItems.Add(false);
            selectedItems.Add(false);
            List<string[]> matrix = new List<string[]>();
            matrix = DatabaseAccess.query(selectedItems, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
            matrix = CaseDistribuctionFiltering(matrix);
            List<string> header = new List<string>();
            header.Add(GetGenericInfoLabel(fileName,"Case.Title"));
            header.Add(GetGenericInfoLabel(fileName, "Case.Description"));
            List<string[]> matrixWithoutLastColumn = new List<string[]>();
            string[] Row = new string[matrix[0].Length - 1];
            for (int i = 0; i < matrix.Count; i++)
            {
                matrixWithoutLastColumn.Add(RemoveIndices(matrix[i],matrix[i].Length-1));
            }
            GenerateNewSpreadSheet(matrixWithoutLastColumn, header);

        }

        private string[] RemoveIndices(string[] IndicesArray, int RemoveAt)
        {
            string[] newIndicesArray = new string[IndicesArray.Length - 1];
            int i = 0;
            int j = 0;
            while (i < IndicesArray.Length)
            {
                if (i != RemoveAt)
                {
                    newIndicesArray[j] = IndicesArray[i];
                    j++;
                }
                i++;
            }
            return newIndicesArray;
        }

        private void GenerateNewSpreadSheet(List<string[]> matrix, List<string> header)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                ShowError(991, "Microsoft Excel");
            }
            else
            {
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets[1];
                string filePath = Application.StartupPath;
                System.Drawing.Color HeaderColor = GeneralHeaderColor;
                for (int k = 1; k <= header.Count; k++)
                {
                    xlWorkSheet.Cells[1, k] = header[k-1];
                    xlWorkSheet.get_Range(xlWorkSheet.Cells[1, k], xlWorkSheet.Cells[1, k]).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.get_Range(xlWorkSheet.Cells[1, k], xlWorkSheet.Cells[1, k]).Interior.Color = System.Drawing.ColorTranslator.ToOle(HeaderColor);
                }
                for (int i = 0; i < matrix.Count; i++)
                {
                    for (int j = 0; j < matrix[0].Length; j++)
                    {
                        xlWorkSheet.get_Range(xlWorkSheet.Cells[i + 2, j + 1], xlWorkSheet.Cells[i + 2, j + 1]).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[i + 2, j + 1] = matrix[i][j];
                    }
                }
                xlWorkSheet.Columns.AutoFit();
                xlApp.Visible = true;
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
        }

        private void panel13_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label31_Click(object sender, EventArgs e)
        {

        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {

        }

        private void label32_Click(object sender, EventArgs e)
        {

        }

        private void label33_Click(object sender, EventArgs e)
        {

        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {

        }

        private void label34_Click(object sender, EventArgs e)
        {

        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {

        }

        private void panel14_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button18_Click(object sender, EventArgs e)
        {
            comboBox7.Text = string.Empty;
            textBox16.Text = string.Empty;
            textBox17.Text = string.Empty;
            textBox18.Text = string.Empty;
            textBox19.Text = string.Empty;
            textBox20.Text = string.Empty;
            textBox21.Text = string.Empty;
            textBox22.Text = string.Empty;
            textBox23.Text = string.Empty;
            textBox24.Text = string.Empty;
        }

        private void textBox26_TextChanged(object sender, EventArgs e)
        {

        }

        private bool ValidateAsNotNullText(TextBox texbox, Label label)
        {
            if (texbox.Text.Equals(string.Empty))
            {
                ShowError(993, label.Text);
                return false;
            }
            else
            {
                return true;
            }
        }

        private bool ValidateAsDouble(TextBox texbox, Label label, out double number)
        {
            number = 0;
            try
            {
                number = Convert.ToDouble(texbox.Text.Replace('.',','));
                return true;
            }
            catch
            {
                ShowError(986,label.Text);
                return false;
            }
        }

        private bool ValidateAsInt(TextBox texbox, Label label, out int number)
        {
            number = 0;
            try
            {
                number = Convert.ToInt32(texbox.Text);
                return true;
            }
            catch
            {
                ShowError(987, label.Text);
                return false;
            }
        }

        private bool ValidateAsSelectedfromCombobox(ComboBox combobox, Label label)
        {
            int selectedIndex = combobox.SelectedIndex;
            if (selectedIndex < 0)
            {
                ShowError(989, label31.Text);
                return false;
            }
            else
            {
                return true;
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            List<bool> TotalValidation = new List<bool>();
            // Pre-validation
            double busminpowervoltage = 0;
            TextBox busMinpowervoltage = textBox24;
            Label busminpowervoltageLabel = label40;
            TotalValidation.Add(ValidateAsDouble(busMinpowervoltage, busminpowervoltageLabel, out busminpowervoltage));
            double busmaxpowervoltage = 0;
            TextBox busMaxpowervoltage = textBox23;
            Label busmaxpowervoltageLabel = label39;
            TotalValidation.Add(ValidateAsDouble(busMaxpowervoltage, busmaxpowervoltageLabel, out busmaxpowervoltage));
            double busdesiredvoltage = 0;
            TextBox busDisiredVoltage = textBox22;
            Label busDisiredVoltageLabel = label38;
            TotalValidation.Add(ValidateAsDouble(busDisiredVoltage, busDisiredVoltageLabel, out busdesiredvoltage));
            double busvoltagebase = 0;
            TextBox busVoltageBase = textBox21;
            Label busVoltageBaseLabel = label37;
            TotalValidation.Add(ValidateAsDouble(busVoltageBase, busVoltageBaseLabel, out busvoltagebase));
            double busphase = 0;
            TextBox busPhase = textBox20;
            Label busPhaseLabel = label36;
            TotalValidation.Add(ValidateAsDouble(busPhase, busPhaseLabel, out busphase));
            double busvoltage = 0;
            TextBox busVoltage = textBox19;
            Label busVoltageLabel = label35;
            TotalValidation.Add(ValidateAsDouble(busVoltage, busVoltageLabel, out busvoltage));
            TextBox busName = textBox18;
            Label busNameLabel = label34;
            TotalValidation.Add(ValidateAsNotNullText(busName, busNameLabel));
            int bussequencialnumber = 0;
            TextBox busSequencialNumber = textBox17;
            Label busSequencialNumberLabel = label33;
            TotalValidation.Add(ValidateAsInt(busSequencialNumber, busSequencialNumberLabel, out bussequencialnumber));
            int busnumber = 0;
            TextBox busNumber = textBox16;
            Label busNumberLabel = label32;
            TotalValidation.Add(ValidateAsInt(busNumber, busNumberLabel, out busnumber));
            ComboBox Case = comboBox7;
            Label CaseLabel = label31;
            TotalValidation.Add(ValidateAsSelectedfromCombobox(Case, CaseLabel));
            int validationCount = 0;
            for (int j = 0; j < TotalValidation.Count; j++)
            {
                if (TotalValidation[j] == true)
                {
                    validationCount = validationCount + 1;
                }
            }

            if (validationCount == TotalValidation.Count)
            {
                BusTable.Class_Bus model = new BusTable.Class_Bus();
                model.BusName = busName.Text;
                model.BusNumber = busnumber;
                List<string[]> TransmissionMatrix = new List<string[]>();
                TransmissionMatrix = GetTransmissionMatrix();
                model.CaseID = Convert.ToInt32(TransmissionMatrix[Case.SelectedIndex][0]);
                model.MaxPowerVoltage = busmaxpowervoltage;
                model.MinPowerVoltage = busminpowervoltage;
                model.SequencialBusNumber = bussequencialnumber;
                model.Phase = busphase;
                model.Voltage = busvoltage;
                model.VoltageBase = busvoltagebase;
                model.DesiredVoltage = busdesiredvoltage;
                BusTable.Insert_Bus databaseAccess = new BusTable.Insert_Bus();
                string msg = databaseAccess.insert(model, textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text);
                if (msg.ToLower().Equals("ok"))
                {
                    label5.Text = GetGenericInfoLabel(fileName, "InsertSuccess");
                    textBox16.Text = string.Empty;
                    textBox17.Text = string.Empty;
                    textBox18.Text = string.Empty;
                    textBox19.Text = string.Empty;
                    textBox20.Text = string.Empty;
                    textBox21.Text = string.Empty;
                    textBox22.Text = string.Empty;
                    textBox23.Text = string.Empty;
                    textBox24.Text = string.Empty;
                    comboBox7.Text = string.Empty;
                }
                else
                {
                    ShowError(992, msg);
                }
            }
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string[]> matrix = new List<string[]>();
            matrix = GetTransmissionMatrix();
            int selectedIndex = comboBox8.SelectedIndex;
            int CaseID = Convert.ToInt32(matrix[selectedIndex][0]);

            GeneralDatabaseAccess.Query DatabaseAccess = new GeneralDatabaseAccess.Query();
            List<string[]> queryMatrix = new List<string[]>();
            queryMatrix = DatabaseAccess.query(textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text, "SELECT * FROM power_system_database.bus where `Case ID` = " + CaseID + ";");
            comboBox9.Items.Clear();
            if (queryMatrix.Count != 0)
            {
                for (int i = 0; i < queryMatrix.Count; i++)
                {
                    comboBox9.Items.Add("BusNumber: " + queryMatrix[i][0] + ", SequencialNumber:" + queryMatrix[i][2] + ", BusName: " + queryMatrix[i][3]);
                }
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            comboBox8.Text = string.Empty;
            comboBox9.Text = string.Empty;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            int CaseIDselectedIndex = comboBox8.SelectedIndex;
            int BusNumberselectedIndex = comboBox9.SelectedIndex;
            if ((CaseIDselectedIndex > -1) && (BusNumberselectedIndex > -1))
            {
                List<string[]> matrix = new List<string[]>();
                matrix = GetTransmissionMatrix();
                int CaseID = Convert.ToInt32(matrix[CaseIDselectedIndex][0]);
                GeneralDatabaseAccess.Query DatabaseAccess = new GeneralDatabaseAccess.Query();
                List<string[]> queryMatrix = new List<string[]>();
                queryMatrix = DatabaseAccess.query(textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text, "SELECT * FROM power_system_database.bus where `Case ID` = " + CaseID + ";");
                int BusNumberID = Convert.ToInt32(queryMatrix[BusNumberselectedIndex][0]);

                GeneralDatabaseAccess.Remove DatabaseAccess1 = new GeneralDatabaseAccess.Remove();
                string response = DatabaseAccess1.remove(textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text, "DELETE FROM `power_system_database`.`bus` WHERE `Bus Number`='" + BusNumberID + "' and`case ID`='" + CaseID + "';");
                if (!response.ToLower().Equals("ok"))
                {
                    ShowError(992, response);
                }
                else
                {
                    label5.Text = GetGenericInfoLabel(fileName,"RemoveSuccess");
                    comboBox8.Text = string.Empty;
                    comboBox9.Text = string.Empty;
                    SetTransmissionItemList(comboBox8);
                }
            }
            else
            {
                ShowError(988, GetGenericInfoLabel(fileName, "GenericItem.Select"));
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            textBox25.Text = string.Empty;
            textBox26.Text = string.Empty;
            textBox27.Text = string.Empty;
            textBox28.Text = string.Empty;
            textBox29.Text = string.Empty;
            textBox30.Text = string.Empty;
            textBox31.Text = string.Empty;
            textBox32.Text = string.Empty;
            comboBox10.Text = string.Empty;
            comboBox11.Text = string.Empty;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            int CaseIDselectedIndex = comboBox11.SelectedIndex;
            int BusNumberSelectedIndex = comboBox10.SelectedIndex;

            if ((CaseIDselectedIndex > -1) && (BusNumberSelectedIndex > -1))
            {
                List<string[]> matrix = new List<string[]>();
                matrix = GetTransmissionMatrix();
                int CaseID = Convert.ToInt32(matrix[CaseIDselectedIndex][0]);

                GeneralDatabaseAccess.Query databaseAccess = new GeneralDatabaseAccess.Query();
                List<string[]> BusMatrix = new List<string[]>();
                BusMatrix = databaseAccess.query(textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text, "SELECT * FROM power_system_database.bus WHERE `Case ID` = " + CaseID + ";");
                int busNumber = Convert.ToInt32(BusMatrix[BusNumberSelectedIndex][0]);

                List<bool> TotalValidation = new List<bool>();
                TotalValidation.Add(ValidateAsNotNullText(textBox31, label49));

                int sequencialNumber = 0;
                TotalValidation.Add(ValidateAsInt(textBox32, label50, out sequencialNumber));

                double voltage = 0;
                TotalValidation.Add(ValidateAsDouble(textBox30, label48, out voltage));

                double phase = 0;
                TotalValidation.Add(ValidateAsDouble(textBox29, label47, out phase));

                double voltageBase = 0;
                TotalValidation.Add(ValidateAsDouble(textBox28, label46, out voltageBase));

                double desiredVoltage = 0;
                TotalValidation.Add(ValidateAsDouble(textBox27, label45, out desiredVoltage));

                double maxPower = 0;
                TotalValidation.Add(ValidateAsDouble(textBox26, label44, out maxPower));

                double minPower = 0;
                TotalValidation.Add(ValidateAsDouble(textBox25, label43, out minPower));

                if(CompleteValidation(TotalValidation) == true)
                {
                    string MsgUpdate = "UPDATE `power_system_database`.`bus` SET `Sequencial Number`='" + sequencialNumber + "', `Bus name`='" + textBox31.Text + "', `Voltage`='" + voltage.ToString().Replace(',', '.') + "', `Phase`='" + phase.ToString().Replace(',', '.') + "', `Voltage Base`='" + voltageBase.ToString().Replace(',', '.') + "', `Desired Voltage`='" + desiredVoltage.ToString().Replace(',', '.') + "', `Max Power Voltage`='" + maxPower.ToString().Replace(',', '.') + "', `Min Power Voltage`='" + minPower.ToString().Replace(',', '.') + "' WHERE `Bus Number`='" + busNumber + "' and`case ID`='" + CaseID + "';";
                    GeneralDatabaseAccess.Update databaseAccess1 = new GeneralDatabaseAccess.Update();
                    string error = databaseAccess1.update(textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text, MsgUpdate);
                    if (!error.ToLower().Equals("ok"))
                    {
                        ShowError(992, error);
                    }
                    else
                    {
                        textBox5.Text = GetGenericInfoLabel(fileName, "UpdateSuccess");
                        textBox25.Text = string.Empty;
                        textBox26.Text = string.Empty;
                        textBox27.Text = string.Empty;
                        textBox28.Text = string.Empty;
                        textBox29.Text = string.Empty;
                        textBox30.Text = string.Empty;
                        textBox31.Text = string.Empty;
                        textBox32.Text = string.Empty;
                        comboBox10.Text = string.Empty;
                        comboBox11.Text = string.Empty;
                    }
                }
            }
            else
            {
                ShowError(989, GetGenericInfoLabel(fileName, "Bus.Case") + " / " + GetGenericInfoLabel(fileName, "Bus.Number"));
            }
        }

        private bool CompleteValidation(List<bool> TotalValidation)
        {
            for (int i = 0; i< TotalValidation.Count; i++)
            {
                if (TotalValidation[i] == false)
                {
                    return false;
                }
            }
            return true;
        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string[]> matrix = new List<string[]>();
            matrix = GetTransmissionMatrix();
            int CaseIDselectedIndex = comboBox11.SelectedIndex;
            int CaseID = Convert.ToInt32(matrix[CaseIDselectedIndex][0]);
            GeneralDatabaseAccess.Query databaseAccess = new GeneralDatabaseAccess.Query();
            List<string[]> BusMatrix = new List<string[]>();
            BusMatrix = databaseAccess.query(textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text, "SELECT * FROM power_system_database.bus WHERE `Case ID` = " + CaseID + ";");
            int busNumber = Convert.ToInt32(BusMatrix[comboBox10.SelectedIndex][0]);
            textBox32.Text = Convert.ToString(Convert.ToInt32(BusMatrix[comboBox10.SelectedIndex][2]));
            textBox31.Text = Convert.ToString(BusMatrix[comboBox10.SelectedIndex][3]);
            textBox30.Text = Convert.ToString(Convert.ToDouble(BusMatrix[comboBox10.SelectedIndex][4])).Replace(',','.');
            textBox29.Text = Convert.ToString(Convert.ToDouble(BusMatrix[comboBox10.SelectedIndex][5])).Replace(',', '.');
            textBox28.Text = Convert.ToString(Convert.ToDouble(BusMatrix[comboBox10.SelectedIndex][6])).Replace(',', '.');
            textBox27.Text = Convert.ToString(Convert.ToDouble(BusMatrix[comboBox10.SelectedIndex][7])).Replace(',', '.');
            textBox26.Text = Convert.ToString(Convert.ToDouble(BusMatrix[comboBox10.SelectedIndex][8])).Replace(',', '.');
            textBox25.Text = Convert.ToString(Convert.ToDouble(BusMatrix[comboBox10.SelectedIndex][9])).Replace(',', '.');
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string[]> matrix = new List<string[]>();
            matrix = GetTransmissionMatrix();
            int CaseIDselectedIndex = comboBox11.SelectedIndex;
            int CaseID = Convert.ToInt32(matrix[CaseIDselectedIndex][0]);
            GeneralDatabaseAccess.Query databaseAccess = new GeneralDatabaseAccess.Query();
            List<string[]> BusMatrix = new List<string[]>();
            BusMatrix = databaseAccess.query(textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text, "SELECT * FROM power_system_database.bus WHERE `Case ID` = " + CaseID + ";");
            comboBox10.Items.Clear();
            if (BusMatrix.Count > 0)
            {
                for (int i = 0; i < BusMatrix.Count; i++)
                {
                    comboBox10.Items.Add("BusNumber: " + BusMatrix[i][0] + ", SequencialNumber: " + BusMatrix[i][2] + ", BusName: " + BusMatrix[i][3]);
                }
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            comboBox12.Text = string.Empty;
        }

        private void button23_Click(object sender, EventArgs e)
        {
            int selectedIndex = comboBox12.SelectedIndex;
            if (selectedIndex < 0)
            {
                ShowError(988, GetGenericInfoLabel(fileName, "GenericItem.Select"));
            }
            else
            {
                int CaseID = 0;
                List<string[]> CaseMatrix = new List<string[]>();
                CaseMatrix = GetTransmissionMatrix();
                CaseID = Convert.ToInt32(CaseMatrix[selectedIndex][0]);

                string queryMsg = "SELECT * FROM power_system_database.bus WHERE `case ID` = " + CaseID + ";";
                List<string[]> queryResult = new List<string[]>();
                GeneralDatabaseAccess.Query databaseAccess = new GeneralDatabaseAccess.Query();
                queryResult = databaseAccess.query(textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text, queryMsg);

                List<string> header = new List<string>();
                header.Add(GetGenericInfoLabel(fileName, "Bus.Number"));
                header.Add(GetGenericInfoLabel(fileName, "Bus.Case"));
                header.Add(GetGenericInfoLabel(fileName, "Bus.SequencialNumber"));
                header.Add(GetGenericInfoLabel(fileName, "Bus.Name"));
                header.Add(GetGenericInfoLabel(fileName, "Bus.Voltage"));
                header.Add(GetGenericInfoLabel(fileName, "Bus.Phase"));
                header.Add(GetGenericInfoLabel(fileName, "Bus.VoltageBase"));
                header.Add(GetGenericInfoLabel(fileName, "Bus.DesiredVoltage"));
                header.Add(GetGenericInfoLabel(fileName, "Bus.MaxReactivePowerOrVoltage"));
                header.Add(GetGenericInfoLabel(fileName, "Bus.MinReactivePowerOrVoltage"));
                GenerateNewSpreadSheet(queryResult, header);
            }
        }

        private void panel19_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button25_Click(object sender, EventArgs e)
        {
            int selectedIndex = comboBox13.SelectedIndex;
            List<string[]> DistribuctionMatrix = new List<string[]>();
            DistribuctionMatrix = GetDistributionMatrix();
            int CaseID = Convert.ToInt32(DistribuctionMatrix[selectedIndex][0]);

            int busNumber = 0;
            bool validated = ValidateAsInt(textBox33, label55, out busNumber);
            
            if (validated == true)
            {
                string InsertMsg = "INSERT INTO `power_system_database`.`bus` (`Bus Number`, `case ID`) VALUES ('" + busNumber + "', '" + CaseID + "');";
                GeneralDatabaseAccess.Insert databaseAccess = new GeneralDatabaseAccess.Insert();
                string error = databaseAccess.insert(textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text, InsertMsg);
                if (error.ToLower().Equals("ok"))
                {
                    label5.Text = GetGenericInfoLabel(fileName, "InsertSuccess");
                    comboBox13.Text = string.Empty;
                    textBox33.Text = string.Empty;
                }
                else if ((error.Contains("Duplicate entry")) && (error.Contains("for key 'PRIMARY'")))
                {
                    ShowError(985, GetGenericInfoLabel(fileName, "Bus.Number"));
                }
                else
                {
                    ShowError(992, error);
                }
            }
            else
            {
                ShowError(988, GetGenericInfoLabel(fileName, "Bus.Number"));
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            comboBox13.Text = string.Empty;
            textBox33.Text = string.Empty;
        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedIndex = comboBox14.SelectedIndex;
            comboBox15.Items.Clear();
            if (selectedIndex > -1)
            {
                List<string[]> DistribuctionMatrix = new List<string[]>();
                DistribuctionMatrix = GetDistributionMatrix();
                int CaseID = Convert.ToInt32(DistribuctionMatrix[selectedIndex][0]);

                List<string[]> BusFilterdMatrix = new List<string[]>();
                string QueryMsg = "SELECT * FROM power_system_database.bus WHERE `case ID` = " + CaseID + ";";
                GeneralDatabaseAccess.Query databaseAccess = new GeneralDatabaseAccess.Query();
                BusFilterdMatrix = databaseAccess.query(textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text, QueryMsg);

                for(int i = 0; i < BusFilterdMatrix.Count; i++)
                {
                    comboBox15.Items.Add("Bus Number: " + BusFilterdMatrix[i][0]);
                }

            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            int selectedIndex1 = comboBox14.SelectedIndex;
            int selectedIndex2 = comboBox15.SelectedIndex;
            
            if ((selectedIndex1 > -1)&&(selectedIndex2 > -1))
            {
                List<string[]> DistributionMatrix = new List<string[]>();
                DistributionMatrix = GetDistributionMatrix();
                int CaseID = Convert.ToInt32(DistributionMatrix[selectedIndex1][0]);

                GeneralDatabaseAccess.Query databaseAccess = new GeneralDatabaseAccess.Query();
                List<string[]> BusMatrix = new List<string[]>();
                BusMatrix = databaseAccess.query(textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text, "SELECT * FROM power_system_database.bus WHERE `case ID` = " + CaseID + ";");
                int BusNumber = Convert.ToInt32(BusMatrix[selectedIndex2][0]);

                GeneralDatabaseAccess.Remove databaseAccess1 = new GeneralDatabaseAccess.Remove();
                string error = databaseAccess1.remove(textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text, "DELETE FROM `power_system_database`.`bus` WHERE `Bus Number`='" + BusNumber + "' and`case ID`='" + CaseID + "';");

                if (error.ToLower().Equals("ok"))
                {
                    label5.Text = GetGenericInfoLabel(fileName, "RemoveSuccess");
                    comboBox14.Text = string.Empty;
                    comboBox15.Text = string.Empty;
                    SetDistribuctionItemList(comboBox14);
                }

            }
            else
            {
                ShowError(988, GetGenericInfoLabel(fileName, "GenericItem.Select"));
            }
        }

        private void button28_Click(object sender, EventArgs e)
        {
            comboBox14.Text = string.Empty;
            comboBox15.Text = string.Empty;
        }

        private void button30_Click(object sender, EventArgs e)
        {
            comboBox16.Text = string.Empty;
        }

        private List<string[]> MakeQuery(string queryMsg)
        {
            GeneralDatabaseAccess.Query databaseAccess = new GeneralDatabaseAccess.Query();
            return databaseAccess.query(textBox1.Text, textBox2.Text, textBox3.Text, maskedTextBox1.Text, queryMsg);
        }

        private bool ValidatedQuery(List<string[]> queryResult)
        {
            if (queryResult[0][0].ToLower().Contains("*error*"))
            {
                ShowError(992, queryResult[0][0].Split('*')[2]);
                return false;
            }
            else
            {
                return true;
            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            int indexSelected = comboBox16.SelectedIndex;
            if (indexSelected < 0)
            {
                if (comboBox16.Text.Equals(string.Empty))
                {
                    string queryMsg = "SELECT `Title`, `Bus Number`  FROM power_system_database.bus, power_system_database.power_system_case where power_system_database.bus.`case ID` = power_system_database.power_system_case.ID and `System Type` = 0;";
                    List<string[]> queryResult = new List<string[]>();
                    queryResult = MakeQuery(queryMsg);

                    if (ValidatedQuery(queryResult) == true)
                    {
                        List<string> header = new List<string>();
                        header.Add(GetGenericInfoLabel(fileName, "Bus.Case"));
                        header.Add(GetGenericInfoLabel(fileName, "Bus.Number"));

                        GenerateNewSpreadSheet(queryResult, header);

                        label5.Text = GetGenericInfoLabel(fileName, "QuerySuccess");
                        comboBox16.Text = string.Empty;
                    }
                }
                else
                {
                    ShowError(988, GetGenericInfoLabel(fileName, "Bus.Case"));
                }
            }
            else
            {
                List<string[]> distrMatrix = new List<string[]>();
                distrMatrix = GetDistributionMatrix();
                int caseID = Convert.ToInt32(distrMatrix[indexSelected][0]);

                string queryMsg = "SELECT `Title`, `Bus Number`  FROM power_system_database.bus, power_system_database.power_system_case where power_system_database.bus.`case ID` = power_system_database.power_system_case.ID and `System Type` = 0 and power_system_database.bus.`case ID` = " + caseID + ";";
                List<string[]> queryResult = new List<string[]>();
                queryResult = MakeQuery(queryMsg);

                if (ValidatedQuery(queryResult) == true)
                {
                    List<string> header = new List<string>();
                    header.Add(GetGenericInfoLabel(fileName, "Bus.Case"));
                    header.Add(GetGenericInfoLabel(fileName, "Bus.Number"));

                    GenerateNewSpreadSheet(queryResult, header);

                    label5.Text = GetGenericInfoLabel(fileName, "QuerySuccess");
                    comboBox16.Text = string.Empty;
                }
            }
        }

        private void label61_Click(object sender, EventArgs e)
        {

        }
    }
}