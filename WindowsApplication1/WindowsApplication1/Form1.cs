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
        public string DictionaryFileName = "dictionary.dat";
        public string convencionalDateNull = "1900-01-01";
        public System.Drawing.Color ExcelQueryHeaderColor = System.Drawing.Color.LightBlue;
        public enum category { Distribution, Transmission, Unknown };
        public enum data { Case, Bus, BusType, Bus_BusTypeRelationship, Line, LineSpacing, Conductor, LineType, Line_LineTypeRelationship, LossZone, Bus_LossZoneRelationship, Line_LossZoneRelationship, LoadModel, Load, DistributedLoad, ShuntElement, Transformer, Line_TransformerRelationship, Bus_BusControl, Line_BusControl, Generation, Regulator, TieLines, Interchange, Area, Unknown};
        public enum operation { Insert, Remove, Update, Query, Unknown };

        public void GenerateNewForm(string password, int categorySelectedValue, int dataSelectedValue, int operationSelectedValue)
        {
            category Category = SelectedValueToEnumeratedCategoryID(categorySelectedValue);
            data Data = SelectedValueToEnumeratedDataID(dataSelectedValue);
            operation Operation = SelectedValueToEnumeratedOperationID(operationSelectedValue);

            if (Category == category.Transmission && Data == data.Bus_BusTypeRelationship && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel8);
                label27.Text = GetLabel(DictionaryFileName,"Bus_BusType.CaseID");
                label28.Text = GetLabel(DictionaryFileName,"Bus_BusType.BusID");
                label29.Text = GetLabel(DictionaryFileName, "Bus_BusType.BusTypeID");
                button8.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button9.Text = GetLabel(DictionaryFileName, "ClearButton");
                SetTransmissionItemList(comboBox6);
            }

            if (Category == category.Transmission && Data == data.BusType && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel7);
                SetTransmissionItemList(comboBox5);
                label20.Text = GetLabel(DictionaryFileName, "BusType.ID");
                label22.Text = GetLabel(DictionaryFileName, "BusType.CaseID");
                label26.Text = GetLabel(DictionaryFileName, "BusType.Description");
                button6.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button7.Text = GetLabel(DictionaryFileName, "ClearButton");
            }

            if (Category == category.Transmission && Data == data.Bus && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel6);
                SetTransmissionItemList(comboBox3);
                label7.Text = GetLabel(DictionaryFileName, "Bus.Case");
                label8.Text = GetLabel(DictionaryFileName, "Bus.Number");
                label9.Text = GetLabel(DictionaryFileName, "Bus.SequencialNumber");
                label10.Text = GetLabel(DictionaryFileName, "Bus.Voltage");
                label11.Text = GetLabel(DictionaryFileName, "Bus.Phase");
                label12.Text = GetLabel(DictionaryFileName, "Bus.VoltageBase");
                label16.Text = GetLabel(DictionaryFileName, "Bus.DesiredVoltage");
                label19.Text = GetLabel(DictionaryFileName, "Bus.MaxReactivePowerOrVoltage");
                label21.Text = GetLabel(DictionaryFileName, "Bus.MinReactivePowerOrVoltage");
                label23.Text = GetLabel(DictionaryFileName, "Bus.Name");
                label24.Text = GetLabel(DictionaryFileName, "Bus.Area");
                button4.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button5.Text = GetLabel(DictionaryFileName, "ClearButton");
            }

            if (Category == category.Transmission && Data == data.Case && Operation == operation.Insert)
            {
                label15.Text = GetLabel(DictionaryFileName, "Case.Title");
                label13.Text = GetLabel(DictionaryFileName, "Case.Description");
                label14.Text = GetLabel(DictionaryFileName, "Case.PowerBase");
                label18.Text = GetLabel(DictionaryFileName, "Case.CaseDate");
                label17.Text = GetLabel(DictionaryFileName, "Case.PublicationDate");
                button2.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button3.Text = GetLabel(DictionaryFileName, "ClearButton");
                SetPanelLocationAndVisibility(panel5);
            }
        }

        public operation SelectedValueToEnumeratedOperationID(int operationID)
        {
            switch (operationID)
            {
                case 0:
                    return operation.Insert;
                case 1:
                    return operation.Remove;
                case 2:
                    return operation.Update;
                case 3:
                    return operation.Query;
            }
            return operation.Unknown;
        }

        public category SelectedValueToEnumeratedCategoryID(int categoryID)
        {
            switch (categoryID)
            {
                case 0:
                    return category.Distribution;
                case 1:
                    return category.Transmission;
            }
            return category.Unknown;
        }

        public data SelectedValueToEnumeratedDataID(int dataID)
        {
            dataID = dataID +1;
            switch (dataID)
            {
                case 1:
                    return data.Case; ;
                case 2:
                    return data.Bus;
                case 3:
                    return data.BusType;
                case 4:
                    return data.Bus_BusTypeRelationship;
                case 5:
                    return data.Line;
                case 6:
                    return data.LineSpacing;
                case 7:
                    return data.Conductor;
                case 8:
                    return data.LineType;
                case 9:
                    return data.Line_LineTypeRelationship;
                case 10:
                    return data.LossZone;
                case 11:
                    return data.Bus_LossZoneRelationship;
                case 12:
                    return data.Line_LossZoneRelationship;
                case 13:
                    return data.LoadModel;
                case 14:
                    return data.Load;
                case 15:
                    return data.DistributedLoad;
                case 16:
                    return data.ShuntElement;
                case 17:
                    return data.Transformer;
                case 18:
                    return data.Line_TransformerRelationship;
                case 19:
                    return data.Bus_BusControl;
                case 20:
                    return data.Line_BusControl;
                case 21:
                    return data.Generation;
                case 22:
                    return data.Regulator;
                case 23:
                    return data.TieLines;
                case 24:
                    return data.Interchange;
                case 25:
                    return data.Area;
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
            DatabaseAccess.Query databaseAccess = new DatabaseAccess.Query();
            List<string[]> matrix = new List<string[]>();
            matrix = databaseAccess.query(GetLabel("config.ini", "Host"), GetLabel("config.ini", "UserID"), GetLabel("config.ini", "DatabaseName"), maskedTextBox1.Text, "SELECT * FROM  `case`");
            matrix = CaseDistribuctionFiltering(matrix);
            return matrix;
        }

        private List<string[]> GetTransmissionMatrix()
        {
            DatabaseAccess.Query databaseAccess = new DatabaseAccess.Query();
            List<string[]> matrix = new List<string[]>();
            matrix = databaseAccess.query(GetLabel("config.ini", "Host"), GetLabel("config.ini", "UserID"), GetLabel("config.ini", "DatabaseName"), maskedTextBox1.Text,"SELECT * FROM  `case`");
            matrix = CaseTransmissionFiltering(matrix);
            return matrix;
        }

        private void SetAreaList(ComboBox comboBox, int caseID)
        {
            comboBox.Text = string.Empty;
            comboBox.Items.Clear();
            DatabaseAccess.Query databaseAccess = new DatabaseAccess.Query();
            List<string[]> matrix = new List<string[]>();
            matrix = databaseAccess.query(GetLabel("config.ini", "Host"), GetLabel("config.ini", "UserID"), GetLabel("config.ini", "DatabaseName"), maskedTextBox1.Text, "SELECT * FROM  `area` WHERE `caseID` = " + caseID.ToString() + "");
            for (int i = 0; i < matrix.Count; i++)
            {
                if (caseID.ToString().Equals(matrix[i][1]))
                { 
                    comboBox.Items.Add(matrix[i][0] + " - " + matrix[i][2]);
                }
            }
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
                if (matrix[i][matrix[0].Length - 1].Equals("DISTRIBUTION"))
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
                if (matrix[i][matrix[0].Length - 1].Equals("TRANSMISSION"))
                {
                    filteredMatrix.Add(matrix[i]);
                }
            }
            return filteredMatrix;
        }

        public string GetSystemLabel(string DictionaryFileName, string ErrorMsg, string Key)
        {
            string filePath = Application.StartupPath;
            List<string> Buffer = new List<string>();
            System.IO.StreamReader file = new System.IO.StreamReader(@"" + filePath + "//" + DictionaryFileName, System.Text.Encoding.UTF7);
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

        public string GetLabel(string DictionaryFileName, string Key)
        {
            string ErrorMsg = "ERROR.998";
            string label = GetSystemLabel(DictionaryFileName, ErrorMsg, Key);
            return label;
        }

        public string GetErrorLabel(string DictionaryFileName, int ErrorCode)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "Error." + ErrorCode;
            string label = GetSystemLabel(DictionaryFileName, ErrorMsg, Key);
            return label;
        }

        public void ShowError(int ErrorCode, string ExtraInfo)
        {
            string Error = GetErrorLabel(DictionaryFileName, ErrorCode);
            label5.Text = Error + GetVerifyInfo(DictionaryFileName) + ": " + ExtraInfo;
        }

        public string GetVerifyInfo(string DictionaryFileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "VerifyInfo";
            string label = GetSystemLabel(DictionaryFileName, ErrorMsg, Key);
            return label;
        }
        
        public string GetSystemTitleLabel()
        {
            string ErrorMsg = "Error: No system title found. Check dictonary.dat file.";
            string Key = "SystemName";
            string label = GetSystemLabel(DictionaryFileName, ErrorMsg, Key);
            return label;
        }

        public Form1()
        {
            InitializeComponent();
            this.Text = GetSystemTitleLabel();
            this.Size = new Size(852, 543);
            panel1.Location = new Point(12, 14);
            panel1.Size = new Size(254, 408);
            panel2.Location = new Point(12, 427);
            panel2.Size = new Size(810, 25);
            panel3.Location = new Point(12, 449);
            panel3.Size = new Size(810, 39);
        }


        private void SetPanelLocationAndVisibility(Panel panel)
        {
            int PanelLocationX = 272;
            int PanelLocationY = 12;
            int PanelLocationH = 550;
            int PanelLocationW = 410;
            panel.Location = new Point(PanelLocationX, PanelLocationY);
            panel.Size = new Size(PanelLocationH, PanelLocationW);
            panel.Visible = true;
            panel4.Visible = false;
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            label5.Text = GetLabel(DictionaryFileName, "Error.RunMsg");
            label1.Text = GetLabel(DictionaryFileName, "CategoryLabel") + ":";
            label2.Text = GetLabel(DictionaryFileName, "OperationLabel") + ":";
            label3.Text = GetLabel(DictionaryFileName, "DataLabel") + ":";
            radioButton1.Text = GetLabel(DictionaryFileName, "OperationInformation.Insert");
            radioButton2.Text = GetLabel(DictionaryFileName, "OperationInformation.Remove");
            radioButton3.Text = GetLabel(DictionaryFileName, "OperationInformation.Update");
            radioButton4.Text = GetLabel(DictionaryFileName, "OperationInformation.Query");

            button1.Text = GetLabel(DictionaryFileName, "RunButton");

            label4.Text = GetLabel(DictionaryFileName, "Connection.DatabasePassword") + ":";

            label6.Text = GetLabel(DictionaryFileName, "UserMsgLabel") + ":";
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Case"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Bus"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.BusType"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Bus_BusTypeRelationship"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Line"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.LineSpacing"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Conductor"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.LineType"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Line_LineTypeRelationship"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.LossZone"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Bus_LossZoneRelationship"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Line_LossZoneRelationship"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.LoadModel"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Load"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.DistributedLoad"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.ShuntElement"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Transformer"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Line_TransformerRelationship"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Bus_BusControl"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Line_BusControl"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Generation"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Regulator"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.TieLines"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Interchange"));
            comboBox2.Items.Add(GetLabel(DictionaryFileName, "DataInformation.Area"));
            comboBox1.Items.Add(GetLabel(DictionaryFileName, "CategoryInformation.Distribution"));
            comboBox1.Items.Add(GetLabel(DictionaryFileName, "CategoryInformation.Transmission"));
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

        public string GetRunLabel(string DictionaryFileName)
        {
            string ErrorMsg = "ERROR.998";
            string Key = "RunButton";
            string label = GetSystemLabel(DictionaryFileName, ErrorMsg, Key);
            return label;
        }

        private void Close_panel()
        {
            panel5.Visible = false;
            panel6.Visible = false;
            panel7.Visible = false;
            panel8.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Visible = false;
            if(button1.Text.Equals(GetLabel(DictionaryFileName,"StopButton")))
            {
                // Clean panel
                Close_panel();

                // Back
                comboBox1.Enabled = true;
                comboBox2.Enabled = true;
                maskedTextBox1.Enabled = true;
                radioButton1.Enabled = true;
                radioButton2.Enabled = true;
                panel4.Visible = true;
                radioButton3.Enabled = true;
                radioButton4.Enabled = true;
                button1.Text = GetLabel(DictionaryFileName, "RunButton");
            }
            else if (button1.Text.Equals(GetLabel(DictionaryFileName,"RunButton")))
            {
                List<bool> validation = new List<bool>();
                validation.Add(ValidateAsSelectedfromCombobox(comboBox2, label3));
               

                if (radioButton1.Checked == false && radioButton2.Checked == false && radioButton3.Checked == false && radioButton4.Checked == false)
                {
                    validation.Add(false);
                    ShowError(989, label2.Text.Split(':')[0]);
                }
                else
                {
                    validation.Add(true);
                }

                validation.Add(ValidateAsSelectedfromCombobox(comboBox1, label1));

                if (CompleteValidation(validation) == true)
                {
                    bool noerror = false;
                    // Testar conexão
                    DatabaseAccess.Query databaseAccess = new DatabaseAccess.Query();
                    List<string[]> matrix = new List<string[]>();
                    bool errorShown = false;
                    try
                    { 
                       matrix = databaseAccess.query(GetLabel("config.ini", "Host"), GetLabel("config.ini", "UserID"), GetLabel("config.ini", "DatabaseName"), maskedTextBox1.Text, "SELECT * FROM `case`");
                       if (matrix[0][0].ToLower().Contains("*error*") && matrix[0][0].ToLower().Contains("access denied for user"))
                       {
                           ShowError(997, label4.Text.Split(':')[0]);
                           errorShown = true;
                       }
                       else if (matrix[0][0].ToLower().Contains("*error*"))
                       {
                           ShowError(984, GetLabel(DictionaryFileName,"InternetConnection"));
                           errorShown = true;
                       }
                       else
                       {
                           noerror = true;
                       }
                    }
                    catch
                    {
                        if (matrix.Count == 0)
                        {
                            // Means that read ok but there  isn't anything on the database
                            noerror = true;
                        }
                    }
                    if (noerror == false)
                    {
                        if ( errorShown == false)
                        { 
                            ShowError(997, label4.Text.Split(':')[0]);
                        }
                    }
                    else
                    {
                        comboBox1.Enabled = false;
                        comboBox2.Enabled = false;
                        label5.Text = GetLabel(DictionaryFileName, "NoError.OK");
                        maskedTextBox1.Enabled = false;
                        radioButton1.Enabled = false;
                        int operationSelectedValue = -1;
                        if (radioButton1.Checked == true)
                        {
                            operationSelectedValue = 0;
                        }
                        radioButton2.Enabled = false;
                        if (radioButton2.Checked == true)
                        {
                            operationSelectedValue = 1;
                        }
                        radioButton3.Enabled = false;
                        if (radioButton3.Checked == true)
                        {
                            operationSelectedValue = 2;
                        }
                        radioButton4.Enabled = false;
                        if (radioButton4.Checked == true)
                        {
                            operationSelectedValue = 3;
                        }

                        GenerateNewForm(maskedTextBox1.Text, comboBox1.SelectedIndex, comboBox2.SelectedIndex, operationSelectedValue);
                        button1.Text = GetLabel(DictionaryFileName, "StopButton");
                    }
                }
                else
                {

                }
            }
            else
            {
                ShowError(992, "dictionary.dat");
            }
            button1.Visible = true;
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

        }

        private void panel7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel7_Paint_1(object sender, PaintEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
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

        
        private void textBox13_TextChanged(object sender, EventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {
            
        }

        private void panel11_Paint(object sender, PaintEventArgs e)
        {

        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            
        }

        private string[] RemoveIndice(string[] IndicesArray, int RemoveAt)
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
                System.Drawing.Color HeaderColor = ExcelQueryHeaderColor;
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

        private bool ValidateAsNotNullRichText(RichTextBox texbox, Label label)
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
                ShowError(989, label.Text);
                return false;
            }
            else
            {
                return true;
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

        private void panel4_Paint(object sender, PaintEventArgs e)
        {
            // INSERT INTO `sql583577`.`case` (`id`, `title`, `description`, `powerBase`, `caseDate`, `publicationDate`) VALUES ('1', 'teste', 'teste', '1.1', '2015-07-08', '2015-07-08');
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void panel5_Paint_1(object sender, PaintEventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            textBox5.Text = string.Empty;
            textBox6.Text = string.Empty;
            richTextBox2.Text = string.Empty;
            dateTimePicker1.Value = new DateTime(Convert.ToInt32(convencionalDateNull.Split('-')[0]), Convert.ToInt32(convencionalDateNull.Split('-')[1]), Convert.ToInt32(convencionalDateNull.Split('-')[2]));
            dateTimePicker2.Value = new DateTime(Convert.ToInt32(convencionalDateNull.Split('-')[0]), Convert.ToInt32(convencionalDateNull.Split('-')[1]), Convert.ToInt32(convencionalDateNull.Split('-')[2]));
        }

        private void showMsg(string Msg)
        {
            label5.Text = Msg;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();

            string publicationDate = string.Empty;
            publicationDate = dateTimePicker2.Value.Year.ToString() + "-" + dateTimePicker2.Value.Month.ToString() + "-" + dateTimePicker2.Value.Day.ToString();

            string caseDate = string.Empty;
            caseDate = dateTimePicker1.Value.Year.ToString() + "-" + dateTimePicker1.Value.Month.ToString() + "-" + dateTimePicker1.Value.Day.ToString();

            double powerBase = 0;
            validationList.Add(ValidateAsDouble(textBox6, label14, out powerBase));

            validationList.Add(ValidateAsNotNullRichText(richTextBox2, label13));
            validationList.Add(ValidateAsNotNullText(textBox5, label15));
            
            if(CompleteValidation(validationList)==true)
            {
                DatabaseAccess.Insert databaseAccess = new DatabaseAccess.Insert();
                string returnedMsg = databaseAccess.insert(GetLabel("config.ini", "Host"), GetLabel("config.ini", "UserID"), GetLabel("config.ini", "DatabaseName"), maskedTextBox1.Text, "INSERT INTO `sql583577`.`case` (`title`, `description`, `powerBase`, `caseDate`, `publicationDate`, `systemType`) VALUES ('" + textBox5.Text + "', '" + richTextBox2.Text + "', '" + textBox6.Text.Replace(',', '.') + "', '" + caseDate + "', '" + publicationDate + "', 'TRANSMISSION');");

                //ValidateInsert(returnedMsg);
                if (returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName, "InsertSuccess"));
                    textBox5.Text = string.Empty;
                    richTextBox2.Text = string.Empty;
                    textBox6.Text = string.Empty;
                }
                else
                {
                    ShowError(992, returnedMsg);
                }
            }
        }

        private void comboBox4_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private int GetAreaID(int selectedIndex)
        {
            if (selectedIndex >= 0)
            {
                List<string[]> matrix = new List<string[]>();
                DatabaseAccess.Query databaseAccess = new DatabaseAccess.Query();
                matrix = databaseAccess.query(GetLabel("config.ini", "Host"), GetLabel("config.ini", "UserID"), GetLabel("config.ini", "DatabaseName"), maskedTextBox1.Text, "SELECT * FROM  `area`");
                try
                {
                    return Convert.ToInt32(matrix[selectedIndex][0]);
                }
                catch
                {
                    return -1;
                }
                
            }
            else
            {
                return -1;
            }
        }

        private int GetTransmissionCaseID(int selectedIndex)
        {
            if (selectedIndex >= 0)
            {
                List<string[]> matrix = new List<string[]>();
                matrix = GetTransmissionMatrix();
                return Convert.ToInt32(matrix[selectedIndex][0]);
            }
            else
            {
                return -1;
            }
        }

        private int GetDistributionCaseID(int selectedIndex)
        {
            if (selectedIndex >= 0)
            {
                List<string[]> matrix = new List<string[]>();
                matrix = GetDistributionMatrix();
                return Convert.ToInt32(matrix[selectedIndex][0]);
            }
            else
            {
                return -1;
            }
        }

        private void comboBox3_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void panel6_Paint_1(object sender, PaintEventArgs e)
        {

        }

        private List<string[]> Query(string msgQuery)
        {
            List<string[]> matrix = new List<string[]>();
            DatabaseAccess.Query databaseAccess = new DatabaseAccess.Query();
            matrix = databaseAccess.query(GetLabel("config.ini", "Host"), GetLabel("config.ini", "UserID"), GetLabel("config.ini", "DatabaseName"), maskedTextBox1.Text, msgQuery);
            return matrix;
        }

        private void comboBox3_TextUpdate(object sender, EventArgs e)
        {
            
        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            int selectedIndex = comboBox3.SelectedIndex;
            int ID = GetTransmissionCaseID(selectedIndex);
            if (ID >= 0)
            {
                SetAreaList(comboBox4, ID);
            }
            else
            {
                comboBox4.Text = string.Empty;
                comboBox4.Items.Clear();
            }
        }

        private void label25_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            comboBox3.Text = string.Empty;
            textBox1.Text = string.Empty;
            textBox8.Text = string.Empty;
            textBox2.Text = string.Empty;
            textBox12.Text = string.Empty;
            textBox3.Text = string.Empty;
            textBox11.Text = string.Empty;
            textBox4.Text = string.Empty;
            textBox7.Text = string.Empty;
            textBox13.Text = string.Empty;
            comboBox4.Text = string.Empty;
        }

        private void ShowInsertError(string returnedMsg)
        {
            if (returnedMsg.Contains("Duplicate entry"))
            {
                ShowError(985, GetLabel(DictionaryFileName, "Database"));
            }
            else
            {
                ShowError(992, returnedMsg);
            }
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox4, label24));
            validationList.Add(ValidateAsNotNullText(textBox13, label23));
            double minLimit = 0;
            validationList.Add(ValidateAsDouble(textBox7, label21, out minLimit));
            double maxLimit = 0;
            validationList.Add(ValidateAsDouble(textBox4, label19, out maxLimit));
            double desiredVoltage = 0;
            validationList.Add(ValidateAsDouble(textBox11, label16, out desiredVoltage));
            double voltageBase = 0;
            validationList.Add(ValidateAsDouble(textBox3, label12, out voltageBase));
            double phase = 0;
            validationList.Add(ValidateAsDouble(textBox12, label11, out phase));
            double voltage = 0;
            validationList.Add(ValidateAsDouble(textBox2, label10, out voltage));
            int sequencialNumber = 0;
            validationList.Add(ValidateAsInt(textBox8,label9,out sequencialNumber));
            int busNumber = 0;
            validationList.Add(ValidateAsInt(textBox1,label9,out busNumber));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox3, label7));

            if(CompleteValidation(validationList)==true)
            {
                int caseID = GetTransmissionCaseID(comboBox3.SelectedIndex);
                int areaID = GetAreaID(comboBox4.SelectedIndex);
                DatabaseAccess.Insert databaseAccess = new DatabaseAccess.Insert();
                string returnedMsg = databaseAccess.insert(GetLabel("config.ini", "Host"), GetLabel("config.ini", "UserID"), GetLabel("config.ini", "DatabaseName"), maskedTextBox1.Text, "INSERT INTO `sql583577`.`bus` (`busNumber`, `caseID`, `sequencialNumber`, `busName`, `Voltage`, `phase`, `voltageBase`, `desiredVoltage`, `maxReactivePower`, `minReactivePower`, `areaID`) VALUES ('" + busNumber.ToString() + "', '" + caseID.ToString() + "', '" + sequencialNumber.ToString() + "', '" + textBox13.Text + "', '" + voltage.ToString().Replace(',', '.') + "', '" + phase.ToString().Replace(',', '.') + "', '" + voltageBase.ToString().Replace(',', '.') + "', '" + desiredVoltage.ToString().Replace(',', '.') + "', '" + maxLimit.ToString().Replace(',', '.') + "', '" + minLimit.ToString().Replace(',', '.') + "', '" + areaID.ToString() + "');");
                if (returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName, "InsertSuccess"));
                    comboBox3.Text = string.Empty;
                    textBox1.Text = string.Empty;
                    textBox8.Text = string.Empty;
                    textBox2.Text = string.Empty;
                    textBox12.Text = string.Empty;
                    textBox3.Text = string.Empty;
                    textBox11.Text = string.Empty;
                    textBox4.Text = string.Empty;
                    textBox7.Text = string.Empty;
                    textBox13.Text = string.Empty;
                    comboBox4.Text = string.Empty;
                }
                else
                {
                    ShowInsertError(returnedMsg);
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            textBox9.Text = string.Empty;
            comboBox5.Text = string.Empty;
            richTextBox1.Text = string.Empty;
        }

        private string Insert(string insertMsg)
        {
            DatabaseAccess.Insert databaseAccess = new DatabaseAccess.Insert();
            return databaseAccess.insert(GetLabel("config.ini", "Host"), GetLabel("config.ini", "UserID"), GetLabel("config.ini", "DatabaseName"), maskedTextBox1.Text, insertMsg);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            validationList.Add(ValidateAsNotNullRichText(richTextBox1, label22));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox5, label21));
            int busTypeID = 0;
            validationList.Add(ValidateAsInt(textBox9, label20, out busTypeID));

            if(CompleteValidation(validationList))
            {
                int caseID = GetTransmissionCaseID(comboBox5.SelectedIndex);
                string returnedMsg = Insert("INSERT INTO `sql583577`.`bustype` (`busTypeID`, `description`, `caseID`) VALUES ('" + busTypeID.ToString() + "', '" + richTextBox1.Text + "', '" + caseID.ToString() + "');");
                if (returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName,"InsertSuccess"));
                    textBox9.Text = string.Empty;
                    comboBox5.Text = string.Empty;
                    richTextBox1.Text = string.Empty;
                }
                else
                {
                    ShowInsertError(returnedMsg);
                }
            }

        }

        private void button9_Click(object sender, EventArgs e)
        {
            comboBox6.Text = string.Empty;
            comboBox7.Text = string.Empty;
            comboBox8.Text = string.Empty;
        }

        private bool checkQueryError(List<string[]> matrix)
        {
            try
            {
                if (matrix[0][0].ToLower().Equals("*error*"))
                {
                    ShowError(984, GetLabel(DictionaryFileName, "InternetConnection"));
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch
            {
                return true;
            }
        }

        private void SetBusList(ComboBox combobox, int caseID)
        {
            combobox.Items.Clear();
            List<string[]> matrix = new List<string[]>();
            matrix = Query("SELECT * FROM  `bus` WHERE `caseID` = " + caseID.ToString() + "");
            if (checkQueryError(matrix))
            {
                for (int i = 0; i < matrix.Count; i++)
                {
                    combobox.Items.Add("BusNumber: " + matrix[i][0] + " / " + "busName: " + matrix[i][3]);
                }
            }
            else
            {
                combobox.Text = string.Empty;
            }
        }

        private int GetBusID(int selectedIndex, int caseID)
        {
            List<string[]> matrix = new List<string[]>();
            matrix = Query("SELECT * FROM  `bus` WHERE `caseID` = " + caseID.ToString() + "");
            if (checkQueryError(matrix))
            {
                return Convert.ToInt32(matrix[selectedIndex][0]);
            }
            else
            {
                return -1;
            }
        }
        
        private void SetBusTypeList(ComboBox combobox, int caseID)
        {
            combobox.Items.Clear();
            List<string[]> matrix = new List<string[]>();
            matrix = Query("SELECT * FROM  `bustype` WHERE `caseID` = " + caseID.ToString() + "");
            if (checkQueryError(matrix))
            {
                for (int i = 0; i < matrix.Count; i++)
                {
                    combobox.Items.Add("BusTypeID: " + matrix[i][0] + " / " + "Description: " + matrix[i][1]);
                    
                }
            }
            else
            {
                combobox.Text = string.Empty;
            }
        }

        private int GetBusTypeID(int selectedIndex, int caseID)
        {
            List<string[]> matrix = new List<string[]>();
            matrix = Query("SELECT * FROM  `bustype` WHERE `caseID` = " + caseID.ToString() + "");
            if (checkQueryError(matrix))
            {
                return Convert.ToInt32(matrix[selectedIndex][0]);
            }
            else
            {
                return -1;
            }
        }

        private void comboBox6_TextChanged(object sender, EventArgs e)
        {
            if(comboBox6.SelectedIndex>-1)
            {
                int caseID = GetTransmissionCaseID(comboBox6.SelectedIndex);
                SetBusList(comboBox7, caseID);
                SetBusTypeList(comboBox8, caseID);
            }
            else
            {
                comboBox7.Text = string.Empty;
                comboBox8.Text = string.Empty;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox8, label29));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox7, label28));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox6, label27));

            if(CompleteValidation(validationList))
            {
                int caseID = GetTransmissionCaseID(comboBox6.SelectedIndex);
                int BusID = GetBusID(comboBox7.SelectedIndex, caseID);
                int BusTypeID = GetBusID(comboBox8.SelectedIndex, caseID);
                string returnedMsg = Insert("INSERT INTO `sql583577`.`bus_bustype` (`idCase`, `busNumber`, `busType`) VALUES ('" + caseID + "', '" + BusID.ToString() + "', '" + BusTypeID.ToString() + "');");
                if(returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName, "InsertSuccess"));
                    comboBox8.Text = string.Empty;
                    comboBox6.Text = string.Empty;
                    comboBox7.Text = string.Empty;
                }
                else
                {
                    ShowInsertError(returnedMsg);
                }
            }
        }
    }
}