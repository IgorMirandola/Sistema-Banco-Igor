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

            if (Category == category.Transmission && Data == data.Generation && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel25);
                SetTransmissionItemList(comboBox38);
                label94.Text = GetLabel(DictionaryFileName, "Generation.CaseID");
                label95.Text = GetLabel(DictionaryFileName, "Generation.BusID");
                label96.Text = GetLabel(DictionaryFileName, "Generation.ActivePower");
                label97.Text = GetLabel(DictionaryFileName, "Generation.ReactivePower");
                button34.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button35.Text = GetLabel(DictionaryFileName, "ClearButton");
            }

            if (Category == category.Transmission && Data == data.Line_BusControl && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel24);
                SetTransmissionItemList(comboBox35);
                label89.Text = GetLabel(DictionaryFileName,"Line_BusControl.CaseID");
                label90.Text = GetLabel(DictionaryFileName, "Line_BusControl.ControlledBusID");
                label91.Text = GetLabel(DictionaryFileName, "Line_BusControl.LineID");
                label92.Text = GetLabel(DictionaryFileName, "Line_BusControl.Side");
                button32.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button33.Text = GetLabel(DictionaryFileName, "ClearButton");
            }

            if (Category == category.Transmission && Data == data.Bus_BusControl && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel23);
                SetTransmissionItemList(comboBox32);
                label86.Text = GetLabel(DictionaryFileName, "Bus_BusControl.CaseID");
                label87.Text = GetLabel(DictionaryFileName, "Bus_BusControl.ControlledBus");
                label88.Text = GetLabel(DictionaryFileName, "Bus_BusControl.ControllerBus");
                button30.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button31.Text = GetLabel(DictionaryFileName, "ClearButton");
            }

            if (Category == category.Transmission && Data == data.Line_TransformerRelationship && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel22);
                SetTransmissionItemList(comboBox29);
                label83.Text = GetLabel(DictionaryFileName, "Line_TransformerRelationship.CaseID");
                label84.Text = GetLabel(DictionaryFileName, "Line_TransformerRelationship.LineID");
                label85.Text = GetLabel(DictionaryFileName, "Line_TransformerRelationship.TransformerID");
                button28.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button29.Text = GetLabel(DictionaryFileName, "ClearButton");

            }

            if (Category == category.Transmission && Data == data.Transformer && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel21);
                SetTransmissionItemList(comboBox28);
                label75.Text = GetLabel(DictionaryFileName, "Transformer.Case");
                label76.Text = GetLabel(DictionaryFileName, "Transformer.FinalVoltageRatio");
                label77.Text = GetLabel(DictionaryFileName, "Transformer.FinalPhaseAngle");
                label78.Text = GetLabel(DictionaryFileName, "Transformer.FinalVoltageRatioOrFinalPhaseAngleMinLimit");
                label79.Text = GetLabel(DictionaryFileName, "Transformer.FinalVoltageRatioOrFinalPhaseAngleMaxLimit");
                label80.Text = GetLabel(DictionaryFileName, "Transformer.stepSize");
                label81.Text = GetLabel(DictionaryFileName, "Transformer.voltageOrPowerMinLimit");
                label82.Text = GetLabel(DictionaryFileName, "Transformer.voltageOrPowerMaxLimit");
                button26.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button27.Text = GetLabel(DictionaryFileName, "ClearButton");
            }

            if (Category == category.Transmission && Data == data.ShuntElement && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel20);
                SetTransmissionItemList(comboBox26);
                label70.Text = GetLabel(DictionaryFileName, "ShuntElement.CaseID");
                label71.Text = GetLabel(DictionaryFileName, "ShuntElement.Bus");
                label72.Text = GetLabel(DictionaryFileName, "ShuntElement.Conductance");
                label73.Text = GetLabel(DictionaryFileName, "ShuntElement.Susceptance");
                label74.Text = GetLabel(DictionaryFileName, "ShuntElement.Description");
                button24.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button25.Text = GetLabel(DictionaryFileName, "ClearButton");
            }

            if (Category == category.Transmission && Data == data.DistributedLoad && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel19);
                label68.Text = GetLabel(DictionaryFileName, "FormNotUsed");
            }

            if (Category == category.Transmission && Data == data.Load && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel18);
                SetTransmissionItemList(comboBox24);
                label64.Text = GetLabel(DictionaryFileName, "Load.CaseID");
                label65.Text = GetLabel(DictionaryFileName, "Load.BusID");
                label66.Text = GetLabel(DictionaryFileName, "Load.ActivePower");
                label67.Text = GetLabel(DictionaryFileName, "Load.ReactivePower");
                button22.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button23.Text = GetLabel(DictionaryFileName, "ClearButton");
            }

            if (Category == category.Transmission && Data == data.LoadModel && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel17);
                label63.Text = GetLabel(DictionaryFileName, "FormNotUsed");
            }

            if (Category == category.Transmission && Data == data.Line_LossZoneRelationship && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel16);
                SetTransmissionItemList(comboBox23);
                label62.Text = GetLabel(DictionaryFileName, "Line_LossZoneRelationship.CaseID");
                label61.Text = GetLabel(DictionaryFileName, "Line_LossZoneRelationship.LineID");
                label60.Text = GetLabel(DictionaryFileName, "Line_LossZoneRelationship.LossZoneID");
                button21.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button20.Text = GetLabel(DictionaryFileName, "ClearButton");
            }

            if (Category == category.Transmission && Data == data.Bus_LossZoneRelationship && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel15);
                SetTransmissionItemList(comboBox18);
                label57.Text = GetLabel(DictionaryFileName, "Bus_LossZoneRelationship.CaseID");
                label57.Text = GetLabel(DictionaryFileName, "Bus_LossZoneRelationship.BusID");
                label57.Text = GetLabel(DictionaryFileName, "Bus_LossZoneRelationship.LossZoneID");
                button18.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button19.Text = GetLabel(DictionaryFileName, "ClearButton");
            }

            if (Category == category.Transmission && Data == data.LossZone && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel14);
                SetTransmissionItemList(comboBox17);
                label52.Text = GetLabel(DictionaryFileName, "LossZone.CaseID");
                label53.Text = GetLabel(DictionaryFileName, "LossZone.LossZone");
                label54.Text = GetLabel(DictionaryFileName, "LossZone.SequencialNumber");
                label55.Text = GetLabel(DictionaryFileName, "LossZone.Description");
                button16.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button17.Text = GetLabel(DictionaryFileName, "ClearButton");
            }

            if (Category == category.Transmission && Data == data.Line_LineTypeRelationship && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel13);
                SetTransmissionItemList(comboBox14);
                label49.Text = GetLabel(DictionaryFileName, "LineLineType.CaseID");
                label50.Text = GetLabel(DictionaryFileName, "LineLineType.LineID");
                label51.Text = GetLabel(DictionaryFileName, "LineLineType.LineTypeID");
                button14.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button15.Text = GetLabel(DictionaryFileName, "ClearButton");
            }

            if (Category == category.Transmission && Data == data.LineType && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel12);
                SetTransmissionItemList(comboBox13);
                label46.Text = GetLabel(DictionaryFileName,"LineType.caseID");
                label47.Text = GetLabel(DictionaryFileName,"LineType.ID");
                label48.Text = GetLabel(DictionaryFileName,"LineType.Description");
                button12.Text = GetLabel(DictionaryFileName,"SubmitButton");
                button13.Text = GetLabel(DictionaryFileName, "ClearButton");
            }

            if (Category == category.Transmission && Data == data.Conductor && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel11);
                label44.Text = GetLabel(DictionaryFileName, "FormNotUsed");
            }

            if (Category == category.Transmission && Data == data.LineSpacing && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel10);
                label43.Text = GetLabel(DictionaryFileName, "FormNotUsed");
            }

            if (Category == category.Transmission && Data == data.Line && Operation == operation.Insert)
            {
                SetPanelLocationAndVisibility(panel9);
                SetTransmissionItemList(comboBox9);
                label30.Text = GetLabel(DictionaryFileName, "Line.Case");
                label31.Text = GetLabel(DictionaryFileName, "Line.InicialBus");
                label32.Text = GetLabel(DictionaryFileName, "Line.FinalBus");
                label33.Text = GetLabel(DictionaryFileName, "Line.SequencialNumber");
                label34.Text = GetLabel(DictionaryFileName, "Line.Resistence");
                label39.Text = GetLabel(DictionaryFileName, "Line.Reactance");
                label38.Text = GetLabel(DictionaryFileName, "Line.Susceptance");
                label37.Text = GetLabel(DictionaryFileName, "Line.PowerRating1");
                label36.Text = GetLabel(DictionaryFileName, "Line.PowerRating2");
                label35.Text = GetLabel(DictionaryFileName, "Line.PowerRating3");
                label40.Text = GetLabel(DictionaryFileName, "Line.Description");
                label42.Text = GetLabel(DictionaryFileName, "Line.CircuitNumber");
                label41.Text = GetLabel(DictionaryFileName, "Line.Area");
                button10.Text = GetLabel(DictionaryFileName, "SubmitButton");
                button11.Text = GetLabel(DictionaryFileName, "ClearButton");
            }

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

            // Define the border style of the form to a dialog box.
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            // Set the MaximizeBox to false to remove the maximize box.
            this.MaximizeBox = false;

            // Set the MinimizeBox to false to remove the minimize box.
            //this.MinimizeBox = false;

            // Set the start position of the form to the center of the screen.
            this.StartPosition = FormStartPosition.CenterScreen;

            // Display the form as a modal dialog box.
            //this.ShowDialog();

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
            panel23.Visible = false;
            panel24.Visible = false;
            panel25.Visible = false;
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
                ClearComboBox(comboBox4);
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

        private void SetAreaList(ComboBox combobox, int caseID)
        {
            combobox.Items.Clear();
            List<string[]> matrix = new List<string[]>();
            matrix = Query("SELECT * FROM  `area` WHERE `caseID` = " + caseID.ToString() + "");
            if (checkQueryError(matrix))
            {
                for (int i = 0; i < matrix.Count; i++)
                {
                    combobox.Items.Add(matrix[i][0] + " - " + matrix[i][2]);
                }
            }
            else
            {
                combobox.Text = string.Empty;
            }
        }

        private int GetAreaID(int selectedIndex, int caseID)
        {
            List<string[]> matrix = new List<string[]>();
            matrix = Query("SELECT * FROM  `area` WHERE `caseID` = " + caseID.ToString() + "");
            if (checkQueryError(matrix))
            {
                return Convert.ToInt32(matrix[selectedIndex][0]);
            }
            else
            {
                return -1;
            }
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

        private void SetLossZoneList(ComboBox combobox, int caseID)
        {
            combobox.Items.Clear();
            List<string[]> matrix = new List<string[]>();
            matrix = Query("SELECT * FROM  `lossZone` WHERE `caseID` = " + caseID.ToString() + "");
            if (checkQueryError(matrix))
            {
                for (int i = 0; i < matrix.Count; i++)
                {
                    combobox.Items.Add(matrix[i][0] + " / " + "description: " + matrix[i][2] + " / " + "sequencial Number: " + matrix[i][3]);
                }
            }
            else
            {
                combobox.Text = string.Empty;
            }
        }

        private int GetLossZoneID(int selectedIndex, int caseID)
        {
            List<string[]> matrix = new List<string[]>();
            matrix = Query("SELECT * FROM  `lossZone` WHERE `caseID` = " + caseID.ToString() + "");
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

        private void SetLineList(ComboBox combobox, int caseID)
        {
            combobox.Items.Clear();
            List<string[]> matrix = new List<string[]>();
            matrix = Query("SELECT * FROM  `line` WHERE `caseID` = " + caseID.ToString() + "");
            if (checkQueryError(matrix))
            {
                for (int i = 0; i < matrix.Count; i++)
                {
                    combobox.Items.Add("Inicial: " + matrix[i][2] + " / " + "Final: " + matrix[i][3] + " / " + "Number: " + matrix[i][3]);
                }
            }
            else
            {
                combobox.Text = string.Empty;
            }
        }


        private int GetLineID(int selectedIndex, int caseID)
        {
            List<string[]> matrix = new List<string[]>();
            matrix = Query("SELECT * FROM  `line` WHERE `caseID` = " + caseID.ToString() + "");
            if (checkQueryError(matrix))
            {
                return Convert.ToInt32(matrix[selectedIndex][0]);
            }
            else
            {
                return -1;
            }
        }

        private void SetTransformerList(ComboBox combobox, int caseID)
        {
            combobox.Items.Clear();
            List<string[]> matrix = new List<string[]>();
            matrix = Query("SELECT * FROM  `transformer` WHERE `caseID` = " + caseID.ToString() + "");
            if (checkQueryError(matrix))
            {
                for (int i = 0; i < matrix.Count; i++)
                {
                    combobox.Items.Add("finalVoltage: " + matrix[i][2] + " / " + "FinalAngle: " + matrix[i][3] + " / " + "stepSize: " + matrix[i][8]);
                }
            }
            else
            {
                combobox.Text = string.Empty;
            }
        }


        private int GetTransformerID(int selectedIndex, int caseID)
        {
            List<string[]> matrix = new List<string[]>();
            matrix = Query("SELECT * FROM  `transformer` WHERE `caseID` = " + caseID.ToString() + "");
            if (checkQueryError(matrix))
            {
                return Convert.ToInt32(matrix[selectedIndex][0]);
            }
            else
            {
                return -1;
            }
        }

        private void SetLineTypeList(ComboBox combobox, int caseID)
        {
            combobox.Items.Clear();
            List<string[]> matrix = new List<string[]>();
            matrix = Query("SELECT * FROM  `linetype` WHERE `caseID` = " + caseID.ToString() + "");
            if (checkQueryError(matrix))
            {
                for (int i = 0; i < matrix.Count; i++)
                {
                    combobox.Items.Add("ID: " + matrix[i][0] + " / " + "Description: " + matrix[i][7]);
                }
            }
            else
            {
                combobox.Text = string.Empty;
            }
        }

        private int GetLineTypeID(int selectedIndex, int caseID)
        {
            List<string[]> matrix = new List<string[]>();
            matrix = Query("SELECT * FROM  `linetype` WHERE `caseID` = " + caseID.ToString() + "");
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
                ClearComboBox(comboBox7);
                ClearComboBox(comboBox8);
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
                int BusTypeID = GetBusTypeID(comboBox8.SelectedIndex, caseID);
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

        private void label34_Click_1(object sender, EventArgs e)
        {

        }

        private void label33_Click_1(object sender, EventArgs e)
        {

        }

        private void comboBox9_TextChanged(object sender, EventArgs e)
        {
            if(comboBox9.SelectedIndex > -1)
            {
                int caseID = GetTransmissionCaseID(comboBox9.SelectedIndex);
                SetAreaList(comboBox10, caseID);
                SetBusList(comboBox11, caseID);
                SetBusList(comboBox12, caseID);
            }
            else
            {
                ClearComboBox(comboBox10);
                ClearComboBox(comboBox11);
                ClearComboBox(comboBox12);
            }
        }

        private void button11_Click_1(object sender, EventArgs e)
        {
            comboBox9.Text = string.Empty;
            comboBox11.Text = string.Empty;
            comboBox12.Text = string.Empty;
            textBox15.Text = string.Empty;
            textBox16.Text = string.Empty;
            textBox17.Text = string.Empty;
            textBox18.Text = string.Empty;
            textBox19.Text = string.Empty;
            textBox20.Text = string.Empty;
            textBox21.Text = string.Empty;
            textBox22.Text = string.Empty;
            textBox23.Text = string.Empty;
            comboBox10.Text = string.Empty;
        }

        private bool ExtraLineValidation(int inicialBus, int finalBus, int caseID)
        {
            List<string[]> matrix = new List<string[]>();
            matrix = Query("SELECT * FROM  `line` WHERE `caseID` = " + caseID.ToString() + " AND `inicialBusNumber` = " + inicialBus.ToString() + " AND `finalBusNumber` = " + finalBus.ToString() + "");
            if(checkQueryError(matrix))
            {
                if(matrix.Count>0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                return false;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox10, label41));
            int circuitNumber = 0;
            validationList.Add(ValidateAsInt(textBox23, label42, out circuitNumber));
            validationList.Add(ValidateAsNotNullText(textBox22, label40));
            double rating3 = 0;
            validationList.Add(ValidateAsDouble(textBox21,label35,out rating3));
            double rating2 = 0;
            validationList.Add(ValidateAsDouble(textBox20, label36, out rating2));
            double rating1 = 0;
            validationList.Add(ValidateAsDouble(textBox19, label37, out rating1));
            double suspectance = 0;
            validationList.Add(ValidateAsDouble(textBox18, label38, out suspectance));
            double reactance = 0;
            validationList.Add(ValidateAsDouble(textBox17, label39, out reactance));
            double resistence = 0;
            validationList.Add(ValidateAsDouble(textBox16, label34, out resistence));
            int sequencialNumber = 0;
            validationList.Add(ValidateAsInt(textBox15, label33, out sequencialNumber));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox12,label32));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox11, label31));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox9, label30));

            if(CompleteValidation(validationList))
            {
                int caseID = GetTransmissionCaseID(comboBox9.SelectedIndex);
                int areaID = GetAreaID(comboBox10.SelectedIndex);
                int finalBus = GetBusID(comboBox12.SelectedIndex, caseID);
                int inicialBus = GetBusID(comboBox11.SelectedIndex, caseID);

                // Check If inicial Bus and Final bus already on the system
                if (ExtraLineValidation(inicialBus, finalBus, caseID))
                {
                    string returnedMsg = Insert("INSERT INTO `sql583577`.`line` (`caseID`, `inicialBusNumber`, `finalBusNumber`, `sequencialNumber`, `length`, `resistence`, `reactance`, `shuntSusceptance`, `rating1`, `rating2`, `rating3`, `description`, `circuitoNumber`, `areaID`) VALUES ('" + caseID.ToString() + "', '" + inicialBus.ToString() + "', '" + finalBus.ToString() + "', '" + sequencialNumber.ToString() + "', NULL, '" + resistence.ToString().Replace(',', '.') + "', '" + reactance.ToString().Replace(',', '.') + "', '" + suspectance.ToString().Replace(',', '.') + "', '" + rating1.ToString().Replace(',', '.') + "', '" + rating2.ToString().Replace(',', '.') + "', '" + rating3.ToString().Replace(',', '.') + "', '" + textBox22.Text + "', '" + circuitNumber.ToString() + "', '" + areaID.ToString() + "');");
                    if (returnedMsg.ToLower().Equals("ok"))
                    {
                        showMsg(GetLabel(DictionaryFileName, "InsertSuccess"));
                        comboBox9.Text = string.Empty;
                        comboBox11.Text = string.Empty;
                        comboBox12.Text = string.Empty;
                        textBox15.Text = string.Empty;
                        textBox16.Text = string.Empty;
                        textBox17.Text = string.Empty;
                        textBox18.Text = string.Empty;
                        textBox19.Text = string.Empty;
                        textBox20.Text = string.Empty;
                        textBox21.Text = string.Empty;
                        textBox22.Text = string.Empty;
                        textBox23.Text = string.Empty;
                        comboBox10.Text = string.Empty;
                    }
                    else
                    {
                        ShowInsertError(returnedMsg);
                    }
                }
                else
                {
                    ShowError(985,GetLabel(DictionaryFileName,"Database"));
                }
            }

        }

        private void maskedTextBox1_MaskInputRejected_1(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void button13_Click_1(object sender, EventArgs e)
        {
            comboBox13.Text = string.Empty;
            textBox10.Text = string.Empty;
            richTextBox3.Text = string.Empty;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            validationList.Add(ValidateAsNotNullRichText(richTextBox3, label48));
            int lineType = 0;
            validationList.Add(ValidateAsInt(textBox10, label47, out lineType));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox13, label46));

            if(CompleteValidation(validationList))
            {
                int caseID = GetTransmissionCaseID(comboBox13.SelectedIndex);
                string returnedMsg = Insert("INSERT INTO `sql583577`.`linetype` (`ID`, `caseID`, `lineSpacing`, `phasing`, `conductorID`, `tapeShieldedConductorID`, `neutralConductorID`, `description`) VALUES ('" + lineType.ToString() + "', '" + caseID.ToString() + "', NULL, '', NULL, NULL, NULL, '" + richTextBox3.Text+ "');");
                if(returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName,"InsertSuccess"));
                    comboBox13.Text = string.Empty;
                    textBox10.Text = string.Empty;
                    richTextBox3.Text = string.Empty;
                }
                else
                {
                    ShowInsertError(returnedMsg);
                }
            }
        }

        private void comboBox14_TextChanged(object sender, EventArgs e)
        {
            if(comboBox14.SelectedIndex>-1)
            {
                int caseID = GetTransmissionCaseID(comboBox14.SelectedIndex);
                SetLineList(comboBox15, caseID);
                SetLineTypeList(comboBox16, caseID);
            }
            else
            {
                ClearComboBox(comboBox15);
                ClearComboBox(comboBox16);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            comboBox14.Text = string.Empty;
            comboBox15.Text = string.Empty;
            comboBox16.Text = string.Empty;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox16, label51));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox15, label50));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox14, label49));
            if(CompleteValidation(validationList))
            {
                int caseID = GetTransmissionCaseID(comboBox14.SelectedIndex);
                int lineID = GetLineID(comboBox15.SelectedIndex,caseID);
                int lineTypeID = GetLineTypeID(comboBox16.SelectedIndex, caseID);
                string returnedMsg = Insert("INSERT INTO `sql583577`.`line_linetype` (`caseID`, `lineID`, `lineTypeID`) VALUES ('"+caseID.ToString()+"', '"+lineID.ToString()+"', '"+lineTypeID.ToString()+"');");
                if(returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName, "InsertSuccess"));
                    comboBox14.Text = string.Empty;
                    comboBox15.Text = string.Empty;
                    comboBox16.Text = string.Empty;
                }
                else
                {
                    ShowInsertError(returnedMsg);
                }
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            comboBox17.Text = string.Empty;
            textBox14.Text = string.Empty;
            textBox24.Text = string.Empty;
            richTextBox4.Text = string.Empty;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            validationList.Add(ValidateAsNotNullRichText(richTextBox4, label55));
            int sequencialNumber;
            validationList.Add(ValidateAsInt(textBox24, label54, out sequencialNumber));
            int lossZone;
            validationList.Add(ValidateAsInt(textBox14, label53, out lossZone));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox17, label52));

            if (CompleteValidation(validationList))
            {
                int caseID = GetTransmissionCaseID(comboBox17.SelectedIndex);
                string returnedMsg = Insert("INSERT INTO `sql583577`.`lossZone` (`zoneNumber`, `caseID`, `description`, `sequencialNumber`) VALUES ('" + lossZone.ToString() + "', '"+caseID.ToString()+"', '"+richTextBox4.Text+"', '"+sequencialNumber.ToString()+"');");
                if (returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName, "InsertSuccess"));
                    comboBox17.Text = string.Empty;
                    textBox14.Text = string.Empty;
                    textBox24.Text = string.Empty;
                    richTextBox4.Text = string.Empty;
                }
                else
                {
                    ShowInsertError(returnedMsg);
                }
            }
        }

        private void ClearComboBox(ComboBox comboBox)
        {
            comboBox.Items.Clear();
            comboBox.Text = string.Empty;
        }

        private void comboBox18_TextChanged(object sender, EventArgs e)
        {
            if(comboBox18.SelectedIndex>-1)
            {
                int caseID = GetTransmissionCaseID(comboBox18.SelectedIndex);
                SetBusList(comboBox19, caseID);
                SetLossZoneList(comboBox20, caseID);
            }
            else
            {
                ClearComboBox(comboBox19);
                ClearComboBox(comboBox20);
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            comboBox18.Text = string.Empty;
            comboBox19.Text = string.Empty;
            comboBox20.Text = string.Empty;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox20, label59));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox19, label58));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox18, label57));

            if (CompleteValidation(validationList))
            {
                int caseID = GetTransmissionCaseID(comboBox18.SelectedIndex);
                int busID = GetBusID(comboBox19.SelectedIndex, caseID);
                int lossZoneID = GetLossZoneID(comboBox20.SelectedIndex, caseID);
                string returnedMsg = Insert("INSERT INTO `sql583577`.`bus_losszone` (`busNumber`, `lossZoneID`, `caseID`) VALUES ('" + busID.ToString() + "', '" + lossZoneID.ToString() + "', '" + caseID.ToString() + "');");
                if(returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName, "InsertSuccess"));
                    comboBox18.Text = string.Empty;
                    comboBox19.Text = string.Empty;
                    comboBox20.Text = string.Empty;
                }
                else
                {
                    ShowInsertError(returnedMsg);
                }
            }
        }

        private void comboBox18_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox23_TextChanged(object sender, EventArgs e)
        {
            if(comboBox23.SelectedIndex>-1)
            {
                int caseID = GetTransmissionCaseID(comboBox23.SelectedIndex);
                SetLineList(comboBox22, caseID);
                SetLossZoneList(comboBox21, caseID);
            }
            else
            {
                ClearComboBox(comboBox22);
                ClearComboBox(comboBox21);
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            comboBox23.Text = string.Empty;
            comboBox22.Text = string.Empty;
            comboBox21.Text = string.Empty;
        }

        private void button21_Click(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox21, label60));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox22, label61));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox23, label62));

            if(CompleteValidation(validationList))
            {
                int caseID = GetTransmissionCaseID(comboBox23.SelectedIndex);
                int LineID = GetLineID(comboBox22.SelectedIndex, caseID);
                int lossZone = GetLossZoneID(comboBox21.SelectedIndex, caseID);

                string returnedMsg = Insert("INSERT INTO `sql583577`.`line_losszone` (`caseID`, `lineID`, `lossZoneID`) VALUES ('" + caseID.ToString() + "', '" + LineID.ToString() + "', '" + lossZone.ToString() + "');");
                if (returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName, "InsertSuccess"));
                    comboBox23.Text = string.Empty;
                    comboBox22.Text = string.Empty;
                    comboBox21.Text = string.Empty;
                }
                else
                {
                    ShowInsertError(returnedMsg);
                }
            }
        }

        private void comboBox24_TextChanged(object sender, EventArgs e)
        {
            if(comboBox24.SelectedIndex>-1)
            {
                int caseID = GetTransmissionCaseID(comboBox24.SelectedIndex);
                SetBusList(comboBox25, caseID);
            }
            else
            {
                ClearComboBox(comboBox25);
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            comboBox24.Text = string.Empty;
            comboBox25.Text = string.Empty;
            textBox25.Text = string.Empty;
            textBox26.Text = string.Empty;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            double reactivePower = 0;
            validationList.Add(ValidateAsDouble(textBox26, label67, out reactivePower));
            double activePower = 0;
            validationList.Add(ValidateAsDouble(textBox25, label66, out activePower));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox25, label65));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox24, label64));

            if (CompleteValidation(validationList))
            {
                int caseID = GetTransmissionCaseID(comboBox24.SelectedIndex);
                int busID = GetBusID(comboBox25.SelectedIndex, caseID);
                string returnedMsg = Insert("INSERT INTO `sql583577`.`load` (`busID`, `caseID`, `activePowerAtPhaseA`, `activePowerAtPhaseB`, `activePowerAtPhaseC`, `reactivePowerAtPhaseA`, `reactivePowerAtPhaseB`, `reactivePowerAtPhaseC`, `modelLoadID`) VALUES ('" + busID.ToString() + "', '" + caseID.ToString() + "', '" + activePower.ToString() + "', '" + activePower.ToString() + "', '" + activePower.ToString() + "', '" + reactivePower.ToString() + "', '" + reactivePower.ToString() + "', '" + reactivePower.ToString() + "', NULL);");
                if (returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName, "InsertSuccess"));
                    comboBox24.Text = string.Empty;
                    comboBox25.Text = string.Empty;
                    textBox25.Text = string.Empty;
                    textBox26.Text = string.Empty;
                }
                else
                {
                    ShowInsertError(returnedMsg);
                }
            }
        }

        private void comboBox26_TextChanged(object sender, EventArgs e)
        {
            if(comboBox26.SelectedIndex>-1)
            {
                int caseID = GetTransmissionCaseID(comboBox26.SelectedIndex);
                SetBusList(comboBox27, caseID);
            }
            else
            {
                ClearComboBox(comboBox27);
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            comboBox26.Text = string.Empty;
            comboBox27.Text = string.Empty;
            textBox27.Text = string.Empty;
            textBox28.Text = string.Empty;
            richTextBox5.Text = string.Empty;
        }

        private void button24_Click(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            validationList.Add(ValidateAsNotNullRichText(richTextBox5, label74));
            double susceptance = 0;
            validationList.Add(ValidateAsDouble(textBox28, label73, out susceptance));
            double conductance = 0;
            validationList.Add(ValidateAsDouble(textBox27, label72, out conductance));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox27, label71));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox26, label70));
            if (CompleteValidation(validationList))
            {
                int caseID = GetTransmissionCaseID(comboBox26.SelectedIndex);
                int busID = GetBusID(comboBox27.SelectedIndex, caseID);
                string returnedMsg = Insert("INSERT INTO `sql583577`.`shuntelement` (`busID`, `caseID`, `conductance`, `susceptance`, `description`, `activePowerAtPhaseA`, `activePowerAtPhaseB`, `activePowerAtPhaseC`, `reactivePowerAtPhaseA`, `reactivePowerAtPhaseB`, `reactivePowerAtPhaseC`) VALUES ('" + busID.ToString() + "', '" + caseID.ToString() + "', '" + conductance.ToString().Replace(',', '.') + "', '" + susceptance.ToString().Replace(',', '.') + "', '" + richTextBox5.Text + "', '', '', '', '', '', '');");
                if(returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName, "InsertSuccess"));
                    comboBox26.Text = string.Empty;
                    comboBox27.Text = string.Empty;
                    textBox27.Text = string.Empty;
                    textBox28.Text = string.Empty;
                    richTextBox5.Text = string.Empty;
                }
                else
                {
                    ShowInsertError(returnedMsg);
                }
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            comboBox28.Text = string.Empty;
            textBox29.Text = string.Empty;
            textBox30.Text = string.Empty;
            textBox31.Text = string.Empty;
            textBox32.Text = string.Empty;
            textBox33.Text = string.Empty;
            textBox34.Text = string.Empty;
            textBox35.Text = string.Empty;
        }

        private void button26_Click(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            double voltagelimitMax = 0;
            validationList.Add(ValidateAsDouble(textBox35, label82, out voltagelimitMax));
            double voltagelimitMin = 0;
            validationList.Add(ValidateAsDouble(textBox34, label81, out voltagelimitMin));
            double stepsize = 0;
            validationList.Add(ValidateAsDouble(textBox33, label80, out stepsize));
            double voltageratiolimmax = 0;
            validationList.Add(ValidateAsDouble(textBox32, label79, out voltageratiolimmax));
            double voltageratiolimmin = 0;
            validationList.Add(ValidateAsDouble(textBox31, label78, out voltageratiolimmin));
            double finalPhaseAngle = 0;
            validationList.Add(ValidateAsDouble(textBox30, label77, out finalPhaseAngle));
            double finalVoltageRatio = 0;
            validationList.Add(ValidateAsDouble(textBox29, label76, out finalVoltageRatio));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox28, label75));

            if(CompleteValidation(validationList))
            {
                int caseID = GetTransmissionCaseID(comboBox28.SelectedIndex);
                string returnedMsg = Insert("INSERT INTO `sql583577`.`transformer` (`ID`, `caseID`, `finalVoltageRatio`, `finalPhaseAngle`, `VoltageRatioMinLimit`, `VoltageRatioMaxLimit`, `kV-high`, `kV-low`, `stepSize`, `kVAPower`, `PowerMinLimit`, `PowerMaxLimit`, `name`, `resistenceVar`, `reactance`) VALUES (NULL, '" + caseID + "', '" + finalVoltageRatio.ToString().Replace(',', '.') + "', '" + finalPhaseAngle.ToString().Replace(',', '.') + "', '" + voltagelimitMin.ToString().Replace(',', '.') + "', '" + voltagelimitMax.ToString().Replace(',', '.') + "', '0', '0', '0', '0', '" + voltageratiolimmin.ToString().Replace(',', '.') + "', '" + voltageratiolimmax.ToString().Replace(',', '.') + "', '', '0', '0');");
                if(returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName, "InsertSuccess"));
                    comboBox28.Text = string.Empty;
                    textBox29.Text = string.Empty;
                    textBox30.Text = string.Empty;
                    textBox31.Text = string.Empty;
                    textBox32.Text = string.Empty;
                    textBox33.Text = string.Empty;
                    textBox34.Text = string.Empty;
                    textBox35.Text = string.Empty;
                }
                else
                {
                    ShowInsertError(returnedMsg);
                }
            }
        }

        private void comboBox29_TextChanged(object sender, EventArgs e)
        {
            if(comboBox29.SelectedIndex>-1)
            {
                int caseID = GetTransmissionCaseID(comboBox29.SelectedIndex);
                SetLineList(comboBox30, caseID);
                SetTransformerList(comboBox31, caseID);
            }
            else
            {
                ClearComboBox(comboBox30);
                ClearComboBox(comboBox31);
            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            comboBox29.Text = string.Empty;
            comboBox30.Text = string.Empty;
            comboBox31.Text = string.Empty;
        }

        private void button28_Click(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox31, label85));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox30, label84));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox29, label83));

            if(CompleteValidation(validationList))
            {
                int caseID = GetTransmissionCaseID(comboBox29.SelectedIndex);
                int lineID = GetLineID(comboBox30.SelectedIndex, caseID);
                int transfID = GetTransformerID(comboBox31.SelectedIndex, caseID);

                string returnedMsg = Insert("INSERT INTO `sql583577`.`transformer_line` (`transformerID`, `lineID`, `caseID`) VALUES ('"+transfID.ToString()+"', '"+lineID.ToString()+"', '"+caseID.ToString()+"');");
                if(returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName, "InsertSuccess"));
                    comboBox29.Text = string.Empty;
                    comboBox30.Text = string.Empty;
                    comboBox31.Text = string.Empty;
                }
                else
                {
                    ShowInsertError(returnedMsg);
                }
            }
        }

        private void comboBox32_TextChanged(object sender, EventArgs e)
        {
            if(comboBox32.SelectedIndex > -1)
            {
                int caseID = GetTransmissionCaseID(comboBox32.SelectedIndex);
                SetBusList(comboBox33, caseID);
                SetBusList(comboBox34, caseID);
            }
            else
            {
                ClearComboBox(comboBox33);
                ClearComboBox(comboBox34);
            }
        }

        private void button31_Click(object sender, EventArgs e)
        {
            comboBox32.Text = string.Empty;
            comboBox33.Text = string.Empty;
            comboBox34.Text = string.Empty;
        }

        private void button30_Click(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox34, label88));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox33, label87));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox32, label86));

            if(CompleteValidation(validationList))
            {
                int caseID = GetTransmissionCaseID(comboBox32.SelectedIndex);
                int busID_controlled = GetBusID(comboBox33.SelectedIndex, caseID);
                int busID_controller = GetBusID(comboBox34.SelectedIndex, caseID);

                string returnedMsg = Insert("INSERT INTO `sql583577`.`bus_bus_control` (`busController`, `busControlled`, `caseID`) VALUES ('"+busID_controller.ToString()+"', '"+busID_controlled.ToString()+"', '"+caseID.ToString()+"');");
                if(returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName, "InsertSuccess"));
                    comboBox32.Text = string.Empty;
                    comboBox33.Text = string.Empty;
                    comboBox34.Text = string.Empty;
                }
                else
                {
                    ShowInsertError(returnedMsg);
                }
            }
        }

        private void comboBox35_TextChanged(object sender, EventArgs e)
        {
            if(comboBox35.SelectedIndex > -1)
            {
                int caseID = GetTransmissionCaseID(comboBox35.SelectedIndex);
                SetBusList(comboBox36, caseID);
                SetLineList(comboBox37, caseID);
            }
            else
            {
                ClearComboBox(comboBox36);
                ClearComboBox(comboBox37);
            }
        }

        private void button33_Click(object sender, EventArgs e)
        {
            comboBox35.Text = string.Empty;
            comboBox36.Text = string.Empty;
            comboBox37.Text = string.Empty;
            textBox36.Text = string.Empty;
        }

        private void button32_Click(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            int side = 0;
            validationList.Add(ValidateAsInt(textBox36, label92, out side));

            if((side == 0)||(side == 1))
            {
                validationList.Add(true);
            }
            else
            {
                validationList.Add(false);
                ShowError(983, label92.Text);
            }

            validationList.Add(ValidateAsSelectedfromCombobox(comboBox37, label91));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox36, label90));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox35, label89));

            if(CompleteValidation(validationList))
            {
                int caseID = GetTransmissionCaseID(comboBox35.SelectedIndex);
                int busID = GetBusID(comboBox36.SelectedIndex, caseID);
                int lineID = GetLineID(comboBox37.SelectedIndex, caseID);

                string returnedMsg = Insert("INSERT INTO `sql583577`.`bus_line_control` (`caseID`, `lineID`, `busID`, `side`) VALUES ('"+caseID.ToString()+"', '"+lineID.ToString()+"', '"+busID.ToString()+"', '"+side.ToString()+"');");
                if(returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName, "InsertSuccess"));
                    comboBox35.Text = string.Empty;
                    comboBox36.Text = string.Empty;
                    comboBox37.Text = string.Empty;
                    textBox36.Text = string.Empty;
                }
                else
                {
                    ShowInsertError(returnedMsg);
                }
            }
        }

        private void comboBox38_TextChanged(object sender, EventArgs e)
        {
            if(comboBox38.SelectedIndex >-1)
            {
                int caseID = GetTransmissionCaseID(comboBox38.SelectedIndex);
                SetBusList(comboBox39, caseID);
            }
            else
            {
                ClearComboBox(comboBox39);
            }
        }

        private void button35_Click(object sender, EventArgs e)
        {
            comboBox38.Text = string.Empty;
            comboBox39.Text = string.Empty;
            textBox37.Text = string.Empty;
            textBox38.Text = string.Empty;

        }

        private void button34_Click(object sender, EventArgs e)
        {
            List<bool> validationList = new List<bool>();
            double reactivePower = 0;
            validationList.Add(ValidateAsDouble(textBox38,label97, out reactivePower));
            double activePower = 0;
            validationList.Add(ValidateAsDouble(textBox37,label96, out activePower));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox39,label95));
            validationList.Add(ValidateAsSelectedfromCombobox(comboBox38,label94));

            if(CompleteValidation(validationList))
            {
                int caseID = GetTransmissionCaseID(comboBox38.SelectedIndex);
                int busID = GetBusID(comboBox39.SelectedIndex,caseID);

                string returnedMsg = Insert("INSERT INTO `sql583577`.`generation` (`busID`, `caseID`, `activePower`, `reactivePower`) VALUES ('" + busID.ToString() + "', '" + caseID.ToString() + "', '" + textBox37.Text + "', '" + textBox38.Text + "');");
                if(returnedMsg.ToLower().Equals("ok"))
                {
                    showMsg(GetLabel(DictionaryFileName, "InsertSuccess"));
                    comboBox38.Text = string.Empty;
                    comboBox39.Text = string.Empty;
                    textBox37.Text = string.Empty;
                    textBox38.Text = string.Empty;
                }
                else
                {
                    ShowInsertError(returnedMsg);
                }
            }
        }
    }
}