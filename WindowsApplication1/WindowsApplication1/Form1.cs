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
            matrix = databaseAccess.query(null, GetLabel("config.ini", "Host"), GetLabel("config.ini", "UserID"), GetLabel("config.ini", "DatabaseName"), maskedTextBox1.Text);
            matrix = CaseDistribuctionFiltering(matrix);
            return matrix;
        }

        private List<string[]> GetTransmissionMatrix()
        {
            DatabaseAccess.Query databaseAccess = new DatabaseAccess.Query();
            List<string[]> matrix = new List<string[]>();
            matrix = databaseAccess.query(null, GetLabel("config.ini", "Host"), GetLabel("config.ini", "UserID"), GetLabel("config.ini", "DatabaseName"), maskedTextBox1.Text);
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
            this.Size = new Size(900, 543);
            panel1.Location = new Point(12, 14);
            panel1.Size = new Size(857, 94);
            panel2.Location = new Point(12, 427);
            panel2.Size = new Size(857, 25);
            panel3.Location = new Point(12, 449);
            panel3.Size = new Size(857, 39);
        }


        private void SetPanelLocation(int PanelLocationX, int PanelLocationY, int PanelLocationH, int PanelLocationW)
        {
            panel4.Location = new Point(PanelLocationX, PanelLocationY);
            panel4.Size = new Size(PanelLocationH, PanelLocationW);
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            label5.Text = GetLabel(DictionaryFileName, "Error.RunMsg");

            // Correct the place of painels. 
            int PanelLocationX = 12;
            int PanelLocationY = 113;
            int PanelLocationH = 857;
            int PanelLocationW = 309;

            // Location of panels with forms.
            SetPanelLocation(PanelLocationX, PanelLocationY, PanelLocationH, PanelLocationW);

            label1.Text = GetLabel(DictionaryFileName, "CategoryLabel") + ":";
            label2.Text = GetLabel(DictionaryFileName, "OperationLabel") + ":";
            label3.Text = GetLabel(DictionaryFileName, "DataLabel") + ":";
            radioButton1.Text = GetLabel(DictionaryFileName, "OperationInformation.Insert");
            radioButton2.Text = GetLabel(DictionaryFileName, "OperationInformation.Remove");
            radioButton3.Text = GetLabel(DictionaryFileName, "OperationInformation.Update");
            radioButton4.Text = GetLabel(DictionaryFileName, "OperationInformation.Query");

            button1.Text = GetLabel(DictionaryFileName, "RunButton");

            label4.Text = GetLabel(DictionaryFileName, "UserMsgLabel") + ":";
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

        private void button1_Click(object sender, EventArgs e)
        {
            
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
    }
}