using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Common;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using System.Collections;

namespace Panel
{
    public partial class Form1 : Form
    {
        string jsonString;
        ArrayList keys = new System.Collections.ArrayList();
        IExcelDataReader excelReader;
        String filePath;
        String fileExt;
        public Form1()
        {
            InitializeComponent();
            jsonString = string.Empty;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            filePath = string.Empty;
            fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                fileExt = System.IO.Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel = new DataTable();
                        dtExcel = ReadExcel(filePath, fileExt); //read excel file 
                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dtExcel;

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
        }

        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable 



                }
                catch { }
            }

            return dtexcel;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var request = WebRequest.CreateHttp("https://nefs-app.firebaseio.com/Leads/.json");
            request.Method = "DELETE";
            request.ContentType = "application/json";
            var response = request.GetResponse();


            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

            // Reading from a binary Excel file ('97-2003 format; *.xls)
            if (fileExt.CompareTo(".xls") == 0)
            {
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }

            // DataSet - The result of each spreadsheet will be created in the result.Tables
            DataSet result = excelReader.AsDataSet();

            excelReader.Close();

            var ind = result.Tables[0].TableName.ToString();

            int row_no = 0;
            int i = 0;

            while (row_no < result.Tables[ind].Rows.Count) // ind is the index of table
                                                           // (sheet name) which you want to convert to csv
            {
                if (row_no != 0)
                {
                    var json = Newtonsoft.Json.JsonConvert.SerializeObject(new
                    {
                        AssignedExecutive = result.Tables[ind].Rows[row_no][i].ToString(),
                        Name = result.Tables[ind].Rows[row_no][i+1].ToString(),
                        Address = result.Tables[ind].Rows[row_no][i+2].ToString(),
                        City = result.Tables[ind].Rows[row_no][i + 3].ToString(),
                        State = result.Tables[ind].Rows[row_no][i + 4].ToString(),
                        Pin = result.Tables[ind].Rows[row_no][i + 5].ToString(),
                        Mobile = result.Tables[ind].Rows[row_no][i + 6].ToString(),
                        Vin = result.Tables[ind].Rows[row_no][i + 7].ToString(),
                        RegNo = result.Tables[ind].Rows[row_no][i + 8].ToString(),
                        Model = result.Tables[ind].Rows[row_no][i + 9].ToString(),
                        SaleDate = result.Tables[ind].Rows[row_no][i + 10].ToString(),
                        NCB = result.Tables[ind].Rows[row_no][i + 11].ToString(),
                        PreviousInsurance = result.Tables[ind].Rows[row_no][i + 12].ToString(),
                        ExecutiveRemark1 = result.Tables[ind].Rows[row_no][i + 13].ToString(),
                        ExecutiveRemark2 = result.Tables[ind].Rows[row_no][i + 14].ToString(),
                        TeamLeaderRemark1 = result.Tables[ind].Rows[row_no][i + 15].ToString(),
                        TeamLeaderRemark2 = result.Tables[ind].Rows[row_no][i + 16].ToString(),
                        Message = result.Tables[ind].Rows[row_no][i+17].ToString()
                    });

                    request = WebRequest.CreateHttp("https://nefs-app.firebaseio.com/Leads/.json");
                    request.Method = "POST";
                    request.ContentType = "application/json";
                    var buffer = Encoding.UTF8.GetBytes(json);
                    request.ContentLength = buffer.Length;
                    request.GetRequestStream().Write(buffer, 0, buffer.Length);
                    response = request.GetResponse();
                    json = (new StreamReader(response.GetResponseStream())).ReadToEnd();
                }
                 row_no++;
            }




            DateTime date = DateTime.Now;
            string Date = date.ToString("yyyy:MM:dd");
            string Time = date.ToString("HH:mm:ss");

           
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var form = new Form2();
            form.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var form = new Form3();
            form.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow item in this.dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.RemoveAt(item.Index);
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            int i = 0;
            int j = 0;

            for (i = 0; i <= dataGridView1.RowCount - 1; i++)
            {
                for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                {
                    DataGridViewCell cell = dataGridView1[j, i];
                    xlWorkSheet.Cells[i + 1, j + 1] = cell.Value;
                }
            }

            xlWorkBook.SaveAs("D:\\UpdatedExcel.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file at D:\\UpdatedExcel.xls");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
                var HTTPrequest = (HttpWebRequest)WebRequest.Create("https://nefs-app.firebaseio.com/Leads/.json");
                var Response = (HttpWebResponse)HTTPrequest.GetResponse();
                var streamReader = new StreamReader(Response.GetResponseStream()).ReadToEnd();
                String jsonString = streamReader;
                var json = JToken.Parse(jsonString);
                var fieldsCollector = new JsonFieldsCollector(json);
                var fields = fieldsCollector.GetAllFields();
                int i = 0;
                foreach (var field in fields)
                {
                    if (checkKey(field.Key.Substring(0, 20)))
                    {
                        field.Key.Substring(0, 20);
                        keys.Add(field.Key.Substring(0, 20));
                        i++;
                    }
                    else
                    {

                    }

                }

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return;
                }
                else
                {
                    Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                    Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;

                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    xlWorkSheet.Cells[1, 1] = "Address";
                    xlWorkSheet.Cells[1, 2] = "Assigned Executive";
                    xlWorkSheet.Cells[1, 3] = "City";
                    xlWorkSheet.Cells[1, 4] = "Executive Remark1";
                    xlWorkSheet.Cells[1, 5] = "Executive Remark2";
                    xlWorkSheet.Cells[1, 6] = "Executive Remark2";
                    xlWorkSheet.Cells[1, 7] = "Mobile";
                    xlWorkSheet.Cells[1, 8] = "Model";
                    xlWorkSheet.Cells[1, 9] = "NCB";
                    xlWorkSheet.Cells[1, 10] = "Name";
                    xlWorkSheet.Cells[1, 11] = "Pin";
                    xlWorkSheet.Cells[1, 12] = "Previous Insurance";
                    xlWorkSheet.Cells[1, 13] = "Registration Number";
                    xlWorkSheet.Cells[1, 14] = "Sale Date";
                    xlWorkSheet.Cells[1, 15] = "State";
                    xlWorkSheet.Cells[1, 16] = "Status";
                    xlWorkSheet.Cells[1, 17] = "Executive Remark1";
                    xlWorkSheet.Cells[1, 18] = "Team Leader Remark2";
                    xlWorkSheet.Cells[1, 19] = "Time";
                    xlWorkSheet.Cells[1, 20] = "VIN";

                    int row = 2;
                    int coloumn = 1;

                    foreach (var key in keys)
                    {
                        string link = "https://nefs-app.firebaseio.com/Leads/" + (string)key + "/.json";
                        var hTTPrequest = (HttpWebRequest)WebRequest.Create(link);
                        var response = (HttpWebResponse)hTTPrequest.GetResponse();
                        var streamReader1 = new StreamReader(response.GetResponseStream()).ReadToEnd();
                        String jsonString1 = streamReader1;

                        var json1 = JToken.Parse(jsonString1);
                        var fieldsCollector1 = new JsonFieldsCollector(json1);
                        var fields1 = fieldsCollector1.GetAllFields();

                        foreach (var field1 in fields1)
                        {
                            xlWorkSheet.Cells[row, coloumn] = field1.Value;
                            coloumn++;
                        }
                        row++;
                        coloumn = 1;


                    }

                    xlWorkBook.SaveAs("D:\\ImportLeads.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);

                    MessageBox.Show("Excel file created , you can find the file d:\\PanelUsers.xls");

                    dataGridView1.Visible = true;
                    dataGridView1.DataSource = ReadExcel("D:\\ImportLeads.xls", ".xls");
                }
            

        }

        bool checkKey(String s)
        {
            if (keys.IndexOf(s) == -1)
            {
                return true;
            }
            return false;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            var policyForm = new Form4();
            policyForm.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            var noticeForm = new Form5();
            noticeForm.Show();
        }
    }

    public class JsonFieldsCollector1
    {
        private readonly Dictionary<string, JValue> fields;

        public JsonFieldsCollector1(JToken token)
        {
            fields = new Dictionary<string, JValue>();
            CollectFields(token);
        }

        private void CollectFields(JToken jToken)
        {
            switch (jToken.Type)
            {
                case JTokenType.Object:
                    foreach (var child in jToken.Children<JProperty>())
                        CollectFields(child);
                    break;
                case JTokenType.Array:
                    foreach (var child in jToken.Children())
                        CollectFields(child);
                    break;
                case JTokenType.Property:
                    CollectFields(((JProperty)jToken).Value);
                    break;
                default:
                    fields.Add(jToken.Path, (JValue)jToken);
                    break;
            }
        }

        public IEnumerable<KeyValuePair<string, JValue>> GetAllFields() => fields;
    }
}
