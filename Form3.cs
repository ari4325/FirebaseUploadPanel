using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office2013.PowerPoint.Roaming;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Panel
{
    public partial class Form3 : Form
    {

        System.Collections.ArrayList keys;
            IExcelDataReader excelReader;
        public Form3()
        {
            InitializeComponent();
            keys = new System.Collections.ArrayList();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                var HTTPrequest = (HttpWebRequest)WebRequest.Create("https://nefs-app.firebaseio.com/Users/.json");
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

                    xlWorkSheet.Cells[1, 1] = "Aadhar";
                    xlWorkSheet.Cells[1, 2] = "Address";
                    xlWorkSheet.Cells[1, 3] = "Mobile No.";
                    xlWorkSheet.Cells[1, 4] = "Name";
                    xlWorkSheet.Cells[1, 5] = "PAN No.";
                    xlWorkSheet.Cells[1, 6] = "Password";

                    int row = 2;
                    int coloumn = 1;

                    foreach (var key in keys)
                    {
                        string link = "https://nefs-app.firebaseio.com/Users/" + (string)key + "/.json";
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

                    xlWorkBook.SaveAs("E:\\PanelUsers.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);

                    MessageBox.Show("Excel file created , you can find the file d:\\PanelUsers.xls");

                    dataGridView1.Visible = true;
                    dataGridView1.DataSource = ReadExcel("E:\\PanelUsers.xls", ".xls");
                }
            }catch(Exception ex)
            {
                MessageBox.Show("Failed to connect to firebase services");
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

        public System.Data.DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            System.Data.DataTable dtexcel = new System.Data.DataTable();
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

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow item in this.dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.RemoveAt(item.Index);
            }

           
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
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

            xlWorkBook.SaveAs("E:\\PanelUsers.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file d:\\PanelUsers.xls");
        }

        private void button3_Click(object sender, EventArgs e)
        {

            var request = WebRequest.CreateHttp("https://nefs-app.firebaseio.com/Users/.json");
            request.Method = "DELETE";
            request.ContentType = "application/json";
            var response = request.GetResponse();

            FileStream stream = File.Open("E:\\PanelUsers.xls", FileMode.Open, FileAccess.Read);

            // Reading from a binary Excel file ('97-2003 format; *.xls)
            if (System.IO.Path.GetExtension("E:\\PanelUsers.xls").CompareTo(".xls") == 0)
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
                        Aadhar = result.Tables[ind].Rows[row_no][i].ToString(),
                        Address = result.Tables[ind].Rows[row_no][i + 1].ToString(),
                        Mobile = result.Tables[ind].Rows[row_no][i + 2].ToString(),
                        Name = result.Tables[ind].Rows[row_no][i + 3].ToString(),
                        Pan = result.Tables[ind].Rows[row_no][i + 4].ToString(),
                        Passowd = result.Tables[ind].Rows[row_no][i + 5].ToString()
                    }) ;

                    request = WebRequest.CreateHttp("https://nefs-app.firebaseio.com/Users/.json");
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

            MessageBox.Show("Uploaded Successfully");


        }
    }

    public class JsonFieldsCollector
    {
        private readonly Dictionary<string, JValue> fields;

        public JsonFieldsCollector(JToken token)
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
