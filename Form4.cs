using ExcelDataReader;
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
    public partial class Form4 : Form
    {
        IExcelDataReader excelReader;
        String filePath;
        String fileExt;
        public Form4()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var request = WebRequest.CreateHttp("https://nefs-app.firebaseio.com/Policies/.json");
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
                        Name = result.Tables[ind].Rows[row_no][i].ToString(),
                        Address = result.Tables[ind].Rows[row_no][i + 1].ToString(),
                        City = result.Tables[ind].Rows[row_no][i + 2].ToString(),
                        State = result.Tables[ind].Rows[row_no][i + 3].ToString(),
                        Pin = result.Tables[ind].Rows[row_no][i + 4].ToString(),
                        Mobile = result.Tables[ind].Rows[row_no][i + 5].ToString(),
                        Vin = result.Tables[ind].Rows[row_no][i + 6].ToString(),
                        regNo = result.Tables[ind].Rows[row_no][i + 7].ToString(),
                        Model = result.Tables[ind].Rows[row_no][i + 8].ToString(),
                        ClosureDate = result.Tables[ind].Rows[row_no][i + 9].ToString(),
                        InsuranceCompany = result.Tables[ind].Rows[row_no][i + 10].ToString(),
                        ExecutiveName = result.Tables[ind].Rows[row_no][i + 11].ToString(),
                        PremiumAccount = result.Tables[ind].Rows[row_no][i + 12].ToString(),
                        OwnDamageAmt = result.Tables[ind].Rows[row_no][i + 13].ToString(),
                        RiskStartDate= result.Tables[ind].Rows[row_no][i + 14].ToString()
                    
                    });

                    request = WebRequest.CreateHttp("https://nefs-app.firebaseio.com/Policies/.json");
                    request.Method = "POST";
                    request.ContentType = "application/json";
                    var buffer = Encoding.UTF8.GetBytes(json);
                    request.ContentLength = buffer.Length;
                    request.GetRequestStream().Write(buffer, 0, buffer.Length);
                    response = request.GetResponse();
                    json = (new StreamReader(response.GetResponseStream())).ReadToEnd();
                }
                row_no++;
                MessageBox.Show("Uploaded Successfully");
            }
        }

        private void button4_Click(object sender, EventArgs e)
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

        private void button2_Click(object sender, EventArgs e)
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

            xlWorkBook.SaveAs("D:\\ImportedPolicies", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file" + filePath);
        }
    }
}
