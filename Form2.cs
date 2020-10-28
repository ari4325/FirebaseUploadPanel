using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Panel
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" ||
                textBox2.Text != "" ||
                textBox3.Text != "" ||
                textBox4.Text != "" ||
                textBox5.Text != "" ||
                textBox6.Text != ""){
                var json = Newtonsoft.Json.JsonConvert.SerializeObject(new
                {
                    UserId = textBox1.Text,
                    Mobile = textBox2.Text,
                    Address = textBox3.Text,
                    PAN = textBox4.Text,
                    Aadhar = textBox5.Text,
                    Password = textBox6.Text
                });

                var request = WebRequest.CreateHttp("https://nefs-app.firebaseio.com/Users/.json");
                request.Method = "POST";
                request.ContentType = "application/json";
                var buffer = Encoding.UTF8.GetBytes(json);
                request.ContentLength = buffer.Length;
                request.GetRequestStream().Write(buffer, 0, buffer.Length);
                var response = request.GetResponse();
                json = (new StreamReader(response.GetResponseStream())).ReadToEnd();

                MessageBox.Show("User Registered Successfully");
            }
            else
            {
                MessageBox.Show("Fields cannot be empty while registering user");
            }
        }
    }
}
