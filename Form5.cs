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
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" ||
                textBox2.Text != "")
            {
                var json = Newtonsoft.Json.JsonConvert.SerializeObject(new
                {
                    NoticeText = textBox2.Text
                });

                var request = WebRequest.CreateHttp("https://nefs-app.firebaseio.com/Notice/"+textBox1.Text+"/.json");
                request.Method = "POST";
                request.ContentType = "application/json";
                var buffer = Encoding.UTF8.GetBytes(json);
                request.ContentLength = buffer.Length;
                request.GetRequestStream().Write(buffer, 0, buffer.Length);
                var response = request.GetResponse();
                json = (new StreamReader(response.GetResponseStream())).ReadToEnd();

                MessageBox.Show("Uploaded Successfully");
            }
            else
            {
                MessageBox.Show("Fields cannot be empty while registering user");
            }
        }
    }
}
