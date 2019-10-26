using ExcelReader.DataAccess;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public DataTable data { get; set; }
        private void Button1_Click(object sender, EventArgs e)
        {
            ExcelDataReader dataReader = new ExcelDataReader();
            data = dataReader.GetAll();
            dataGridView1.DataSource = data;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (data != null || data.Rows.Count > 0)
                {
                    using (StreamWriter file = File.CreateText("../OtoAdmin/DataFiles/siparisler.json"))
                    {
                        JsonSerializer serializer = new JsonSerializer();
                        serializer.Serialize(file, data);
                    }
                }
                else
                {
                    MessageBox.Show("Lütfen önce verileri çekiniz.");
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Lütfen önce verileri çekiniz.");
            }
        }
    }
}
