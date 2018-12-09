using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp3
{
    public partial class Form2 : Form
    {
        DataTable table = new DataTable("tbl");
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            table.WriteXml(@"C:\Users\qwerty\source\repos\WindowsFormsApp3\WindowsFormsApp3\XMLFile1.xml", XmlWriteMode.WriteSchema);
            MessageBox.Show("Data Exported");
        }
    }
}
