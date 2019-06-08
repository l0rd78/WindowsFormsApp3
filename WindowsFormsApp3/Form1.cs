using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using System.Data.OleDb;
using DataTable = System.Data.DataTable;
using Microsoft.Office.Interop.Excel;
using System.IO;


namespace WindowsFormsApp3
{

    public partial class Form1 : Form
    {
        string connectString = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=C:\\WindowsFormsApp3\\WindowsFormsApp3\\projectNo1.mdf;Integrated Security=True;Connect Timeout=30";
        DataTable table = new DataTable("tbl");
        Dbclass cl = new Dbclass();
        TWP cl1 = new TWP();
        EXCELLDATAREADER cl2 = new EXCELLDATAREADER();
        public Form1()
        {
            InitializeComponent();
            textBox4.Text = "Addess";
            textBox4.ForeColor = Color.Silver;
            textBox5.Text = "Name";
            textBox5.ForeColor = Color.Silver;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadData();
            cl1.CreateProxySrvr();
        }
        public void ItemsRedirect()
        {
            string[] ItemsRedirect = new string[dataGridView1.RowCount - 1];
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                ItemsRedirect[i] = dataGridView1.Rows[i].Cells[1].Value.ToString();
            }
            cl1.items = ItemsRedirect;
        }
        private void LoadData()
        {
            SqlConnection myConnection = new SqlConnection(connectString);

            string query = "SELECT * FROM [Table] ORDER BY Id";

            using (SqlConnection connection = new SqlConnection(connectString))
            {
                connection.Open();
                SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                DataSet ds = new DataSet();
                adapter.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                
            }
            ItemsRedirect();
            Rediritems();
        }
        private void Rediritems()
        {
            string redir = Properties.Settings.Default.Save_link;
            cl1.redirect = redir;
        }
        private void LoadData(string query)
        {
            SqlConnection myConnection = new SqlConnection(connectString);


            using (SqlConnection connection = new SqlConnection(connectString))
            {
                connection.Open();
                SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                DataSet ds = new DataSet();
                adapter.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];

            }
        }

        private void button5_Click(object sender, EventArgs e) //add button
        {
            if (textBox4.TextLength == 0 && textBox5.TextLength == 0)
            {

                textBox4.ForeColor = Color.Red;
                textBox4.Text = "Enter the address!";
                textBox5.ForeColor = Color.Red;
                textBox5.Text = "Enter the name!";

            }
            else if (textBox4.ForeColor == Color.Silver && textBox5.ForeColor == Color.Silver)
            {
                textBox5.ForeColor = Color.Red;
                textBox5.Text = "Enter the name!";
                textBox4.ForeColor = Color.Red;
                textBox4.Text = "Enter the address!";
            }
            else if (textBox5.TextLength == 0 || textBox5.ForeColor == Color.Silver || textBox5.ForeColor == Color.Red)
            {
                textBox5.ForeColor = Color.Red;
                textBox5.Text = "Enter the name!";
            }
            else if (textBox4.TextLength == 0 || textBox4.ForeColor == Color.Silver || textBox4.ForeColor == Color.Red)
            {
                textBox4.ForeColor = Color.Red;
                textBox4.Text = "Enter the address!";
            }
            else if (textBox4.TextLength > 0 && textBox5.TextLength > 0)
            {

                cl.connect2Db(string.Format("INSERT INTO [TABLE] (Address, Name) VALUES (N'{0}',N'{1}')", textBox4.Text, textBox5.Text));
                LoadData();
                textBox4.Text = "";
                textBox5.Text = "";
            }

        } 
        private void copyAlltoClipboard()
        {
        dataGridView1.SelectAll();
        DataObject dataObj = dataGridView1.GetClipboardContent();
        if (dataObj != null)
            Clipboard.SetDataObject(dataObj);
        }
        private void button3_Click(object sender, EventArgs e) //                          export button
        {
            SaveFileDialog sfd = new SaveFileDialog();
            string filePath = "";
            sfd.Filter = "CSV (*.CSV)|*.CSV";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                filePath = sfd.FileName;
                try
                {
                    SaveDataGridViewToCSV(filePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
                }
            }

        }

        public void button4_Click(object sender, EventArgs e) //                            import button
        {
            OpenFileDialog opf = new OpenFileDialog();
            string filePath = "";
            opf.Filter = "CSV (*.CSV)|*.CSV";
            if (opf.ShowDialog() == DialogResult.OK)
            {
                filePath = opf.FileName;
                try
                {
                    var dt = EXCELLDATAREADER.readCSV(opf.FileName);
                    //dataGridView1.DataSource = dt;
                    UpdateData(dt);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
                }
                //var dt = readCSV();
                //dataGridView1.DataSource = dt.Tables[0];
            }
        }
        void SaveDataGridViewToCSV(string filename)
        {
            // Choose whether to write header. Use EnableWithoutHeaderText instead to omit header.
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            // Select all the cells
            dataGridView1.SelectAll();
            // Copy selected cells to DataObject
            DataObject dataObject = dataGridView1.GetClipboardContent();
            // Get the text of the DataObject, and serialize it to a file
            File.WriteAllText(filename, dataObject.GetText(TextDataFormat.CommaSeparatedValue));
        }
        private void  UpdateData(DataTable dt)
        {
            cl.connect2Db("DELETE  FROM [Table] WHERE Id !='0'");
           // cl.connect2Db ("DBCC CHECKIDENT (projectNo1.dbo.Table, RESEED, 0) WITH TABLERESULTS ");
           
            
            for(int i =0; i< dt.Rows.Count; i++)
            {
                cl.connect2Db(string.Format("INSERT INTO [TABLE] (Address, Name) VALUES (N'{0}',N'{1}')", dt.Rows[i].ItemArray[1], dt.Rows[i].ItemArray[2]));
            }
            
            LoadData();
        }
      
        private void button6_Click(object sender, EventArgs e) //delete button
        {
            try
            {
                if (dataGridView1.RowCount >1)
                {
                    cl.connect2Db(string.Format("DELETE [Table] WHERE Id = {0}", Convert.ToInt32(this.dataGridView1.CurrentRow.Cells[0].Value)));

                    LoadData();
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
            }

        }
        

        private void button1_Click(object sender, EventArgs e) // search button (id - textBox2; address - textBox1; name - textBox3)
        {
            string filterStr = " ";

            if (textBox3.TextLength > 0 && textBox1.TextLength > 0 && textBox2.TextLength > 0)
            {
                 filterStr = string.Format("SELECT * FROM [Table] WHERE Name LIKE N'{0}%' AND Address LIKE N'{1}%' AND Id LIKE '{2}%'", textBox3.Text, textBox1.Text, textBox2.Text);
            }
            else if (textBox3.TextLength > 0 && textBox1.TextLength > 0)
            {
                 filterStr = string.Format(" SELECT * FROM [Table] WHERE Name LIKE N'{0}%' AND Address LIKE N'{1}%'", textBox3.Text, textBox1.Text);
            }
            else if (textBox2.TextLength > 0 && textBox1.TextLength > 0)
            {
                 filterStr = string.Format("SELECT * FROM [Table] WHERE Id LIKE '{0}%' AND Address LIKE N'{1}%'", textBox2.Text, textBox1.Text);
            }
            else if (textBox3.TextLength > 0 && textBox2.TextLength > 0)
            {
                 filterStr = string.Format("SELECT * FROM [Table] WHERE Name LIKE N'{0}%' AND Id LIKE '{1}%' ", textBox3.Text, textBox2.Text);
            }
            else if (textBox3.TextLength > 0)
            {
                 filterStr = string.Format("SELECT * FROM [Table] WHERE Name LIKE N'{0}%'", textBox3.Text);
            }
            else if (textBox2.TextLength > 0)
            {
                 filterStr = string.Format("SELECT * FROM [Table] WHERE Id LIKE '{0}%'", textBox2.Text);
            }
            else if (textBox1.TextLength > 0)
            {
                 filterStr = string.Format("SELECT * FROM [Table] WHERE Address LIKE N'{0}%'", textBox1.Text);
            }
            LoadData(filterStr);
        } 

        private void button2_Click(object sender, EventArgs e) //                show all button
        {
            LoadData();
        }

        private void textBox4_Enter(object sender, EventArgs e)
        {
                textBox4.Text = null;
                textBox4.ForeColor = Color.Black;
        }

        private void textBox5_Enter(object sender, EventArgs e)
        {
                textBox5.Text = null;
                textBox5.ForeColor = Color.Black;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            cl1.StopProxySrvr();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Redirect_btn_Click(object sender, EventArgs e)
        {
            Form2 newForm = new Form2();
            newForm.Owner = this;
            newForm.Show();
        }
    }

  
}
