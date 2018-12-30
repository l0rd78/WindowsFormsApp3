using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;
using xNet;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Jacob.UIAutomation;
using UIAutomationClient;
using System.Windows.Automation;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using EasyXLS;



namespace WindowsFormsApp3
{
    
    public partial class Form1 : Form
    {
        string connectString = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=C:\\Users\\qwerty\\Documents\\projectNo1.mdf;Integrated Security=True;Connect Timeout=30";
        DataTable table = new DataTable("tbl");
        bool b = false;

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
            blockItem();
            
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
                flushdns();
            }
        }
        public void flushdns()
        {
            Process cmd = new Process();
            cmd.StartInfo.FileName = "cmd.exe";
            cmd.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            cmd.StartInfo.Arguments = "/C ipconfig /flushdns";
            cmd.Start();
        }
        private void blockItem()
        {
            string writePath = @"C:\Windows\System32\drivers\etc\hosts";
            using (FileStream fstream = new FileStream(writePath, FileMode.OpenOrCreate))

            {
                using (StreamWriter strWriter = new StreamWriter(fstream))
                {
                    for (int i = 0; i < dataGridView1.RowCount - 1; i++) //-1 обязательно, иначе попытается считать пустую строку
                    {
                        strWriter.WriteLine("127.0.0.1" + " " + dataGridView1.Rows[i].Cells[1].Value.ToString());
                        strWriter.WriteLine("127.0.0.1" + " " + "www." + dataGridView1.Rows[i].Cells[1].Value.ToString());
                    }
                }
            }
            flushdns();
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
                if(b == false)
                {
                    SqlConnection myConnection = new SqlConnection(connectString);
                    myConnection.Open();
                    string addItem = string.Format("INSERT INTO [TABLE] (Address, Name) VALUES (N'{0}',N'{1}')", textBox4.Text, textBox5.Text);
                    SqlCommand command = new SqlCommand(addItem, myConnection);
                    command.CommandText = addItem;
                    command.ExecuteNonQuery();
                    myConnection.Close();
                    LoadData();
                }
                else
                {


                }
                
            }
        }

        
        private void button3_Click(object sender, EventArgs e) //                          export button
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            sfd.FileName = "Inventory_Adjustment_Export.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {

                // Create an instance of the class that exports Excel files, having one sheet
                ExcelDocument workbook = new ExcelDocument(1);

                // Set sheet name
                ExcelWorksheet xlsWorksheet = (ExcelWorksheet)workbook.easy_getSheetAt(0);
                xlsWorksheet.setSheetName("DataGridView");

                // Get the sheet table that stores the data
                ExcelTable xlsTable = xlsWorksheet.easy_getExcelTable();
                int tableRow = 0;

                // Export DataGridView header if the header is visible
                if (dataGridView1.ColumnHeadersVisible)
                {

                    // Add data in cells for header
                    for (int column = 0; column < dataGridView1.Columns.Count; column++)
                    {
                        xlsTable.easy_getCell(tableRow, column).setValue(
                                              dataGridView1.Columns[column].HeaderText);
                    }
                    tableRow++;
                }

                // Add data in cells
                for (int row = 0; row < dataGridView1.Rows.Count - 1; row++)
                {
                    for (int column = 0; column < dataGridView1.Columns.Count; column++)
                    {
                        xlsTable.easy_getCell(tableRow, column).setValue(
                                              dataGridView1.Rows[row].Cells[column].Value.ToString());
                    }
                    tableRow++;
                }

                // Export Excel file
                workbook.easy_WriteXLSFile(sfd.FileName);
            }
        }

        private void button4_Click(object sender, EventArgs e) //                            import button
        {
            OpenFileDialog opf = new OpenFileDialog();
            
            opf.Filter = "Excel (*.XLS)|*.XLS";
            if (opf.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string filename = opf.FileName;
                    string ConStr = String.Format("Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Extended Properties='Excel 8.0;HDR=YES;';", filename);
                    using (OleDbConnection cn = new OleDbConnection(ConStr))
                    {
                        cn.Open();
                        DataTable schemaTable = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                        string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];

                        //string sheet1 = @"Лист1";
                        string select = String.Format("SELECT * FROM [{0}]", sheet1);
                        //string select = "SELECT * FROM [" + sheet1 + "$]";

                        using (OleDbDataAdapter ad = new OleDbDataAdapter(select, cn))
                        {
                            DataTable dt = new DataTable();
                            ad.Fill(dt);
                            dataGridView1.DataSource = dt;
                            UpdateData(dt);
                            b = true;

                        }
                        cn.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
                }
            }
            
        }
        private void  UpdateData(DataTable dt)
        {
            string sqlExpression = "DELETE  FROM [Table] WHERE Id !='0'";
            using (SqlConnection connection = new SqlConnection(connectString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand(sqlExpression, connection);
                command.ExecuteNonQuery();
            }
            for(int i =0; i< dt.Rows.Count; i++)
            {
                string sqlExpression2 = string.Format("INSERT INTO [Table](Id, Address, Name) VALUES({0}, {1}, {2})", dt.Rows[i].ItemArray[0], dt.Rows[i].ItemArray[1], dt.Rows[i].ItemArray[2]);
                using (SqlConnection connection = new SqlConnection(connectString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    command.ExecuteNonQuery();
                }
            }
            
        }
      
        private void button6_Click(object sender, EventArgs e) //delete button
        {
            if (b == false)
            {
                SqlConnection myConnection = new SqlConnection(connectString);
                myConnection.Open();
                string delItem = string.Format("DELETE [Table] WHERE Id = {0}", Convert.ToInt32(this.dataGridView1.CurrentRow.Cells[0].Value));
                SqlCommand command = new SqlCommand(delItem, myConnection);
                command.CommandText = delItem;
                command.ExecuteNonQuery();
                myConnection.Close();
                LoadData();
            }else
            {

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
    }
}
