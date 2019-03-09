using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using System.Data.OleDb;
//using EasyXLS;


namespace WindowsFormsApp3
{

    public partial class Form1 : Form
    {
        string connectString = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=C:\\Users\\qwerty\\source\\repos\\WindowsFormsApp3\\WindowsFormsApp3\\projectNo1.mdf;Integrated Security=True;Connect Timeout=30";
        DataTable table = new DataTable("tbl");
        Dbclass cl = new Dbclass();
        TWP cl1 = new TWP();
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
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            sfd.FileName = "Inventory_Adjustment_Export.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                copyAlltoClipboard();
                Microsoft.Office.Interop.Excel.Application xlexcel;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlexcel = new Excel.Application();
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);          
                /*

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
                */
            }
            
        }

        public void button4_Click(object sender, EventArgs e) //                            import button
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
                        string select = String.Format("SELECT * FROM [{0}]", sheet1);
                        using (OleDbDataAdapter ad = new OleDbDataAdapter(select, cn))
                        {
                            DataTable dt = new DataTable();
                            ad.Fill(dt);
                            UpdateData(dt);
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
    }

  
}
