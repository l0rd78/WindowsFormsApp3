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


namespace WindowsFormsApp3
{
    
    public partial class Form1 : Form
    {
        string connectString = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=C:\\Users\\qwerty\\Documents\\projectNo1.mdf;Integrated Security=True;Connect Timeout=30";
        DataTable table = new DataTable("tbl");
        
        public Form1()
        {
            InitializeComponent();
            textBox4.Text = "Addess";
            textBox4.ForeColor = Color.Silver;  
            textBox5.Text = "Name";
            textBox5.ForeColor = Color.Silver;
            /*textBox2.Text = "Id?";
            textBox2.ForeColor = Color.Silver;
            textBox1.Text = "Addres?";
            textBox1.ForeColor = Color.Silver;
            textBox3.Text = "Name?";
            textBox3.ForeColor = Color.Silver;*/


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
            }
        }
        private void blockItem()
        {
            /*string writePath = @"C:\Windows\System32\drivers\etc\hosts";
            using (FileStream fstream = new FileStream(writePath, FileMode.OpenOrCreate))
                
            {
                using (StreamWriter strWriter = new StreamWriter(fstream))
                {
                    for (int i = 0; i < dataGridView1.RowCount - 1; i++) //-1 обязательно, иначе попытается считать пустую строку
                    {
                        strWriter.WriteLine("173.194.122.168" + " " + dataGridView1.Rows[i].Cells[1].Value.ToString());
                        strWriter.WriteLine("173.194.122.168" + " " + "www." + dataGridView1.Rows[i].Cells[1].Value.ToString());
                    }
             }*/



        }
        public string GetBrowserUrl()
        {
            /* if (Browser == BrowserType.Chrome)
             {
                 //"Chrome_WidgetWin_1"

                 Process[] procsChrome = Process.GetProcessesByName("chrome");
                 foreach (Process chrome in procsChrome)
                 {
                     // the chrome process must have a window
                     if (chrome.MainWindowHandle == IntPtr.Zero)
                     {
                         continue;
                     }
                     //AutomationElement elm = AutomationElement.RootElement.FindFirst(TreeScope.Children,
                     //         new PropertyCondition(AutomationElement.ClassNameProperty, "Chrome_WidgetWin_1"));
                     // find the automation element
                     AutomationElement elm = AutomationElement.FromHandle(chrome.MainWindowHandle);

                     // manually walk through the tree, searching using TreeScope.Descendants is too slow (even if it more reliable)
                     AutomationElement elmUrlBar = null;
                     try
                     {
                         // walking path found using inspect.exe (Windows SDK) for Chrome 29.0.1547.76 m (currently the latest stable)
                         var elm1 = elm.FindFirst(System.Windows.Automation.TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Google Chrome"));
                         var elm2 = TreeWalker.ControlViewWalker.GetLastChild(elm1); // I don't know a Condition for this for finding :(
                         var elm3 = elm2.FindFirst(System.Windows.Automation.TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, ""));
                         var elm4 = elm3.FindFirst(System.Windows.Automation.TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ToolBar));
                         elmUrlBar = elm4.FindFirst(System.Windows.Automation.TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Address and search bar"));
                     }
                     catch
                     {
                         // Chrome has probably changed something, and above walking needs to be modified. :(
                         // put an assertion here or something to make sure you don't miss it
                         continue;
                     }

                     // make sure it valid
                     if (elmUrlBar == null)
                     {
                         // it not..
                         continue;
                     }

                     // elmUrlBar is now the URL bar element. we have to make sure that it out of keyboard focus if we want to get a valid URL
                     if ((bool)elmUrlBar.GetCurrentPropertyValue(AutomationElement.HasKeyboardFocusProperty))
                     {
                         continue;
                     }

                     // there might not be a valid pattern to use, so we have to make sure we have one
                     AutomationPattern[] patterns = elmUrlBar.GetSupportedPatterns();
                     if (patterns.Length == 1)
                     {
                         string ret = "";
                         try
                         {
                             ret = ((ValuePattern)elmUrlBar.GetCurrentPattern(patterns[0])).Current.Value;
                         }
                         catch { }
                         if (ret != "")
                         {
                             // must match a domain name (and possibly "https://" in front)
                             if (Regex.IsMatch(ret, @"^(https:\/\/)?[a-zA-Z0-9\-\.]+(\.[a-zA-Z]{2,4}).*$"))
                             {
                                 // prepend http:// to the url, because Chrome hides it if it not SSL
                                 if (!ret.StartsWith("http"))
                                 {
                                     ret = "http://" + ret;
                                 }
                                 return ret;

                             }
                         }
                         continue;
                     }
                 }

             //}*/
            return "";
            
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

        private void textBox5_TextChanged(object sender, EventArgs e) //name add
        {

        }

        private void button5_Click(object sender, EventArgs e) //add button
        {
            if ( textBox4.TextLength == 0 && textBox5.TextLength == 0)
            {
                textBox4.Text = "Enter the address!";
                textBox5.Text = "Enter the name!";
            }   
            else if (textBox5.TextLength == 0 || textBox5.ForeColor == Color.Silver)
                textBox5.Text = "Enter the name!";
            else if (textBox4.TextLength == 0 || textBox4.ForeColor == Color.Silver)
                textBox4.Text = "Enter the address!";
            else if (textBox4.TextLength > 0 && textBox5.TextLength > 0)
            {
                SqlConnection myConnection = new SqlConnection(connectString);
                myConnection.Open();
                string addItem = string.Format ("INSERT INTO [TABLE] (Address, Name) VALUES (N'{0}',N'{1}')", textBox4.Text, textBox5.Text);
                SqlCommand command = new SqlCommand(addItem, myConnection);
                command.CommandText = addItem;
                command.ExecuteNonQuery();
                myConnection.Close();
                LoadData();
            }
        }

        
        private void button3_Click(object sender, EventArgs e) //                          export button
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            sfd.FileName = "Inventory_Adjustment_Export.xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                // Copy DataGridView results to clipboard
                copyAlltoClipboard();

                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlexcel = new Excel.Application();

                xlexcel.DisplayAlerts = false; // Without this you will get two confirm overwrite prompts
                Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                // Format column D as text before pasting results, this was required for my data
                Excel.Range rng = xlWorkSheet.get_Range("D:D").Cells;
                rng.NumberFormat = "@";

                // Paste clipboard results to worksheet range
                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                // For some reason column A is always blank in the worksheet. ¯\_(ツ)_/¯
                // Delete blank column A and select cell A1
                Excel.Range delRng = xlWorkSheet.get_Range("A:A").Cells;
                delRng.Delete(Type.Missing);
                xlWorkSheet.get_Range("A1").Select();

                // Save the excel file under the captured location from the SaveFileDialog
                xlWorkBook.SaveAs(sfd.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlexcel.DisplayAlerts = true;
                xlWorkBook.Close(true, misValue, misValue);
                xlexcel.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlexcel);

                // Clear Clipboard and DataGridView selection
                Clipboard.Clear();
                dataGridView1.ClearSelection();

                // Open the newly saved excel file
                if (System.IO.File.Exists(sfd.FileName))
                    System.Diagnostics.Process.Start(sfd.FileName);
                 }
            }
        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }


        private void button4_Click(object sender, EventArgs e) //                            import button
        {
            dataGridView1.DataSource = null;
            OpenFileDialog opf = new OpenFileDialog();
            
            opf.Filter = "Excel (*.XLS)|*.XLS";
            opf.ShowDialog();
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

                    }
                    cn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
            }



            /*OleDbConnection con;
            OleDbDataAdapter adapter;
            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter = "Excel (*.XLS)|*.XLS";
            opf.ShowDialog();
            string filename = opf.FileName;
            string conString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Extended Properties=Excel 12.0;");
            
                con = new OleDbConnection(conString);
            DataTable schemaTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];
            string sql = String.Format("SELECT * FROM [{0}]", sheet1);

            //ADAPTER
            adapter = new OleDbDataAdapter(sql, con);

            DataSet ds = new DataSet();

            adapter.Fill(ds);

            dataGridView1.DataSource = ds.Tables[0];

            //CLOSE CON
            con.Close();*/
        }
      




       

        private void button6_Click(object sender, EventArgs e) //delete button
        {
            SqlConnection myConnection = new SqlConnection(connectString);
            myConnection.Open();
            string delItem = string.Format("DELETE [Table] WHERE Id = {0}", Convert.ToInt32(this.dataGridView1.CurrentRow.Cells[0].Value));
            SqlCommand command = new SqlCommand(delItem, myConnection);
            command.CommandText = delItem;
            command.ExecuteNonQuery();
            myConnection.Close();
            LoadData();
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

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

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
