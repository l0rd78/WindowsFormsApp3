using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace WindowsFormsApp3
{
    class EXCELLDATAREADER
    {
       public static DataSet Reader()
        {
            OpenFileDialog opf = new OpenFileDialog();
            string filePath = "";
            opf.Filter = "CSV (*.CSV)|*.CSV";
            if (opf.ShowDialog() == DialogResult.OK)
            {
                filePath = opf.FileName;
                try
                {
                    using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {

                            var conf = new ExcelDataSetConfiguration
                            {
                                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                                {
                                    UseHeaderRow = true
                                }
                            };
                            var dataSet = reader.AsDataSet(conf);
                            var dataTable = dataSet.Tables[0];
                            return dataSet;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
                }

            }
            return new DataSet();
        }

        static public DataTable readCSV(string filePath)
        {
            var dt = new DataTable();
            // Creating the columns
            foreach (var headerLine in File.ReadLines(filePath).Take(1))
            {
                foreach (var headerItem in headerLine.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    dt.Columns.Add(headerItem.Trim());
                }
            }

            // Adding the rows
            foreach (var line in File.ReadLines(filePath).Skip(1))
            {
                dt.Rows.Add(line.Split(',').Skip(1).ToArray());
            }
            return dt;
        }


    }
}
