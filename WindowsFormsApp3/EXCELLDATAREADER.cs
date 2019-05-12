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
            opf.Filter = "Excel (*.XLS)|*.XLS| Excel (*.XLSX)|*.XLSX";
            if (opf.ShowDialog() == DialogResult.OK)
            {
                filePath = opf.FileName;
                try
                {
                    using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            do
                            {
                                while (reader.Read())
                                {
                                }
                            } while (reader.NextResult());

                            var result = reader.AsDataSet();
                            return result;
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
            

    }
}
