using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace WindowsFormsApp3
{
    class Dbclass
    {
        string connectString = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=C:\\WindowsFormsApp3\\WindowsFormsApp3\\projectNo1.mdf;Integrated Security=True;Connect Timeout=30";
        public void connect2Db(string query)
        {

            try
            {
                string sqlExpression = string.Format("{0}", query);
                using (SqlConnection connection = new SqlConnection(connectString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    command.ExecuteNonQuery();
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
            }
        }
        public Dbclass()
        {

        }

    }
}
