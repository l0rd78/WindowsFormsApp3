using System;
using System.Windows.Forms;
namespace WindowsFormsApp3
{
    public partial class Form2 : Form
    {
        
       
        public Form2()
        {
            InitializeComponent();
            Form1 frm = (Form1)this.Owner;
            
            name.Text = Properties.Settings.Default.Save_name;
            link.Text = Properties.Settings.Default.Save_link;
        }
        private void name_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save_name = name.Text;
            Properties.Settings.Default.Save();
        }

        private void link_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save_link = link.Text;
            Properties.Settings.Default.Save();
        }
        public void saveBTTN_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void Form2_Load(object sender, EventArgs e)
        {
        }
    }
}
