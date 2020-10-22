using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using MachineConnectLicenseDTO;


namespace MachineConnectLicenseDTO
{
    public partial class Login : Form
    {          
        public Login()
        {
            InitializeComponent();
        }

        private void Login_Load(object sender, EventArgs e)
        {
            txtUsername.Focus();
        }

        private void btnLogin_Click_1(object sender, EventArgs e)
        {
            try
            {        
                if (txtUsername.Text.Trim().Equals("pct") && txtpassword.Text.Equals("PCT"))
                {
                    LicenseGUI gui = new LicenseGUI();
                    gui.Show();
                    this.Hide();
                }     
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error,.!! \n" + ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }      
    }        
}
