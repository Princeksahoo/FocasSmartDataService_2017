using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using System.IO;

namespace MachineConnectLicenseDTO
{
    public partial class LicenseGUI : Form
    {
       
        public LicenseGUI()
        {
            InitializeComponent();           
            cmdLicType.SelectedIndex = 0;
            bindingSource1.DataSource = new List<CNCData>(); ;
        }

        private void btnSelectLicenseFile_Click(object sender, EventArgs e)
        {
           var result = openFileDialog.ShowDialog(this);
           if (result == System.Windows.Forms.DialogResult.OK)
           {
               txtLicFiletoRead.Text = openFileDialog.FileName;
               LicenseTerms terms = null;
               LicenseServer.validateLicenseFile(openFileDialog.FileName, ref terms);
               if (terms != null)
               {                   
                   txtCustomer.Text = terms.Customer;
                   txtPlant.Text = terms.Plant;
                   txtEmail.Text = terms.Email;
                   cmdLicType.SelectedItem = terms.LicType;
                   dtpStartDate.Text = terms.StartDate;
                   dtpExpiresAt.Text = terms.ExpiresAt;                   
                   txtSerialNo.Text = terms.ComputerSerialNo;                  
                   bindingSource1.DataSource = terms.CNCData;
               }
           }
        }      

        private void btnGenerateLic_Click(object sender, EventArgs e)
        {           
            saveFileDialog.CheckFileExists = false;
            saveFileDialog.FileName = "SmartDataServiceFocas.lic";          
            List<CNCData> cNCDataList = new List<CNCData>();
            foreach (var item in bindingSource1.DataSource as IEnumerable<CNCData>)
            {
                cNCDataList.Add(new CNCData { IsOEM = item.IsOEM, CNCdata1 = item.CNCdata1, CNCdata2 = item.CNCdata2, CNCdata3 = item.CNCdata3 });
            }
           
            LicenseTerms terms = new LicenseTerms 
            { 
                Customer = txtCustomer.Text.Trim(),
                Plant = txtPlant.Text.Trim(),
                Email = txtEmail.Text.Trim(),
                LicType = cmdLicType.SelectedItem.ToString(),
                StartDate = dtpStartDate.Value.Date.ToString("yyyy-MM-dd"),
                ExpiresAt = dtpExpiresAt.Value.Date.ToString("yyyy-MM-dd"),                
                ComputerSerialNo = txtSerialNo.Text.Trim(),               
                CNCData = cNCDataList,           
            };
            LicenseDTO licDto = LicenseServer.CreateLicense(terms);            
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                licDto.SaveLicenseStringFinal(saveFileDialog.FileName);
            }           
        }      

      
        private void cmdLicType_SelectionChangeCommitted(object sender, EventArgs e)
        {           
            dtpExpiresAt.Enabled = cmdLicType.SelectedItem.ToString() == "Trial";
            dtpStartDate.Enabled = cmdLicType.SelectedItem.ToString() == "Trial";
        }

        private void LicenseGUI_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }     
    }
}
