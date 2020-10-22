using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Net.NetworkInformation;
using FocasLibrary;
using System.Management;
using System.IO;

namespace CNCId_GUI
{
    public partial class ReadCNCIdGUI : Form
    {       
        string CNCIDs = string.Empty;
        public ReadCNCIdGUI()
        {
            InitializeComponent();
        }

        public void ReadCNCID(string ipAddress, ushort portNo)
        {
            if (String.IsNullOrEmpty(txtIpAddress.Text.Trim()))
            {
                MessageBox.Show("Please enter IP address");
                return;
            }
            
            this.Cursor = Cursors.WaitCursor;
            ushort focasLibHandleMain = ushort.MinValue;
            Ping ping = new Ping();
            PingReply reply = ping.Send(ipAddress, 4000);
           
            if (reply.Status == IPStatus.Success)
            {
                short ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 4, out focasLibHandleMain);
                 if (ret == 0)
                 {
                     //txtCncId.Text = ReadCNCId(focasLibHandleMain);
                     CNCIDs = txtIpAddress.Text.PadRight(15) + "[ " + ReadCNCId(focasLibHandleMain) + " ]";
                      if (!lstCNCID.Items.Contains(CNCIDs))
                      {
                          lstCNCID.Items.Add(CNCIDs);
                      } 
                 }
                else
                {
                    MessageBox.Show("Not able to connect to CNC machine, Please check IP Address, Port No (8193). Return value from function is " + ret + " .");
                }
                ret = FocasLib.cnc_freelibhndl(focasLibHandleMain);
            }
            else
            {
                MessageBox.Show("Not able to Ping the CNC Machine. Please check the ip address and try again. Ping Status = " + reply.Status);
            }
            this.Cursor = Cursors.Default;
        }

        public string ReadCNCId(ushort handle)
        {
            string cncIDstr = "";
            uint[] cncid = new uint[4];
            int ret = FocasLib.cnc_rdcncid(handle, cncid);
            if (ret == 0)
            {
                foreach (var item in cncid)
                {
                    cncIDstr = cncIDstr + item.ToString("X") + "-";
                }               
            }
           cncIDstr = cncIDstr.Trim('-');

           return (cncIDstr);// Base64Encode(cncIDstr);
        }

        private void btnReadCncId_Click(object sender, EventArgs e)
        {
            ReadCNCID(txtIpAddress.Text.Trim(), 8193);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            ushort focasLibHandleMain = ushort.MinValue;
            Ping ping = new Ping();
            PingReply reply = ping.Send(txtIpAddress.Text.Trim(), 4000);
            if (reply.Status == IPStatus.Success)
            {
                short ret = FocasLib.cnc_allclibhndl3(txtIpAddress.Text.Trim(), 8193, 4, out focasLibHandleMain);
                if (ret == 0)
                {
                    ret = FocasLib.cnc_freelibhndl(focasLibHandleMain);
                    MessageBox.Show("Successfully connected to CNC machine.");
                }
                else
                {
                    MessageBox.Show("Not able to connect to CNC machine, Please check IP Address, Port No(8193). Return value from function is " + ret + " .");
                }
            }
            else
            {
                MessageBox.Show("Not able to Ping the CNC Machine. Please check the ip address and try again. Ping Status = " + reply.Status);
            }
            this.Cursor = Cursors.Default;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {

            if (lstCNCID.Items.Count == 0)
            {
                return;
            }
          
            string name = string.Empty;
            SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
            SaveFileDialog1.Title = "Save File...";

            SaveFileDialog1.CheckFileExists = false;
            SaveFileDialog1.CheckPathExists = false;

            SaveFileDialog1.DefaultExt = "txt";
            SaveFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            SaveFileDialog1.FilterIndex = 2;
            SaveFileDialog1.RestoreDirectory = true;
            SaveFileDialog1.OverwritePrompt = false;
            SaveFileDialog1.FileName = "CNCData.txt";

            if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
            {
             name = SaveFileDialog1.FileName;
            }

            if (String.IsNullOrEmpty(name))
            {
                return;
            }

            if (!File.Exists(name))
            {
                using (StreamWriter SaveFile = File.CreateText(name))
                {
                    foreach (var item in lstCNCID.Items)
                    {
                        SaveFile.WriteLine(item);
                    }
                    SaveFile.Close();
                }
            }            
            else 
            {
                using (StreamWriter SaveFile = new StreamWriter(name, false))
                {
                    foreach (var item in lstCNCID.Items)
                    {
                        SaveFile.WriteLine(item);
                    }
                    SaveFile.Close();
                }
            }

           // MessageBox.Show("CNC Data saved Successfully!");
            MessageBox.Show("CNC Data saved Successfully!!!","Message",MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        private void lstCNCID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                for (int i = lstCNCID.SelectedIndices.Count - 1; i >= 0; i--)
                {
                    lstCNCID.Items.RemoveAt(lstCNCID.SelectedIndices[i]);
                }
            }

        }

        public static string Base64Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }

        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            lstCNCID.Items.Clear();

        }

    }
}
