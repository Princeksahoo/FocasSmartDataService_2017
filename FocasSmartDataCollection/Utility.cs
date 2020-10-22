using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.Cache;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using MachineConnectLicenseDTO;

namespace FocasSmartDataCollection
{
    public static class Utility
    {
        [DllImport("iphlpapi.dll", ExactSpelling = true)]
        public static extern int SendARP(int DestIP, int SrcIP, byte[] pMacAddr, ref uint PhyAddrLen);

        public static string SafeFileName(string name)
        {
            StringBuilder str = new StringBuilder(name);
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                str = str.Replace(c, '_');
            }
            return str.ToString();
        }

        private static LicenseTerms ValidateLicense(LicenseDTO license)
        {
            RSACryptoServiceProvider rsa = GetPublicKeyFromAssemblyAtClient(Assembly.GetExecutingAssembly());

            byte[] dataToVerify = Convert.FromBase64String(license.LicenseTerms);
            byte[] signature = Convert.FromBase64String(license.Signature);

            // verify that the license-terms match the signature data
            if (rsa.VerifyData(dataToVerify, new SHA1CryptoServiceProvider(), signature))
                return LicenseTerms.FromString(license.LicenseTerms);
            else
                throw new Exception("license file is tempered.");
        }

        public static bool validateLicenseFile(string licFile, ref LicenseTerms lic)
        {
            LicenseDTO license = null;
            if (File.Exists(licFile))
            {
                license = LicenseDTO.FromString(File.ReadAllText(licFile));
            }
            else
            {
                throw new FileNotFoundException("Lic file not found");
            }

            if (license != null)
            {
                lic = ValidateLicense(license);
                return true;
            }
            else
            {
                //TODO : "Invalid License File Not Supplied!"
                return false;
            }
            return false;
        }

        private static RSACryptoServiceProvider GetPublicKeyFromAssemblyAtClient(Assembly assembly)
        {
            // Extract public key - note that public key stored in assembly has an extra 3 headers
            // (12 bytes) at the front that are not part of a CAPI public key blob, so they must be removed
            byte[] rawPublicKeyData = assembly.GetName().GetPublicKey();

            int extraHeadersLen = 12;
            int bytesToRead = rawPublicKeyData.Length - extraHeadersLen;
            byte[] publicKeyData = new byte[bytesToRead];
            Buffer.BlockCopy(rawPublicKeyData, extraHeadersLen, publicKeyData, 0, bytesToRead);
            RSACryptoServiceProvider publicKey = new RSACryptoServiceProvider();
            publicKey.ImportCspBlob(publicKeyData);
            return publicKey;
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


        public static bool ValidateHardDiskSerialnumber(string hardDiskSerialnumber)
        {
            bool isHardDiskSerialNumberValid = false;
            List<string> serialNumberToValidate = new List<string>
                                                  {"202020202020202020202020325a48414c423947","E3G24262C41WRD",
                                                      "2020202057202d44585731393541363031373336","W627TYTK","E3G24262C41WRD"
                                                  };
            serialNumberToValidate.Add(hardDiskSerialnumber);
            List<string> serialNumbers = new List<string>();
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive where InterfaceType = 'IDE'");

                foreach (ManagementObject wmi_HD in searcher.Get())
                {
                    if (!string.IsNullOrWhiteSpace(wmi_HD["SerialNumber"].ToString().Trim()))
                    {                       
                        serialNumbers.Add(wmi_HD["SerialNumber"].ToString().Trim());
                    }
                }
            }
            catch (Exception exx)
            {
                Logger.WriteErrorLog(exx.ToString());
            }
            foreach (string item in serialNumbers)
            {
                if (serialNumberToValidate.Contains(item))
                {
                    isHardDiskSerialNumberValid = true;
                    break;
                }
            }
            return isHardDiskSerialNumberValid;
        }
                
        public static string GetMacAddress(string ipAddressStr)
        {
            IPAddress iPAddress;            
            if (!IPAddress.TryParse(ipAddressStr, out iPAddress))
            {
                return string.Empty;            
            }

            if (!IsHostAccessible(iPAddress)) return string.Empty;

            //IPHostEntry hostEntry = Dns.GetHostEntry(ipAddressStr);
            //if (hostEntry.AddressList.Length == 0) return string.Empty;

            byte[] macAddr = new byte[6];
            uint macAddrLen = (uint)macAddr.Length;
            if (SendARP((int)iPAddress.Address, 0, macAddr, ref macAddrLen) != 0)
            {
                Logger.WriteErrorLog("ARP command failed");
                return string.Empty;
            }

            string[] str = new string[(int)macAddrLen];
            for (int i = 0; i < macAddrLen; i++)
            {
                str[i] = macAddr[i].ToString("x2");
            }
            return string.Join(":", str).ToUpper();
        }

        private static bool IsHostAccessible(IPAddress IPAddress)
        {
            Ping ping = new Ping();
            PingReply reply = ping.Send(IPAddress, 4000);
            return reply.Status == IPStatus.Success;
        }

        public static DateTime GetNetworkTime()
        {
            DateTime localDateTime = DateTime.Now;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://nist.time.gov/actualtime.cgi?lzbc=siqm9b");
            request.Method = "GET";
            request.Accept = "text/html, application/xhtml+xml, */*";
            request.UserAgent = "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; Trident/6.0)";
            request.ContentType = "application/x-www-form-urlencoded";
            request.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore); //No caching
            try
            {
                request.Timeout = 40 * 1000;
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    StreamReader stream = new StreamReader(response.GetResponseStream());
                    string html = stream.ReadToEnd();//<timestamp time=\"1395772696469995\" delay=\"1395772696469995\"/>
                    string time = System.Text.RegularExpressions.Regex.Match(html, @"(?<=\btime="")[^""]*").Value;
                    double milliseconds = Convert.ToInt64(time) / 1000.0;
                    localDateTime = new DateTime(1970, 1, 1).AddMilliseconds(milliseconds).ToLocalTime();
                }
            }
            catch (Exception exx)
            {

            }

            return localDateTime;
        }

        public static string get_actual_date(long number)
        {
            //TODO
            string strDate = string.Empty;
            int length = number.ToString().Length;
            if (length == 6)
            {
                strDate = number.ToString("####-#-#");
            }
            else if (length == 7)
            {
                strDate = DateTime.Now.Month >= 10 ? number.ToString("####-##-#") : number.ToString("####-#-##");
            }
            else if (length == 8)
            {
                strDate = number.ToString("####-##-##");
            }
            return strDate;
        }
        public static string get_actual_time(long number)
        {
            string strTime = number.ToString("0#:0#:0#");
            return strTime;
            //TODO            
        }

        public static bool ValidateServiceTagName()
        {
            List<string> allowedServiceTags = new List<string> { "J4F0HZ1", "5Y66D12", "2C5N7X1", "67B6D12", "7B8B6BS", "2XGXC12", "9695042", 
                "6G46D12", "INA8340L1T", "INA8340L11", "CP66D12", "9XDXC12", "INA8340L1F", "JMQB6BS", "BXGXC12", "BP66D12", "8XGXC12","F6FMWQ1","R8F4GKY","J695042","7GXPYB2" };
            return allowedServiceTags.Contains(GetServiceTag().ToUpper());
        }

        static private string GetComputerName()
        {
            return Environment.MachineName;
        }

        public static bool ValidateDomainName()
        {
            string domainName = "acemicromatic.com";
            return System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName.Equals(domainName, StringComparison.OrdinalIgnoreCase);
        }

        public static bool ValidateMACAddress()
        {
            List<string> allowedServiceTags = new List<string> {"A0-88-B4-5B-FA-50","00-21-CC-5F-58-CA", "D4-BE-D9-C1-8C-36", "34-C3-D2-E3-97-0D", "36-C3-D2-E3-97-0D", "34-C3-D2-94-22-2D", "36-C3-D2-94-22-2D", 
                "34-C3-D2-94-30-B8", "36-C3-D2-94-30-B8","14-18-77-BE-B9-55","F8-CA-B8-61-CF-F8","00-22-64-6D-96-E9","1C-39-47-33-ED-D5","14-18-77-BE-B9-56","28-D2-44-A6-6F-52",
                "18-03-73-9B-9E-82","00-1C-BF-A8-F6-85","18-03-73-9B-9E-CB","18-03-73-9B-9E-82","3C-97-0E-C1-1E-B0","18-03-73-95-C9-87","34-23-87-6A-63-BC","34-C3-D2-E3-97-0D","14-18-77-BB-2A-EE","18-03-73-9B-9E-CB"};
            List<string> MACAddress = GetMACAddress();
            return allowedServiceTags.Intersect(MACAddress).Count() > 0;
        }

        static private List<string> GetMACAddress()
        {
            List<string> MACList = new List<string>();

            foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (nic.NetworkInterfaceType == NetworkInterfaceType.Wireless80211 || nic.NetworkInterfaceType == NetworkInterfaceType.Ethernet)
                {
                    string MAC = nic.GetPhysicalAddress().ToString();
                    //  Insert hyphens
                    MACList.Add(String.Join("-", Regex.Matches(MAC, @"\w{2}").Cast<Match>()).ToUpper());
                }
            }
            return MACList;
        }

        static private string GetModel()
        {
            string model = "";

            ManagementClass wmi = new ManagementClass("Win32_ComputerSystem");
            foreach (ManagementObject computer in wmi.GetInstances())
            {
                model = computer.Properties["Model"].Value.ToString().Trim();
            }
            return model;
        }

        static private string GetManufacturer()
        {
            string manufacturer = "";

            ManagementClass wmi = new ManagementClass("Win32_ComputerSystem");
            foreach (ManagementObject computer in wmi.GetInstances())
            {
                manufacturer = computer.Properties["Manufacturer"].Value.ToString().Trim();
            }
            return manufacturer;
        }

        private static string GetServiceTag()
        {
            string dellServiceTag = "";

            try
            {
                ManagementClass wmi = new ManagementClass("Win32_Bios");

                foreach (ManagementObject bios in wmi.GetInstances())
                {
                    dellServiceTag = bios.Properties["Serialnumber"].Value.ToString().Trim();
                }
                return dellServiceTag;
            }
            catch (Exception e)
            {
                Logger.WriteErrorLog(e.Message);
            }
            return dellServiceTag;
        }

        public static DateTime GetDatetimeFromtpmString(int date, int time)
        {
            string[] formats = new string[]
			{
				"yyyyMMdd HHmmss"
			};
            DateTime minValue = DateTime.MinValue;
            if (!DateTime.TryParseExact(date + " " + time, formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out minValue))
            {
                string strDate = Utility.get_actual_date(date);
                string strTime = Utility.get_actual_time(time);
                DateTime.TryParse(strDate + " " + strTime, out minValue);
            }
            return minValue;
        }

        public static bool CheckPingStatus(string ipAddress)
        {
            bool pingStatus = false;
            IPStatus status = IPStatus.Unknown;
            Ping ping = null;
            try
            {
                ping = new Ping();
                PingReply pingReply = ping.Send(ipAddress, 10000);
                status = pingReply.Status;
                if (pingReply.Status == IPStatus.Success)
                {
                    pingStatus = true;
                }
                else
                {
                    Logger.WriteErrorLog(string.Format("Not able to ping IP Address {0} . Ping status {1} ", ipAddress, status.ToString()));
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(string.Format("Not able to ping IP Address {0} . Ping status {1} ", ipAddress, status.ToString()));
            }
            finally
            {
                if (ping != null)
                {
                    ping.Dispose();
                }
            }
            return pingStatus;
        }

    }
}
