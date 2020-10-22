using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Reflection;
using System.Windows.Forms;
using General.IO;
using System.Security.Cryptography;
using System.Security;

namespace MachineConnectLicenseDTO
{
    public class LicenseServer
    {
        //create the LicenseTerms object from GUI and call this function to create lic
        public static LicenseDTO CreateLicense(LicenseTerms terms)
        {   
            //Get the rsa from snk file
            RSACryptoServiceProvider rsa = Utility.GetRSAFromSnkFileAtServer(@"SmartDataServiceFocas.snk");          

            // get the byte-array of the licence terms:
            byte[] license = terms.GetLicenseData();

            // get the signature:
            byte[] signature = rsa.SignData(license, new SHA1CryptoServiceProvider());

            // now create the license object:
            return new LicenseDTO()
            {
                LicenseTerms = Convert.ToBase64String(license),
                Signature = Convert.ToBase64String(signature)
            };

            // save the license file:
            //license.Save(licenseResourceFolder + "\\" + userName + ".lic");
        }      

        internal static LicenseTerms ValidateLicense(LicenseDTO license)
        {          
            RSACryptoServiceProvider rsa = Utility.GetPublicKeyFromAssemblyAtClient(Assembly.GetExecutingAssembly());          
                       
            byte[] dataToVerify = Convert.FromBase64String(license.LicenseTerms);                     
            byte[] signature = Convert.FromBase64String(license.Signature);

            // verify that the license-terms match the signature data
            if (rsa.VerifyData(dataToVerify, new SHA1CryptoServiceProvider(), signature))
                return LicenseTerms.FromString(license.LicenseTerms);
            else
                return null;
        }

        internal static bool validateLicenseFile(string licFile, ref LicenseTerms lic)
        {
            try
            {
                LicenseDTO license = null;               
                if (File.Exists(licFile))
                {
                    license = LicenseDTO.FromString(File.ReadAllText(licFile));
                }
                else
                {
                    //TODO : file not exists
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
            }
            catch (SecurityException se)
            {
                // display the reason for the license check failure:
               //TODO :
            }           
            return false;
        }      
       
    }
}
