using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization;
using System.Reflection;

namespace MachineConnectLicenseDTO
{

    public sealed class CurrentAssemblyDeserializationBinder : SerializationBinder
    {
        public override Type BindToType(string assemblyName, string typeName)
        {
            if (assemblyName.Contains("mscorlib") && typeName.Contains("[MachineConnectLicenseDTO.CNCData"))
            {
                return Type.GetType("System.Collections.Generic.List`1[[MachineConnectLicenseDTO.CNCData, MachineConnectSmartDataService]]");
            }
            else
            {            
                return Type.GetType(String.Format("{0}, {1}", typeName, Assembly.GetExecutingAssembly().FullName));
            }
        }
    }

    [Serializable]
    public class LicenseDTO
    {       
        public string LicenseTerms { get; set; }
        public string Signature { get; set; }


        #region Methods 
        internal static LicenseDTO FromString(string licenseTerms)
        {
            using (MemoryStream ms = new MemoryStream(Convert.FromBase64String(licenseTerms)))
            {
                BinaryFormatter bnfmt = new BinaryFormatter();
                bnfmt.Binder = new CurrentAssemblyDeserializationBinder();
                object value = bnfmt.Deserialize(ms);
                if (value is LicenseDTO)
                    return (LicenseDTO)value;
                else
                    throw new Exception("Invalid lic Type!");            
            }
        }
        #endregion
    }


    [Serializable]
    public class CNCData
    {
        public string CNCdata1 { get; set; }
        public string CNCdata2 { get; set; }
        public string CNCdata3 { get; set; }
        public bool IsOEM { get; set; }
    }

    [Serializable]
    public class LicenseTerms
    {
        public string Customer { get; set; }
        public string Plant { get; set; }
        public string Email { get; set; }
        public string LicType { get; set; }
        public string ExpiresAt { get; set; }       
        public string ComputerSerialNo { get; set; }       
        public List<CNCData> CNCData { get; set; }
        public string StartDate { get; set; }         

        #region Methods       
        internal static LicenseTerms FromString(string licenseTerms)
        {
            using (MemoryStream ms = new MemoryStream(Convert.FromBase64String(licenseTerms)))
            {
                BinaryFormatter bnfmt = new BinaryFormatter();
                bnfmt.Binder = new CurrentAssemblyDeserializationBinder();
                object value = bnfmt.Deserialize(ms);
                if (value is LicenseTerms)
                    return (LicenseTerms)value;
                else
                    throw new Exception("Invalid Lic Type!");

            }
        }
        #endregion
    }
}

       