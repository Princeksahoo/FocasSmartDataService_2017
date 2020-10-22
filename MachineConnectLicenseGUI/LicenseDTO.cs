using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using General.IO;
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
                return Type.GetType("System.Collections.Generic.List`1[[MachineConnectLicenseDTO.CNCData, MachineConnectLicenseGUI]]");
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
       
        public void Save(String fileName)
        {
            Serializer.Save<LicenseDTO>(this, fileName);
        }
      
        public void Save(Stream stream)
        {
            Serializer.Save<LicenseDTO>(this, stream);
        }
       
        public static LicenseDTO Load(String fileName)
        {
            // read the filename:
            return Serializer.Load<LicenseDTO>(new FileInfo(fileName));
        }
      
        public static LicenseDTO Load(Stream data)
        {
            // read the data stream:
            return Serializer.Load<LicenseDTO>(data);
        }
        public void SaveLicenseStringFinal(String fileName)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bnfmt = new BinaryFormatter();
                bnfmt.Serialize(ms, this);
                File.WriteAllText(fileName, Convert.ToBase64String(ms.GetBuffer()));
            }
        }
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
       // public int NoOfMachines { get; set; }
        public string ComputerSerialNo { get; set; }
       // public List<string> CNCSerialNumbers { get; set; }
        public List<CNCData> CNCData { get; set; }
        public string StartDate { get; set; }        

        #region Methods
        public String GetLicenseString()
        {
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bnfmt = new BinaryFormatter();
                bnfmt.Serialize(ms, this);
                return Convert.ToBase64String(ms.GetBuffer());
            }
        }

        public byte[] GetLicenseData()
        {
            using (MemoryStream ms = new MemoryStream())
            {
                BinaryFormatter bnfmt = new BinaryFormatter();
                bnfmt.Serialize(ms, this);
                return ms.GetBuffer();
            }
        }

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

       