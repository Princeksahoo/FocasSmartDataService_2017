using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using System.Reflection;
using System.IO;

namespace MachineConnectLicenseDTO
{
    public static class Utility
    {
        public static RSACryptoServiceProvider GetPublicKeyFromAssemblyAtClient(Assembly assembly)
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

        public static RSACryptoServiceProvider GetRSAFromSnkFileAtServer(string snkFilePath)
        {
            byte[] snkFileBytes = File.ReadAllBytes(snkFilePath);
            RSACryptoServiceProvider rsaCsp = new RSACryptoServiceProvider();
            rsaCsp.ImportCspBlob(snkFileBytes);
            return rsaCsp;
        }
    }
}
