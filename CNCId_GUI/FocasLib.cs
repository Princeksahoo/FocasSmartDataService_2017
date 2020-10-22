using System;
using System.Runtime.InteropServices;

namespace FocasLibrary
{
    public class FocasLib
    {
        [DllImport("FWLIB32.dll")]    
        public static extern short cnc_allclibhndl3([In, MarshalAs(UnmanagedType.AsAny)] object ip, ushort port, int timeout, out ushort FlibHndl);
        [DllImport("FWLIB32.dll")]        
        public static extern short cnc_freelibhndl(ushort FlibHndl);
        [DllImport("FWLIB32.dll")]
        public static extern short  cnc_rdcncid( ushort FlibHndl, uint [] cncid);
    }
}

