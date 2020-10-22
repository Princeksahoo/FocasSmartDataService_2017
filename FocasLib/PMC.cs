using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using FocasLib;
namespace FocasLibrary
{
    public static partial class FocasData
    {
        //Reads the contents of the operator's message in CNC.
        public static List<string> ReadPMCMessages(ushort handle)
        {
            List<string> msgs = new List<string>();
            FocasLibBase.OPMSG message = new FocasLibBase.OPMSG();
            short length = (short) Marshal.SizeOf(message);
            int ret = FocasLib.cnc_rdopmsg(handle, 0, 6 + 256, message);
            if (ret == 0)
            {
                if (message.msg1.datano != -1) msgs.Add(message.msg1.data);
                if (message.msg2.datano != -1) msgs.Add(message.msg2.data);
                if (message.msg3.datano != -1) msgs.Add(message.msg3.data);
                if (message.msg4.datano != -1) msgs.Add(message.msg4.data);
                if (message.msg5.datano != -1) msgs.Add(message.msg5.data);
                if (message.msg6.datano != -1) msgs.Add(message.msg6.data);
            }
            return msgs;
        }

        //Read PMC alarm messages
        public static List<string> ReadPMCAlarmMessage(ushort handle)
        {
            List<string> msgs = new List<string>();
            FocasLibBase.ODBPMCALM message = new FocasLibBase.ODBPMCALM();
            short recordsExists = 0;
            short s_number = 1;
            short read_num = 10;
            int ret = FocasLib.pmc_rdalmmsg(handle, s_number, ref read_num, out recordsExists, message);
            if (ret == 0)
            {
                if (message.msg1.almmsg != string.Empty) msgs.Add(message.msg1.almmsg);
                if (message.msg2.almmsg != string.Empty) msgs.Add(message.msg2.almmsg);
                if (message.msg3.almmsg != string.Empty) msgs.Add(message.msg3.almmsg);
                if (message.msg4.almmsg != string.Empty) msgs.Add(message.msg4.almmsg);
                if (message.msg5.almmsg != string.Empty) msgs.Add(message.msg5.almmsg);
                if (message.msg6.almmsg != string.Empty) msgs.Add(message.msg6.almmsg);
                if (message.msg7.almmsg != string.Empty) msgs.Add(message.msg7.almmsg);
                if (message.msg8.almmsg != string.Empty) msgs.Add(message.msg8.almmsg);
                if (message.msg9.almmsg != string.Empty) msgs.Add(message.msg9.almmsg);
                if (message.msg10.almmsg != string.Empty) msgs.Add(message.msg10.almmsg);
            }
            return msgs;
        }

        public static string GetFeedrateOverride(ushort handle)
        {
            FocasLibBase.IODBPMC0 f = new FocasLibBase.IODBPMC0();
            FocasLib.pmc_rdpmcrng(handle, 1, 0, 3, 3, 9, f);
            if (f.cdata[0] < 8)
            {
                return "Manual Mode";
            }
            f = new FocasLibBase.IODBPMC0();
            FocasLib.pmc_rdpmcrng(handle, 0, 0, 12, 19, 16, f);
            int num = 0xff - f.cdata[0];
            return (num.ToString() + "%");
        }

        public static string GetRapidTraverseOverride(ushort handle)
        {
            string str = string.Empty;
            FocasLibBase.IODBPMC0 f = new FocasLibBase.IODBPMC0();
            FocasLibBase.IODBPMC0 iodbpmc2 = new FocasLibBase.IODBPMC0();
            FocasLib.pmc_rdpmcrng(handle,0, 0, 14, 15, 10, f);//pmc_rdpmcrng(0, 0, 14, 15, 10, f);
            FocasLib.pmc_rdpmcrng(handle, 0, 0, 96, 103, 16, iodbpmc2);//pmc_rdpmcrng(0, 0, 0x60, 0x67, 0x10, iodbpmc2);
            if ((iodbpmc2.cdata[0] & 128) == 128)
            {
                return iodbpmc2.cdata[0].ToString();
            }
            switch (f.cdata[0])
            {
                case 0:
                    return "100%";

                case 1:
                    return "50%";

                case 2:
                    return "25%";

                case 3:
                    return "Slow";
            }
            return str;
        }

        public static string GetSpindleOverride(ushort handle)
        {
            FocasLibBase.IODBPMC0 f = new FocasLibBase.IODBPMC0();
            FocasLib.pmc_rdpmcrng(handle, 0, 0, 30, 37, 16, f);
            return f.cdata[0].ToString() + "%";
        }

        public static string GetTPMTrakFlagStatus(ushort handle)
        {
            FocasLibBase.IODBPMC0 f = new FocasLibBase.IODBPMC0();
            FocasLib.pmc_rdpmcrng(handle, 5, 0, 600, 600, 9, f);
            //Use the  first byte f.cdata[0] to check the pm trak flag on CNC machine and update it
            return f.cdata[0].ToString();
        }

        public static void GetCoolantLubricantLevel(ushort handle,ushort startingLocation, ushort endLocation, out short coolentLevel, out short lubOilLevel)
        {
            /*
                R0200,..,R0209(It is assumed the byte type) is read.
                adr_type 5 
                data_type 0 
                s_number 200 
                e_number 209 
                length 8+1×10 (=18) 
                buf.u.cdata[0]
                 ,..,buf.u.cdata[9] The contents of R0200,..,R0209 are stored. 
             FWLIBAPI short WINAPI pmc_rdpmcrng(unsigned short FlibHndl, short adr_type, short data_type, short s_number, short e_number, short length, IODBPMC *buf); 
             */
            FocasLibBase.IODBPMC1 f = new FocasLibBase.IODBPMC1();
            short ret = short.MinValue;
            coolentLevel = 0;
            lubOilLevel = 0;
            ret = FocasLib.pmc_rdpmcrng(handle, 5, 0, startingLocation, endLocation, (8 + 2 * 2), f);
            if (ret != 0)
            {
                ret = FocasLib.pmc_rdpmcrng(handle, 5, 0, 662, 665, (8 + 2 * 2), f);
                Logger.WriteErrorLog("pmc_rdpmcrng() failed. return value is = " + ret);
                return;
            }
            coolentLevel = f.idata[1];
            lubOilLevel = f.idata[0];
           // return "Coolent = " + ((f.idata[0] / 4095.0 ) * 100.0).ToString();
        }

        public static void GetPredictiveMaintenanceTargetCurrent(ushort handle, ushort startingLocation, ushort endLocation, out int targetValue, out int currentValue, string MTB = "ACE")
        {
            /*
                R0200,..,R0209(It is assumed the byte type) is read.
                adr_type 5 
                data_type 0 
                s_number 200 
                e_number 209 
                length 8+1×10 (=18) 
                buf.u.cdata[0]
                 ,..,buf.u.cdata[9] The contents of R0200,..,R0209 are stored. 
             FWLIBAPI short WINAPI pmc_rdpmcrng(unsigned short FlibHndl, short adr_type, short data_type, short s_number, short e_number, short length, IODBPMC *buf); 
             */
            FocasLibBase.IODBPMC1 f = new FocasLibBase.IODBPMC1();
            short ret = short.MinValue;
            targetValue = int.MaxValue;
            currentValue = int.MaxValue;

            if (MTB.Equals("ACE", StringComparison.OrdinalIgnoreCase))
            {
                ret = FocasLib.pmc_rdpmcrng(handle, 9, 1, startingLocation, endLocation, (8 + 2 * 2), f);
                if (ret != 0)
                {
                    Logger.WriteErrorLog(string.Format("pmc_rdpmcrng() failed. return value is = {0} . Read D parameters {1} - {2}", ret, startingLocation, endLocation));
                    return;
                }
                targetValue = f.idata[0];
                currentValue = f.idata[1];
            }
            else if (MTB.Equals("AMS", StringComparison.OrdinalIgnoreCase))
            {
                ushort tempEndLocation = (ushort)(startingLocation + 2);
                ret = FocasLib.pmc_rdpmcrng(handle, 9, 1, startingLocation, tempEndLocation, (8 + 2 * 2), f);
                if (ret != 0)
                {
                    Logger.WriteErrorLog(string.Format("pmc_rdpmcrng() failed. return value is = {1} . read D parameter {0}  ", startingLocation, ret));
                    return;
                }
                targetValue = f.idata[0];

                tempEndLocation = (ushort)(endLocation + 2);
                ret = FocasLib.pmc_rdpmcrng(handle, 9, 1, endLocation, tempEndLocation, (8 + 2 * 2), f);
                if (ret != 0)
                {                    
                    Logger.WriteErrorLog(string.Format("pmc_rdpmcrng() failed. return value is = {1} . read D parameter {0}  ", endLocation, ret));
                    return;
                }
                currentValue = f.idata[0];
            }
        }


        public static List<int> ReadPMCDataTableInt(ushort handle, ushort startingAddress, ushort endingAddress)
        {
            FocasLibBase.IODBPMC2 f = new FocasLibBase.IODBPMC2();
            FocasLib.pmc_rdpmcrng(handle, 9, 1, startingAddress, endingAddress, 8 + (20 * 2), f);
            return f.ldata.ToList();
        }

        public static void WritePMCDataTableInt(ushort handle, ushort startingAddress, ushort endingAddress)
        {
            FocasLibBase.IODBPMC2 f = new FocasLibBase.IODBPMC2();
            f.ldata = new int[10];
            f.type_a = 9;
            f.type_d = 1;
            f.datano_s = (short)startingAddress;
            f.datano_e = (short)endingAddress;
            for (int i = 0; i < 10; i++)
            {
                f.ldata[i] = 0;
            }
            var ret = FocasLib.pmc_wrpmcrng(handle, 8 + (20 * 2), f);
        }

        public static List<short> ReadPMCDataTableWord(ushort handle, ushort startingAddress, ushort endingAddress)
        {
            FocasLibBase.IODBPMC1 f = new FocasLibBase.IODBPMC1();
            FocasLib.pmc_rdpmcrng(handle, 9, 1, startingAddress, endingAddress, 8 + (10 * 2), f);
            return f.idata.ToList();
        }

        public static void WritePMCDataTableWord(ushort handle, ushort startingAddress, ushort endingAddress) // to write spindle load 0
        {
            FocasLibBase.IODBPMC1 f = new FocasLibBase.IODBPMC1();
            f.idata = new short[10];
            f.type_a = 9;
            f.type_d = 1;
            f.datano_s = (short)startingAddress;
            f.datano_e = (short)endingAddress;
            for (int i = 0; i < 10; i++)
            {
                f.idata[i] = 0;
            }
            FocasLib.pmc_wrpmcrng(handle, 8 + (10 * 2), f);
        }

        public static List<byte> ReadPMCDataTableChar(ushort handle, ushort startingAddress, ushort endingAddress)
        {
            FocasLibBase.IODBPMC0 f = new FocasLibBase.IODBPMC0();
            FocasLib.pmc_rdpmcrng(handle, 9, 1, startingAddress, endingAddress, 8 + (10 * 1), f);
            string str = string.Empty;
            return f.cdata.ToList();
        }  
        

        public static List<byte> ReadPMCRangeByte(ushort handle, short addType, ushort startingAddress, ushort endingAddress)
        {
            //3 = X adr_type , 2 = Y adr_type 
            //data_type byte = 0
            //pmc_rdpmcrng(unsigned short FlibHndl, short adr_type, short data_type, short s_number, short e_number, short length, IODBPMC *buf); 
            FocasLibBase.IODBPMC0 f = new FocasLibBase.IODBPMC0();
            short ret = FocasLib.pmc_rdpmcrng(handle, addType, 0, startingAddress, endingAddress, (ushort)(8 + (endingAddress - startingAddress) + 1), f);
            //if (ret != 0) throw new DivideByZeroException();
            return f.cdata.ToList<byte>();

        }


        public static short ReadPMCOneWord(ushort handle, short addType, ushort startingAddress, ushort endingAddress)
        {
            //3 = X adr_type , 2 = Y adr_type 
            //data_type byte = 0
            //pmc_rdpmcrng(unsigned short FlibHndl, short adr_type, short data_type, short s_number, short e_number, short length, IODBPMC *buf); 
            FocasLibBase.IODBPMC1 f = new FocasLibBase.IODBPMC1();
            short ret = FocasLib.pmc_rdpmcrng(handle, addType, 1, startingAddress, endingAddress, (ushort)(8 + (endingAddress - startingAddress) * 2), f);
            //if (ret != 0) throw new DivideByZeroException();
            return f.idata[0];

        }

        public static List<short> ReadPMCRangeWord(ushort handle, short addType, ushort startingAddress, ushort endingAddress)
        {
            //3 = X adr_type , 2 = Y adr_type 
            //data_type byte = 0
            //pmc_rdpmcrng(unsigned short FlibHndl, short adr_type, short data_type, short s_number, short e_number, short length, IODBPMC *buf); 
            FocasLibBase.IODBPMC1 f = new FocasLibBase.IODBPMC1();
            short ret = FocasLib.pmc_rdpmcrng(handle, addType, 1, startingAddress, endingAddress, (ushort)(8 + (endingAddress - startingAddress) * 2), f);
            //if (ret != 0) throw new DivideByZeroException();
            return f.idata.ToList<short>();

        }

    }
}
