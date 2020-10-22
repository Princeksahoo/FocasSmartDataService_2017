using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EZSOCKETNCLib;
using EZNCAUTLib;
using FocasSmartDataCollection;

namespace TPMTrakSmartDataServiceForMazak
{
    public static class MitsubishiLib
    {
        static int M_ALM_ALL_ALARM = 0;
        static int EZNC_MAINPRG = 0;
        static int EZNC_S = 1;
        static DispEZNcCommunication EZNcCom = null;

        public static int SetTCPIPProtocol(string ipAddress, int portNo)
        {
            if (EZNcCom == null)
            {
                EZNcCom = new DispEZNcCommunication();
                return EZNcCom.SetTCPIPProtocol(ipAddress, 683);
            }
            return 0;
        }

        public static int Open2(string ipAddress, int portNo, int lSystemType, int lMachine, int lTimeOut)
        {
            int lResult = -1;
            if (EZNcCom != null) return 0;

            EZNcCom = new DispEZNcCommunication();
            lResult = EZNcCom.SetTCPIPProtocol(ipAddress, portNo);

            if (lResult == 0)
            {
                lResult = EZNcCom.Open2(lSystemType, lMachine, lTimeOut, "EZNC_LOCALHOST");
            }
            return lResult;
        }

        public static DateTime ReadCNCTimeStamp()
        {
            int plDate;
            int plTime;
            int lResult;
            DateTime CNCTimestamp = DateTime.MinValue;
            lResult = EZNcCom.Time_GetClockData(out plDate, out plTime);
            if (lResult == 0)
            {
                CNCTimestamp = new DateTime(plDate / 10000, (plDate % 10000) / 100, plDate % 100, plTime / 10000, (plTime % 10000) / 100, plTime % 100);
            }
            else
            {
                Logger.WriteDebugLog(string.Format("could not set timestamp from controller, using system time"));
                CNCTimestamp = DateTime.Now;
            }
            return CNCTimestamp;
        }

        public static DateTime ReadCNCCycleEndTimeStamp(int plDate, int plTime)
        {
            //int plDate;
            //int plTime;
            //int lResult;
            DateTime CNCTimestamp = DateTime.MinValue;

            int date = (int)ReadMacroParameter(plDate);
            int time = (int)ReadMacroParameter(plTime);
            string strDate = Utility.get_actual_date(date);
            string strTime = Utility.get_actual_time(time);
            DateTime.TryParse(strDate + " " + strTime, out CNCTimestamp);

            //lResult = EZNcCom.Time_GetClockData(out plDate, out plTime);
            //if (lResult == 0)
            //{
            //    CNCTimestamp = new DateTime(plDate / 10000, (plDate % 10000) / 100, plDate % 100, plTime / 10000, (plTime % 10000) / 100, plTime % 100);
            //}
            //else
            //{
            //    Logger.WriteDebugLog(string.Format("could not set timestamp from controller, using system time"));
            //    CNCTimestamp = DateTime.Now;
            //}


            return CNCTimestamp;
        }

        public static int Device_ReadBlock(int lLength, string strDevice, int lDataType, out object val)
        {
            return EZNcCom.Device_ReadBlock(2, "YC2C", 17, out val); // Emergency
            
        }


        public static string ReadMachineStatusMode(out string status, out short runningStatus, out string EmergencyStatus)
        {
            object val;
            int[] stat;
            int ret = -1;
            bool isrunning = false;
            runningStatus = 0;
            ret = EZNcCom.Device_ReadBlock(2, "YC2C", 17, out val);
            EmergencyStatus = "Unknown";
            if (ret == 0)
            {
                stat = (int[])val;
                if (stat.Length > 0)
                {
                    EmergencyStatus = stat[0] == 1 ? "Emergency" : "Reset";
                }
            }
            else
            {
                EmergencyStatus = "Error";
            }
            status = GetMachineStatus(out isrunning);
            runningStatus = isrunning ? (short)1 : (short)0;
            return GetMachineMode();
        }

        private static string GetMachineMode()
        {
            object pvValue1;
            string mode = "";
            int[] stat;
            int lResult = EZNcCom.Device_ReadBlock(2, "XC00", 17, out pvValue1); // jog mode
            if (lResult == 0)
            {
                stat = (int[])pvValue1;
                mode = stat[0] == 1 ? "JOG" : "";
            }
            
            if (mode.Equals(""))
            {
                if (lResult == 0)
                {
                    lResult = EZNcCom.Device_ReadBlock(2, "XC08", 17, out pvValue1); // mem
                    stat = (int[])pvValue1;
                    mode = stat[0] == 1 ? "MEM" : "";
                }
            }
            if (mode.Equals(""))
            {
                lResult = EZNcCom.Device_ReadBlock(2, "XC0b", 17, out pvValue1); // mdi
                if (lResult == 0)
                {
                    stat = (int[])pvValue1;
                    mode = stat[0] == 1 ? "MDI" : "";
                }
            }
            if (mode.Equals(""))
            {
                lResult = EZNcCom.Device_ReadBlock(2, "XC09", 17, out pvValue1); // edit
                if (lResult == 0)
                {
                    stat = (int[])pvValue1;
                    mode = stat[0] == 1 ? "EDIT" : "";
                }
            }
            return mode.Equals("") ? "Error" : mode;
        }

        //private static string GetMachineStatus(out bool isStarted)
        //{
        //    int[] stat = new int[2];
        //    int lResult;
        //    object val;
        //    string status = "";
        //    isStarted = false;

        //    lResult = EZNcCom.Device_ReadBlock(2, "XC12", 17, out val); // start & running
        //    stat = (int[])val;
        //    if (lResult == 0)
        //    {
        //        status = stat[0] == 1 ? "In Cycle" : "";
        //        isStarted = stat[0] == 1;
        //    }
        //    if (status.Equals(""))
        //    {
        //        lResult = EZNcCom.Device_ReadBlock(2, "XC13", 17, out val); // start
        //        stat = (int[])val;
        //        if (lResult == 0)
        //        {
        //            status = stat[0] == 1 ? "Idle" : "";
        //            isStarted = stat[0] == 1;
        //        }
        //    }

        //    if (status.Equals(""))
        //    {
        //        lResult = EZNcCom.Device_ReadBlock(2, "XC14", 17, out val); // pause
        //        stat = (int[])val;
        //        if (lResult == 0)
        //        {
        //            status = stat[0] == 1 ? "STOP" : "";
        //        }
        //    }

        //    if (status.Equals(""))
        //    {
        //        lResult = EZNcCom.Device_ReadBlock(2, "XC15", 17, out val); // MSTR
        //        stat = (int[])val;
        //        if (lResult == 0)
        //        {
        //            status = stat[0] == 1 ? "MSTR" : "";
        //        }
        //    }
        //    //lResult = EZNcCom.Device_ReadBlock(2, "XC11", 17, out pvValue1); // Servo ready
        //    //lResult = EZNcCom.Device_ReadBlock(2, "Y1898", 17, out pvValue1); // IWP: Spindle on
        //    //lResult = EZNcCom.Device_ReadBlock(2, "Y1899", 17, out pvValue1); // Rev: Spindle on

        //    return status.Equals("") ? "Unknown" : status;
        //}

        private static string GetMachineStatus(out bool isStarted)
        {
            /*
           |------+------+------+----------+----------+------------|
           |      | Idle | Stop | FeedHold | In-cycle | MSTR (RST) |
           |------+------+------+----------+----------+------------|
           | XC12 |    0 |    1 |        1 |        1 |          0 |
           | XC13 |    0 |    0 |        1 |        1 |          0 |
           | XC14 |    0 |    1 |        0 |        0 |          0 |
           | XC15 |    0 |    0 |        0 |        0 |          1 |
           | XC9B |    0 |    0 |        1 |        0 |          0 |
           |------+------+------+----------+----------+------------| 
            */
            int[] stat = new int[2];
            int lResult;
            object val;
            string status = "";
            isStarted = false;
            int statVal = 0; // combine the returned values and compare

            lResult = EZNcCom.Device_ReadBlock(2, "XC12", 17, out val); // start & running
            stat = (int[])val;
            if (lResult == 0)
            {
                statVal |= stat[0];
            }
            else
            {
                status = "Error";
            }

            lResult = EZNcCom.Device_ReadBlock(2, "XC13", 17, out val); // start
            stat = (int[])val;
            if (lResult == 0)
            {
                statVal |= (stat[0] << 1);
            }
            else
            {
                status = "Error";
            }

            lResult = EZNcCom.Device_ReadBlock(2, "XC14", 17, out val); // pause
            stat = (int[])val;
            if (lResult == 0)
            {
                statVal |= (stat[0] << 2);
            }
            else
            {
                status = "Error";
            }

            lResult = EZNcCom.Device_ReadBlock(2, "XC15", 17, out val); // MSTR
            stat = (int[])val;
            if (lResult == 0)
            {
                statVal |= (stat[0] << 3);
            }
            else
            {
                status = "Error";
            }

            lResult = EZNcCom.Device_ReadBlock(2, "XC9B", 17, out val); // Emergency - on feed hold
            stat = (int[])val;
            if (lResult == 0)
            {
                statVal |= (stat[0] << 4);
            }
            else
            {
                status = "Error";
            }

            if (!status.Equals("Error", StringComparison.OrdinalIgnoreCase))
            {
                switch(statVal)
                {
                    case 0:
                        status = "Idle";
                        break;
                    case 5:
                        status = "STOP";
                        break;
                    case 19:
                        status = "Feed Hold";
                        isStarted = true;
                        break;
                    case 3:
                        status = "In Cycle";
                        isStarted = true;
                        break;
                    case 8:
                        status = "MSTR";
                        break;
                    default:
                        status = "";
                        break;
                }
            }
            return status.Equals("") ? "Error" : status;
        }

        internal static string GetProgramNo()
        {
            int lResult;
            string szPrgNo = "";
            lResult = EZNcCom.Program_GetProgramNumber2(EZNC_MAINPRG, out szPrgNo);
            return szPrgNo;
        }

        internal static int GetSpindleSpeed()
        {
            int speed = 0;
            int lResult;
            lResult = EZNcCom.Command_GetCommand2(EZNC_S, 1, out speed);
            return speed;
        }

        internal static int GetSpindleLoad(out int ret)
        {
            int plData;
            string pbstrBuffer;
            ret = EZNcCom.Monitor_GetSpindleMonitor(3, 1, out plData, out pbstrBuffer);
            return plData;
        }

        //internal static List<DTO.LiveAlarm> ReadLiveAlarms(DateTime cncTS)
        //{
        //    List<DTO.LiveAlarm> lst = new List<DTO.LiveAlarm>();
        //    string AlarmMsg;
        //    int lResult;
        //    lResult = EZNcCom.System_GetAlarm2(3, M_ALM_ALL_ALARM, out AlarmMsg);
        //    string[] spl = AlarmMsg.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
        //    int i = 1;
        //    foreach (string s in spl)
        //    {
        //        string[] splAlm = s.Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries);
        //        lst.Add(new DTO.LiveAlarm() {  AlarmNo=splAlm.Length > 0 ? splAlm[0] : i.ToString(), StartTime=cncTS, Message=s});
        //        i += 1;
        //    }
        //    return lst;
        //}

        internal static int GetPartsCount()
        {
            object partsCount;
            int pc = 0;
            int lResult = EZNcCom.Parameter_GetData3(1, 8002, 1, 1, out partsCount);
            if (lResult == 0)
            {
                pc = int.Parse(((string[])partsCount)[0]);
            }
            return pc;
        }

        internal static int GetCycleTime()
        {
            object cyctim = 0;
            int lResult = EZNcCom.Generic_ReadData(1, 40, 8, ref cyctim);
            return (int)cyctim;
        }

        internal static void Close()
        {
            int lResult = -1;
            if (EZNcCom != null)
            {
                lResult = EZNcCom.Close();
                EZNcCom = null;
            }
        }
        public static double ReadMacroParameter(int macroAddress)
        {
            int retValue=0;
            double paravalue=0;
            try
            {
                
                int AA = 1;
                retValue = EZNcCom.CommonVariable_Read2(macroAddress, out paravalue, out AA);
                //var ret = EZNcCom.CommonVariable_Write2(500, 600, 1);
                if (retValue != 0)
                {
                    return double.MinValue;
                }
                
            }
            catch(Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            return paravalue;
        }

        public static void WriteMacroParameter(int macroAddress, double macroValue)
        {
            int retValue = 0;
            try
            {

                int AA = 1;
               // retValue = EZNcCom.CommonVariable_Read2(macroAddress, out paravalue, out AA);
                retValue = EZNcCom.CommonVariable_Write2(macroAddress, macroValue, 1);
                if (retValue != 0)
                {

                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
        }
    }
}
