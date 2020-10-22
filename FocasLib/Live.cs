using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Diagnostics;
using FocasLib;
using System.Linq;
using DTO;

namespace FocasLibrary
{

    public static partial class FocasData
    {      
        public static short cnc_allclibhndl3(string ip, ushort port, int timeout,out ushort handle)
        {
            return FocasLib.cnc_allclibhndl3(ip, port, timeout, out handle);
        }

        public static short cnc_freelibhndl(ushort handle)
        {
            return FocasLib.cnc_freelibhndl(handle);
        }
        public static short ReadActiveProgram(ushort handle)
        {
            FocasLibBase.ODBPRO odbpro = new FocasLibBase.ODBPRO();
            FocasLib.cnc_rdprgnum(handle,odbpro);
            return odbpro.data;
        }

        public static short ReadMainProgram(ushort handle)
        {
            FocasLibBase.ODBPRO odbpro = new FocasLibBase.ODBPRO();
            FocasLib.cnc_rdprgnum(handle, odbpro);
            return odbpro.mdata;
        }

        /// <summary>
        /// Get Active and Main program No
        /// </summary>
        /// <param name="Main Program"></param>
        /// <returns>ActiveProgramNo</returns>
        public static short ReadActiveMainProgramNo(ushort handle,out short mainProgram)
        {
            FocasLibBase.ODBPRO odbpro = new FocasLibBase.ODBPRO();
            FocasLib.cnc_rdprgnum(handle, odbpro);
            mainProgram = odbpro.mdata;
            return odbpro.data;
        }


        public static string ReadMachineMode(ushort handle)
        {
            FocasLibBase.ODBST a = new FocasLibBase.ODBST();
            FocasLib.cnc_statinfo(handle, a);
            return Utility.GetMachineMode(a);

        }

        public static string ReadMachineStatus(ushort handle)
        {
            Utility.CncMachineState _cncMachineState = Utility.CncMachineState.Unknown;
            FocasLibBase.ODBST a = new FocasLibBase.ODBST();
            FocasLib.cnc_statinfo(handle, a);
            return Utility.GetMachineStatus(out _cncMachineState, a);
        }

        /// <summary>
        /// Read Machine Status and Mode
        /// </summary>
        /// <param name="Out MachineStatus"></param>
        /// <returns>MachineMode</returns>
        public static string ReadMachineStatusMode(ushort handle,out string status)
        {
            Utility.CncMachineState _cncMachineState = Utility.CncMachineState.Unknown;
            FocasLibBase.ODBST a = new FocasLibBase.ODBST();
            FocasLib.cnc_statinfo(handle, a);
            status = Utility.GetMachineStatus(out _cncMachineState, a);
            return Utility.GetMachineMode(a);
        }

        /// <summary>
        /// Read RPM(Spindle Speed) and FeedRate
        /// </summary>
        /// <param name="feedRate"></param>
        /// <returns>RPM</returns>
        public static double ReadRPMFeedRate(ushort handle,out double feedRate, out string feedRateUnit, out string spinddleUnit)
        {
            FocasLibBase.ODBSPEED b = new FocasLibBase.ODBSPEED();
            FocasLib.cnc_rdspeed(handle, -1, b);
            feedRate = (((double)b.actf.data) / Math.Pow(10.0, (double)b.actf.dec));
            feedRateUnit = b.actf.unit.ToString();
            spinddleUnit = b.acts.unit.ToString();
            return (((double)b.acts.data) / Math.Pow(10.0, (double)b.acts.dec));
            /*
               Unit 0 : mm/min 
                    1 : inch/min 
                    2 : rpm 
                    3 : mm/rev 
                    4 : inch/rev 
             */
        }

        public static int ReadToolNo(ushort handle)
        {
            return (int)ReadMacro(handle, 4120);
        }

        public static int ReadOffsetNo(ushort handle)
        {
            int tool = (int)ReadMacro(handle, 4120);
            return tool % 100;
        }
        /// <summary>
        /// Read tool and offset no
        /// </summary>
        /// <param name="offsetNoforCurrectRunningTool"></param>
        /// <returns>ToolNo</returns>
        public static int ReadTool_Offset(ushort handle,out int offset)
        {
            int tool = ReadMacro(handle, 4120);
            offset = tool % 100;
            return tool;
        }        

        public static double ReadCuttingTime(ushort handle)
        {           
            long minute = ReadParameterInt(handle, 6754);
            long ms = ReadParameterInt(handle, 6753);
            TimeSpan ts = TimeSpan.FromMinutes(minute) + TimeSpan.FromMilliseconds(ms);                
            return ts.TotalSeconds;
        }

        public static double ReadOperatorTime(ushort handle)
        {
            long minute = ReadParameterInt(handle, 6752);
            long ms = ReadParameterInt(handle, 6751);
            TimeSpan ts = TimeSpan.FromMinutes(minute) + TimeSpan.FromMilliseconds(ms);           
            return ts.TotalSeconds;
        }

        public static DateTime ReadCNCTimeStamp(ushort handle)
        {           
            DateTime cncDateTime = DateTime.MinValue;
            int date = (int)ReadMacroDouble(handle, 3011);
            int time = (int)ReadMacroDouble(handle, 3012);
            string strDate = Utility.get_actual_date(date);
            string strTime = Utility.get_actual_time(time);
            DateTime.TryParse(strDate + " " + strTime, out cncDateTime);           
            return cncDateTime;
        }

        public static DateTime ReadCNCTimeStamp(ushort handle,short dateL, short timeL )
        {
            DateTime cncDateTime = DateTime.MinValue;
            int date = (int)ReadMacroDouble(handle, dateL);
            int time = (int)ReadMacroDouble(handle, timeL);
            string strDate = Utility.get_actual_date(date);
            string strTime = Utility.get_actual_time(time);
            DateTime.TryParse(strDate + " " + strTime, out cncDateTime);
            return cncDateTime;
        }

        public static int ReadPartsCount(ushort handle)
        {
            int count = ReadParameterInt(handle, 6712);
            return count;
        }
               
        public static FocasLibBase.ODBDY2_1 cnc_rddynamic2(ushort handle)
        {
            FocasLibBase.ODBDY2_1 obj = new FocasLibBase.ODBDY2_1();
            int ret = FocasLib.cnc_rddynamic2(handle, 4, 540, obj);
            return obj;
        }

        public static short ReadSpindleLoad(ushort handle)
        {
            FocasLibBase.ODBDGN_1 obj = new FocasLibBase.ODBDGN_1();
            short ret = FocasLib.cnc_diagnoss(handle, 410, 1, 8, obj);
            return obj.idata;
        }       

        public static int ReadServoMotorTempX2(ushort handle)
        {
            FocasLibBase.ODBDGN_1 obj = new FocasLibBase.ODBDGN_1();
            short ret = FocasLib.cnc_diagnoss(handle, 403, 1, 8, obj);//temp spindle motor 403
            return obj.ldata;
        }

        public static int ReadServoMotorTempY2(ushort handle)
        {
            FocasLibBase.ODBDGN_1 obj1 = new FocasLibBase.ODBDGN_1();
            short ret = FocasLib.cnc_diagnoss(handle, 403, 2, 8, obj1);//temp spindle motor 403
            return obj1.ldata;
        }
        
        public static long ReadPowerOnTime(ushort handle)
        {
            return ReadParameterInt(handle, 6750);
        }        
       
        public static short GetFanucMachineType(ushort handle, ref CncMachineType cncMachineType, out string cncSeries)
        {
            short num = 0;
            string p = string.Empty;
            cncSeries = string.Empty;
            FocasLibBase.ODBSYS a = new FocasLibBase.ODBSYS();
            num = FocasLib.cnc_sysinfo(handle, a);
            if (num != 0)
            {
                return num;
            }
            else
            {
                p = new string(a.cnc_type);
                switch (p)
                {
                    case "15":
                        cncMachineType = CncMachineType.Series150i;
                        break;

                    case "16":
                        cncMachineType = CncMachineType.Series160i;
                        break;

                    case "18":
                        cncMachineType = CncMachineType.Series180i;
                        break;

                    case "21":
                        cncMachineType = CncMachineType.Series210i;
                        break;

                    case "30":
                        cncMachineType = CncMachineType.Series300i;
                        break;

                    case "31":
                        cncMachineType = CncMachineType.Series310i;
                        break;

                    case "32":
                        cncMachineType = CncMachineType.Series320i;
                        break;

                    case " 0":
                        cncMachineType = CncMachineType.Series0i;
                        break;

                    case "PD":
                        cncMachineType = CncMachineType.SeriesPowerMateiD;
                        break;

                    case "PH":
                        cncMachineType = CncMachineType.SeriesPowerMateiH;
                        break;

                    default:
                        cncMachineType = CncMachineType.cncUnknown;                        
                        break;
                }
                
                p = new string(a.mt_type);
                switch (p)
                {
                    case " M":
                        cncSeries = "Machining center";
                        break;

                    case " T":
                        cncSeries = "Lathe";
                        break;

                    case "MM":
                        cncSeries = "M series with 2 path control";
                        break;

                    case "TT":
                        cncSeries = "T series with 2/3 path control";
                        break;

                    case "MT":
                        cncSeries = "T series with compound machining function";
                        break;

                    case " P":
                        cncSeries = "Punch press";
                        break;

                    case " L":
                        cncSeries = "Laser";
                        break;

                    case " W":
                        cncSeries = "Wire cut";
                        break;

                    default:
                        cncSeries = p;
                        break;
                }
            }
            return 0;
        }

        public static bool SetCNCDate(ushort handle,DateTime date)
        {
            FocasLibBase.IODBTIMER a = new FocasLibBase.IODBTIMER();
            a.type = 0; // 0 = Date ; 1 = time
            a.date.date =  (short)date.Day;
            a.date.month = (short)date.Month;
            a.date.year = (short)date.Year;
            short ret = FocasLib.cnc_settimer(handle, a);
            if (ret != 0)
            {
                Logger.WriteErrorLog("Setting CNC date failed (cnc_settimer() return value is = " + ret);
                return false;
            }
            return true;            
        }
        public static bool SetCNCTime(ushort handle, DateTime date)
        {
            FocasLibBase.IODBTIMER a = new FocasLibBase.IODBTIMER();
            a.type = 1; // 0 = Date ; 1 = time
            a.time.hour = (short)date.Hour;
            a.time.minute = (short)date.Minute;
            a.time.second = (short)date.Second;
            short ret = FocasLib.cnc_settimer(handle, a);
            if (ret != 0)
            {
                Logger.WriteErrorLog("Setting CNC Time failed (cnc_settimer() return value is = " + ret);
                return false;
            }
            return true;            
        }

        public static bool GetCNCTime(ushort handle, out TimeSpan time)
        {
            time = TimeSpan.Zero;
            FocasLibBase.IODBTIMER a = new FocasLibBase.IODBTIMER();
            a.type = 1; // 0 = Date ; 1 = time           
            short ret = FocasLib.cnc_gettimer(handle, a);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_gettimer() failed. return value is = " + ret);
                return false;
            }
            time = new TimeSpan(a.time.hour, a.time.minute, a.time.second);
            return true;
        }

        public static bool GetCNCDate(ushort handle, out DateTime date)
        {
            date = DateTime.MinValue;
            FocasLibBase.IODBTIMER a = new FocasLibBase.IODBTIMER();
            a.type = 0; // 0 = Date ; 1 = time           
            short ret = FocasLib.cnc_gettimer(handle, a);
            if (ret != 0)
            {
                Logger.WriteErrorLog("cnc_gettimer() failed. return value is = " + ret);
                return false;
            }
            date = new DateTime(a.date.year, a.date.month, a.date.date);
            return true;
        }
        
     
        public static void ClearOperationHistory(ushort handle)
        {
           int ret = FocasLib.cnc_clearophis(handle,0);
           ret = FocasLib.cnc_clearomhis(handle);
        }

        public static int ReadAxixServoMotorTempX(ushort handle)
        {
            FocasLibBase.ODBDGN_1 obj = new FocasLibBase.ODBDGN_1();
            short ret = FocasLib.cnc_diagnoss(handle, 308, 1, 8, obj);//temp spindle motor 308
            Debug.WriteLine("AXIS X temp=" + obj.ldata);
            return obj.ldata;
        }
       
        public static int ReadSpindleMotorTemp(ushort handle)
        {
            FocasLibBase.ODBDGN_1 obj = new FocasLibBase.ODBDGN_1();
            short ret = FocasLib.cnc_diagnoss(handle, 403, 1, 8, obj);//temp spindle motor 408           
            return obj.ldata;
        }

        public static int ReadSpindleSpeed(ushort handle)
        {
            FocasLibBase.ODBACT obj = new FocasLibBase.ODBACT();
            short ret = FocasLib.cnc_acts(handle, obj);
            return obj.data;
        }

        private const short X_WEAR_DEFAULT_MACRO = 2000;
        private const short Z_WEAR_DEFAULT_MACRO = 2100;
        private const short R_WEAR_DEFAULT_MACRO = 2200;
        private const short T_WEAR_DEFAULT_MACRO = 2300;

        private const short X_GEOMETRY_DEFAULT_MACRO = 2700;
        private const short Y_GEOMETRY_DEFAULT_MACRO = 2800;
        private const short R_GEOMETRY_DEFAULT_MACRO = 2900;
        private const short T_GEOMETRY_DEFAULT_MACRO = 2300;
        public static double get_wear_offset(ushort handle, int offsetno, char axis)
        {
            double act_val = 0;
            if (offsetno == 0)
            {
                return 0;
            }
            short nmbr = 0;
            switch (axis)
            {

                case 'x':
                    nmbr = (short)(X_WEAR_DEFAULT_MACRO + offsetno);
                    break;

                case 'z':
                    nmbr = (short)(Z_WEAR_DEFAULT_MACRO + offsetno);
                    break;

                case 'r':
                    nmbr = (short)(R_WEAR_DEFAULT_MACRO + offsetno);
                    break;

                case 't':
                    nmbr = (short)(T_WEAR_DEFAULT_MACRO + offsetno);
                    break;
                default:
                    return 0;
            }

            act_val = ReadMacroDouble(handle, nmbr);
            return act_val;

        }
        public static double get_geometry_offset(ushort handle, int offsetno, char axis)
        {
            double act_val = 0;
            if (offsetno == 0)
            {
                return 0;
            }
            short nmbr = 0;
            switch (axis)
            {
                case 'x':
                    nmbr = (short)(X_GEOMETRY_DEFAULT_MACRO + offsetno);
                    break;

                case 'y':
                    nmbr = (short)(Y_GEOMETRY_DEFAULT_MACRO + offsetno);
                    break;

                case 'r':
                    nmbr = (short)(R_GEOMETRY_DEFAULT_MACRO + offsetno);
                    break;

                case 't':
                    nmbr = (short)(T_GEOMETRY_DEFAULT_MACRO + offsetno);
                    break;

                default:
                    return 0;

            }
            act_val = ReadMacroDouble(handle, nmbr);
            return act_val;
        }

        /// <summary>
        /// ReadServo Load and Current Details on All axis
        /// </summary>
        /// <param name="current">Current in Amp</param>
        /// <returns>Servo Load</returns>
        public static string ReadServoLoadCurrentDetails(ushort handle, out string current)
        {
            current = string.Empty;
            string buff = "";
            //string sb = "";
            string DL = "#";
            short[] types = { 0, 1, 2 };
            /*
             cls = 2 (Servo) 
             type = 0 : Servo load meter 
                    1 : Load current (% unit) 
                    2 : Load current (Ampere unit) 
                    */

            short MAX_AXIS = FocasLibBase.MAX_AXIS;
            short len = MAX_AXIS;
            FocasLibBase.ODBAXDT odbaxdt = new FocasLibBase.ODBAXDT();

            List<string> pos = new List<string>();
            //1st parameter - Position(=1), Servo(=2), Spindle(=3)
            //4th parameter - After execution, "(*len)" will have the actual number of axes.
            short ret = FocasLib.cnc_rdaxisdata(handle, 2, types, (short)types.Length, ref len, odbaxdt);
            if (ret == 0)
            {
                MAX_AXIS = len;
                get_actual_axis_value(odbaxdt, ref pos, len, types.Length);

                int i;
                //Servo load meter
                for (i = 0 * MAX_AXIS; i < 0 * MAX_AXIS + len; i++)
                {
                    buff += pos[i];
                    buff += ",";
                }
                buff += DL;

                // Load current (% unit)   
                for (i = 1 * MAX_AXIS; i < 1 * MAX_AXIS + len; i++)
                {
                    buff += pos[i];
                    buff += ",";
                }
                buff += DL;

                //Load current (Ampere unit) 
                for (i = 2 * MAX_AXIS; i < 2 * MAX_AXIS + len; i++)
                {
                    buff += pos[i];
                    buff += ",";
                    current += pos[i];
                    current += ",";
                }
            }
            return buff;
        }
        private static void get_actual_axis_value(FocasLibBase.ODBAXDT odbaxdt, ref List<string> pos, short len, int types)
        {
            pos.Add(FocasData.get_ax_val2(odbaxdt.data1.data, odbaxdt.data1.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data2.data, odbaxdt.data2.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data3.data, odbaxdt.data3.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data4.data, odbaxdt.data4.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data5.data, odbaxdt.data5.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data6.data, odbaxdt.data6.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data7.data, odbaxdt.data7.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data8.data, odbaxdt.data8.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data9.data, odbaxdt.data9.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data10.data, odbaxdt.data10.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data11.data, odbaxdt.data11.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data12.data, odbaxdt.data12.dec).ToString());
            if (types == 3 && len == 4)
                return;
            pos.Add(FocasData.get_ax_val2(odbaxdt.data13.data, odbaxdt.data13.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data14.data, odbaxdt.data14.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data15.data, odbaxdt.data15.dec).ToString());
            if (types == 3 && len == 5)
                return;
            pos.Add(FocasData.get_ax_val2(odbaxdt.data16.data, odbaxdt.data16.dec).ToString());
            if (types == 4 && len == 4)
                return;
            pos.Add(FocasData.get_ax_val2(odbaxdt.data17.data, odbaxdt.data17.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data18.data, odbaxdt.data18.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data19.data, odbaxdt.data19.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data20.data, odbaxdt.data20.dec).ToString());
            if (types == 4 && len == 5)
                return;
            pos.Add(FocasData.get_ax_val2(odbaxdt.data21.data, odbaxdt.data21.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data22.data, odbaxdt.data22.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data23.data, odbaxdt.data23.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data24.data, odbaxdt.data24.dec).ToString());
            pos.Add(FocasData.get_ax_val2(odbaxdt.data25.data, odbaxdt.data25.dec).ToString());
        }
        public static string get_ax_val2(long i, short d)
        {
            float pow_val = (float)Math.Pow(10, d);
            float ret_val = (float)(i / pow_val);
            return ret_val.ToString();
        }


        public static short WriteWearOffset2(ushort handle, short offsetno, decimal value)
        {
            short act_val = 0;
            if (offsetno == 0)
            {
                return -1;
            }
            act_val = WriteMacro(handle, offsetno, value);
            return act_val;
        }

        public static short WriteWearOffset_TEST(ushort handle, short offsetno, decimal value)
        {
            short act_val = 0;
            if (offsetno == 0)
            {
                return -1;
            }
          
            float valueF = Convert.ToSingle(value);
            //// float to int
            var y = BitConverter.ToInt32(BitConverter.GetBytes(valueF), 0);
            ////int to float
            //float d = BitConverter.ToSingle(BitConverter.GetBytes(y), 0);

            act_val = FocasLib.cnc_wrtofs(handle, offsetno, 0, 8, y);
            return act_val;
        }

        public static double ReadWearOffset2(ushort handle, short offsetno)
        {
            double act_val = 0;
            if (offsetno == 0)
            {
                return 0;
            }
            act_val = ReadMacroDouble2(handle, offsetno);
            return act_val;
        }

        public static void ReadMachineInfo(ushort handle,out string serialnumber,out string mtb)
        {
            serialnumber = "";
            mtb = "";
            FocasLibBase.ODBPMCTITLE a = new FocasLibBase.ODBPMCTITLE();
           int ret =  FocasLibrary.FocasLib.pmc_rdpmctitle(handle, a);
           if (ret == 0)
           {
               serialnumber = a.machine;
               mtb = a.mtb;
           }
        }

        public static string ReadCNCId(ushort handle)
        {
            string cncIDstr = "";
            uint[] cncid = new uint[4];
            int ret = FocasLibrary.FocasLib.cnc_rdcncid(handle, cncid);
            if (ret == 0)
            {
                foreach (var item in cncid)
                {
                    cncIDstr = cncIDstr + item.ToString("X") + "-";
                }
                //cncIDstr = string.Join("-", cncid);
            }
            return cncIDstr.Trim('-');
        }

        public static List<ServoLoad> ReadServoMotorLoad(ushort handle, short axisNo)
        {
            List<ServoLoad> load = new List<ServoLoad>();
            FocasLibBase.ODBSVLOAD a = new FocasLibBase.ODBSVLOAD();
            short ret = FocasLib.cnc_rdsvmeter(handle, ref axisNo, a);
            if (ret == 0)
            {
                if (a.svload1.name > 0)
                {
                    load.Add(new ServoLoad { Load = a.svload1.data, Axis = System.Text.Encoding.ASCII.GetString(new[] { a.svload1.name }), AxisNumber = 1 });
                }
                if (a.svload2.name > 0)
                {
                    load.Add(new ServoLoad { Load = a.svload2.data, Axis = System.Text.Encoding.ASCII.GetString(new[] { a.svload2.name }), AxisNumber = 2 });
                }
                if (a.svload3.name > 0)
                {
                    load.Add(new ServoLoad { Load = a.svload3.data, Axis = System.Text.Encoding.ASCII.GetString(new[] { a.svload3.name }), AxisNumber = 3 });
                }
                if (a.svload4.name > 0)
                {
                    load.Add(new ServoLoad { Load = a.svload4.data, Axis = System.Text.Encoding.ASCII.GetString(new[] { a.svload4.name }), AxisNumber = 4 });
                }
                if (a.svload5.name > 0)
                {
                    load.Add(new ServoLoad { Load = a.svload5.data, Axis = System.Text.Encoding.ASCII.GetString(new[] { a.svload5.name }), AxisNumber = 5 });
                }
            }
            return load;
        }

        public static int ReadServoMotorTemp(ushort handle, short axisNo)
        {
            FocasLibBase.ODBDGN_1 obj = new FocasLibBase.ODBDGN_1();
            short ret = FocasLib.cnc_diagnoss(handle, 308, axisNo, 8, obj);//temp spindle motor 408           
            return obj.ldata;
        }
    }
}


