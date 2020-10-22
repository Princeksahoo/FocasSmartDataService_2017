using System;
using System.Collections.Generic;
using System.Text;
using FocasLib;

namespace FocasLibrary
{
    public static partial class FocasData
    {
        public static int ReadMacro(ushort handle, short macroLocation)
        {
            FocasLibBase.ODBM odbm = new FocasLibBase.ODBM();
            int ret = FocasLib.cnc_rdmacro(handle, macroLocation, 10, odbm);
            if (ret == 0)
            {
                if (odbm.mcr_val == 0 && odbm.dec_val == -1)
                {
                    return -1;
                }
                else
                {
                    return (int)(odbm.mcr_val / Math.Pow(10.0, odbm.dec_val));
                }
            }
            else
            {
                Logger.WriteErrorLog("cnc_rdmacro() failed. return value is = " + ret);
            }
            return -1;
        }      

        public static double ReadMacroDouble(ushort handle, short macroLocation)
        {
            FocasLibBase.ODBM odbm = new FocasLibBase.ODBM();
            int ret = FocasLib.cnc_rdmacro(handle, macroLocation, 10, odbm);
            if (ret == 0)
            {
                if (odbm.mcr_val == 0 && odbm.dec_val == -1)
                {
                    return -1;
                }
                else
                {
                    return odbm.mcr_val / Math.Pow(10.0, odbm.dec_val);
                }
            }
            else
            {
                Logger.WriteErrorLog("cnc_rdmacro() failed. return value is = " + ret);
            }
            return -1;
        }

        public static double ReadMacroDouble2(ushort handle, short macroLocation)
        {
            FocasLibBase.ODBM odbm = new FocasLibBase.ODBM();
            int ret = FocasLib.cnc_rdmacro(handle, macroLocation, 10, odbm);
            if (ret == 0)
            {
                if (odbm.mcr_val == 0 && odbm.dec_val == -1)
                {
                    return double.MaxValue;
                }
                else
                {
                    return odbm.mcr_val / Math.Pow(10.0, odbm.dec_val);
                }
            }
            else
            {
                Logger.WriteErrorLog("cnc_rdmacro() failed. return value is = " + ret);
            }
            return double.MaxValue;
        }

        public static short WriteMacro(ushort handle, short macroLocation, int value)
        {
            return FocasLib.cnc_wrmacro(handle, macroLocation, 10, value, 0);
        }

        public static short WriteMacro(ushort handle, short macroLocation, decimal value)
        {
            short decimalPlaces = BitConverter.GetBytes(decimal.GetBits(value)[3])[2];
            short intValue = Convert.ToInt16(value * (decimal)Math.Pow(10, decimalPlaces));
            return FocasLib.cnc_wrmacro(handle, macroLocation, 10, intValue, decimalPlaces);
        }  

        /// <summary>
        /// currently only supports max 10 macro location read at a time
        /// </summary>
        /// <param name="startLoc"></param>
        /// <param name="EndLoc"></param>
        /// <returns>List of macro values</returns>
        public static List<int> ReadMacroRange(ushort handle, short startLoc, short EndLoc)
        {
            List<int> values = new List<int>();            
            FocasLibBase.IODBMR odbmr = new FocasLibBase.IODBMR();
            //int ret = FocasLib.cnc_rdmacroRange(711, 720, 8 + (8 * 10), odbmr);
            int ret = FocasLib.cnc_rdmacror(handle, startLoc, EndLoc, 8 + (8 * 14), odbmr);
            if (ret == 0)
            {
                values.Add(GetMacroValue(odbmr.data.data1));
                values.Add(GetMacroValue(odbmr.data.data2));
                values.Add(GetMacroValue(odbmr.data.data3));
                values.Add(GetMacroValue(odbmr.data.data4));
                values.Add(GetMacroValue(odbmr.data.data5));
                values.Add(GetMacroValue(odbmr.data.data6));
                values.Add(GetMacroValue(odbmr.data.data7));
                values.Add(GetMacroValue(odbmr.data.data8));
                values.Add(GetMacroValue(odbmr.data.data9));
                values.Add(GetMacroValue(odbmr.data.data10));
                values.Add(GetMacroValue(odbmr.data.data11));
                values.Add(GetMacroValue(odbmr.data.data12));
                values.Add(GetMacroValue(odbmr.data.data13));
                values.Add(GetMacroValue(odbmr.data.data14));
            }
            else
            {
                Logger.WriteErrorLog("cnc_rdmacror() failed. return value is = " + ret);
            }
            return values;
        }

        public static int GetMacroValue(FocasLibBase.IODBMR_data data)
        {
            if (data.mcr_val == 0 && data.dec_val == -1)
            {
                return 0;
            }
            else
            {
                return (int)(data.mcr_val / Math.Pow(10.0, data.dec_val));
            }
        }

        public static short ReadParametershort(ushort handle, short parameter)
        {
            short partsCount = 0;
            FocasLibBase.IODBPSD_1 para = new FocasLibBase.IODBPSD_1();
            short ret = FocasLib.cnc_rdparam(handle, parameter, 0, 8, para);
            if (ret == 0)
            {
                partsCount = para.idata;
            }
            else
            {
                Logger.WriteErrorLog(string.Format("Parameter : {0} - cnc_rdparam() failed. return value is = {1}", parameter, ret));
            }
            
            return partsCount;
        }

        public static int ReadParameterInt(ushort handle, short parameter)
        {
            int partsCount = 0;
            FocasLibBase.IODBPSD_1 para = new FocasLibBase.IODBPSD_1();
            short ret = FocasLib.cnc_rdparam(handle, parameter, 0, 8, para);
            if (ret == 0)
            {
                partsCount = para.ldata;
            }
            else
            {
                Logger.WriteErrorLog(string.Format("Parameter : {0} - cnc_rdparam() failed. return value is = {1}", parameter,ret));
            }
            return partsCount;
        }

        /// <summary>
        /// currently only supports max 10 macro location read at a time
        /// </summary>
        /// <param name="startLoc"></param>
        /// <param name="EndLoc"></param>
        /// <returns>List of macro values</returns>
        public static List<double> ReadMacroRangeDouble(ushort handle, short startLoc, short EndLoc)
        {
            List<double> values = new List<double>();
            FocasLibBase.IODBMR odbmr = new FocasLibBase.IODBMR();
            //int ret = FocasLib.cnc_rdmacroRange(711, 720, 8 + (8 * 10), odbmr);
            int ret = FocasLib.cnc_rdmacror(handle, startLoc, EndLoc, 8 + (8 * 14), odbmr);
            if (ret == 0)
            {
                values.Add(GetMacroValueDouble(odbmr.data.data1));
                values.Add(GetMacroValueDouble(odbmr.data.data2));
                values.Add(GetMacroValueDouble(odbmr.data.data3));
                values.Add(GetMacroValueDouble(odbmr.data.data4));
                values.Add(GetMacroValueDouble(odbmr.data.data5));
                values.Add(GetMacroValueDouble(odbmr.data.data6));
                values.Add(GetMacroValueDouble(odbmr.data.data7));
                values.Add(GetMacroValueDouble(odbmr.data.data8));
                values.Add(GetMacroValueDouble(odbmr.data.data9));
                values.Add(GetMacroValueDouble(odbmr.data.data10));
                values.Add(GetMacroValueDouble(odbmr.data.data11));
                values.Add(GetMacroValueDouble(odbmr.data.data12));               
            }
            else
            {
                Logger.WriteErrorLog("cnc_rdmacror() failed. return value is = " + ret);
            }
            return values;
        }

        public static double GetMacroValueDouble(FocasLibBase.IODBMR_data data)
        {
            if (data.mcr_val == 0 && data.dec_val == -1)
            {
                return double.MaxValue;
            }
            else
            {
                return data.mcr_val / Math.Pow(10.0, data.dec_val);
            }            
        }

    }
}
