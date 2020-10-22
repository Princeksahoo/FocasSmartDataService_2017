using System;
using System.Data;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.ComponentModel;
using System.Reflection;
using System.Linq;
using FocasLib;

namespace FocasLibrary
{
    public static class Utility
    {
        public enum CncMachineState
        {
            Available = 2,
            FeedHold = 6,
            Idle = 5,
            InCycle = 4,
            NewScan = 7,
            Offline = 1,
            Unavailable = 3,
            Unknown = -1
        }

        public static string GetMachineMode(FocasLibBase.ODBST odbst)
        {
            switch (odbst.aut)
            {
                case 0:
                    return "MDI";

                case 1:
                    return "MEM";

                case 3:
                    return "EDIT";

                case 4:
                    return "HANDLE";

                case 5:
                    return "JOG";

                case 6:
                    return "Teach in JOG";

                case 7:
                    return "Teach in HANDLE";

                case 8:
                    return "INC-feed";

                case 9:
                    return "Reference";

                case 10:
                    return "MEM";

                case 11:
                    return "TEST";
            }
            return ("Code: " + odbst.aut.ToString());
        }

        public static string GetMachineStatus(out Utility.CncMachineState _cncMachineState, FocasLibBase.ODBST odbst)
        {
            string str = string.Empty;
            _cncMachineState = Utility.CncMachineState.Unknown;
            return GetMachineStatusSeriesDefault(odbst, ref _cncMachineState);
        }

        private static string GetMachineStatusSeriesDefault(FocasLibBase.ODBST odbst, ref Utility.CncMachineState _cncMachineState)
        {
            string str = string.Empty;
            if ((odbst.emergency != 0))
            {
                _cncMachineState = Utility.CncMachineState.Unavailable;
                str = "Unavailable";
                if (odbst.emergency != 0)
                {
                    return "Emergency";
                }
                //if (odbst.alarm != 0)
                //{
                //    str = "Alarm";
                //}
                return str;
            }
            //if (odbst.aut == 1 || odbst.aut == 10)
            {
                switch (odbst.run)
                {
                    case 0:
                        _cncMachineState = Utility.CncMachineState.Idle;
                        return "Idle";

                    case 1:
                        _cncMachineState = Utility.CncMachineState.Unavailable;
                        return "STOP";

                    case 2:
                        _cncMachineState = Utility.CncMachineState.FeedHold;
                        return "Feed Hold";

                    case 3:
                        _cncMachineState = Utility.CncMachineState.InCycle;
                        return "In Cycle";

                    case 4:
                        _cncMachineState = Utility.CncMachineState.InCycle;
                        return "MSTR";
                }
                _cncMachineState = Utility.CncMachineState.Unknown;
                return ("Code: " + odbst.run.ToString());
            }
            _cncMachineState = Utility.CncMachineState.Unavailable;
            return "Unavailable";
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

        public static DateTime ConvertToDate(FocasLibBase.ALM_HIS5_data obj)
        {
            DateTime alarmTime = DateTime.Now;
            string str = string.Format("{0}-{1}-{2} {3}:{4}:{5}", obj.year, obj.month, obj.day, obj.hour, obj.minute, obj.second);
            if (!DateTime.TryParse(str, out alarmTime))
            {
                Logger.WriteErrorLog("Not able to parse date time string to date " + str);
            }

            //DateTime alarmTime = DateTime.MinValue;
            //string[] formats2 = { "yyyy-MM-dd HH:mm:ss"};
            //string str = string.Format("{0}-{1}-{2} {3}:{4}:{5}", obj.year, obj.month.ToString("00"), obj.day.ToString("00"), obj.hour.ToString("00"), obj.minute.ToString("00"), obj.second.ToString("00"));
            //if (!DateTime.TryParseExact(str, formats2, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out alarmTime))
            //{
            //    Logger.WriteErrorLog("Not able to parse date time string to date " + str);
            //}            
            return alarmTime;
        }
        public static DateTime ConvertToDate(FocasLibBase.ALM_HIS_data obj)
        {            
            DateTime alarmTime = DateTime.Now;
            string str = string.Format("{0}-{1}-{2} {3}:{4}:{5}", Convert.ToInt16(obj.year) < 2000 ? obj.year + 2000 : obj.year, obj.month, obj.day, obj.hour, obj.minute, obj.second);           
            if (! DateTime.TryParse(str, out alarmTime))
            {
                Logger.WriteErrorLog("Not able to parse date time string to date " + str);
            }
            return alarmTime;

            //DateTime alarmTime = DateTime.MinValue;
            //string[] formats2 = { "yyyy-MM-dd HH:mm:ss" };
            //string str = string.Format("{0}-{1}-{2} {3}:{4}:{5}", Convert.ToInt16(obj.year) < 2000 ? obj.year + 2000 : obj.year, obj.month.ToString("00"), obj.day.ToString("00"), obj.hour.ToString("00"), obj.minute.ToString("00"), obj.second.ToString("00"));
            //if (!DateTime.TryParseExact(str, formats2, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out alarmTime))
            //{
            //    Logger.WriteErrorLog("Not able to parse date time string to date " + str);
            //}
            //return alarmTime;
        }

        public static DataTable get_table()
        {
            DataTable dt = new DataTable("Alarms");

            dt.Columns.Add("AlarmNo", typeof(long));
            dt.Columns.Add("AlarmGroupNo", typeof(long));
            dt.Columns.Add("AlarmMSG", typeof(string));
            dt.Columns.Add("AlarmAxisNo", typeof(long));
            dt.Columns.Add("AlarmTotAxisNo", typeof(long));

            dt.Columns.Add("AlarmGCode", typeof(string));
            dt.Columns.Add("AlarmOtherCode", typeof(string));

            dt.Columns.Add("AlarmMPos", typeof(string));
            dt.Columns.Add("AlarmAPos", typeof(string));

            dt.Columns.Add("AlarmTime", typeof(DateTime));
            dt.Columns.Add("MachineID", typeof(string));
            return dt;
        }

        public static void get_datatable_row(FocasLibBase.ALM_HIS5_data obj, ref DataRow row)
        {
            try
            {
                row["AlarmNo"] = obj.alm_no;
                row["AlarmGroupNo"] = obj.alm_grp;
                row["AlarmMSG"] = obj.alm_msg;
                row["AlarmAxisNo"] = obj.axis_no;
                row["AlarmTotAxisNo"] = obj.axis_num;
                row["AlarmTime"] = Utility.ConvertToDate(obj);

                int count = 0;
                string alm_gcode = "";
                for (count = 0; count < 10; count++)
                {
                    alm_gcode += obj.g_modal[count].ToString();
                    alm_gcode += ",";
                    alm_gcode += obj.g_dp[count].ToString();
                    alm_gcode += ",";
                }
                row["AlarmGCode"] = alm_gcode;


                string alm_othCode = "";
                for (count = 0; count < 10; count++)
                {
                    alm_gcode += obj.a_modal[count].ToString();
                    alm_othCode += ",";
                    alm_gcode += obj.a_dp[count].ToString();
                    alm_othCode += ",";
                }
                row["AlarmOtherCode"] = alm_othCode;

                string str = "";
                for (count = 0; count < 31; count++)
                {
                    str += obj.abs_pos[count].ToString();
                    str += ",";
                    str += obj.abs_dp[count].ToString();
                    str += ",";

                }
                row["AlarmAPos"] = str;//absolute pos

                str = "";
                for (count = 0; count < 31; count++)
                {
                    str += obj.mcn_pos[count].ToString();
                    str += ",";
                    str += obj.mcn_dp[count].ToString();
                    str += ",";

                }
                row["AlarmMPos"] = str;//machine pos
            }
            catch (Exception ex)
            {

            }
        }

        public static void get_datatable_row(FocasLibBase.ALM_HIS_data obj, ref DataRow row)
        {
            try
            {
                row["AlarmNo"] = obj.alm_no;
                row["AlarmGroupNo"] = obj.alm_grp;
                row["AlarmMSG"] = obj.alm_msg;
                row["AlarmAxisNo"] = obj.axis_no;
                row["AlarmTotAxisNo"] = 0;
                row["AlarmTime"] = Utility.ConvertToDate(obj);

                row["AlarmGCode"] = string.Empty;
                row["AlarmOtherCode"] = string.Empty;
                row["AlarmAPos"] = string.Empty;
                row["AlarmMPos"] = string.Empty;
            }
            catch (Exception ex)
            {

            }
        }

        internal static DataTable get_OperationTable()
        {
            DataTable dt = new DataTable("OperationHistory");

            dt.Columns.Add("OperationType", typeof(string));
            dt.Columns.Add("OperationValue", typeof(string));
            dt.Columns.Add("ODateTime", typeof(DateTime));
            dt.Columns.Add("MachineID", typeof(string));
            dt.Columns.Add("TypeNumber", typeof(Int32));           
            return dt;
        }      

        public static DataTable ToDataTable<T>(this IList<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }

       
        private static DataTable GetDataTable<T>(this IList<T> data, IEnumerable<PropertyDescriptor> mappedProperties)
        {
            DataTable table = new DataTable();
            // columns
            foreach (PropertyDescriptor prop in mappedProperties)
            {
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            // row values
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in mappedProperties)
                {
                    object value = prop.GetValue(item) ?? DBNull.Value;
                    row[prop.Name] = value;
                }
                table.Rows.Add(row);
            }

            return table;
        }

        /* uses:
        properties to exclude from mapping
        var nonMappedProperties = new string[] { "P1", "P2" };
        Func<PropertyDescriptor, bool> expression = x => !nonMappedProperties.Contains(x.Name);
        */
        public static DataTable ToDataTable<T>(this IList<T> data, Func<PropertyDescriptor, bool> expression)
        {
            var properties = TypeDescriptor.GetProperties(typeof(T)).Cast<PropertyDescriptor>().Where(expression);

            DataTable table = GetDataTable(data, properties);
            return table;
        }
    }
}

