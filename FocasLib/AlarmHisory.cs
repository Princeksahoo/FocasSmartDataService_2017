using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Data;
using FocasLib;
using DTO;

namespace FocasLibrary
{
    public static partial class FocasData
    {        
       public static AlarmHistoryDTO ReadCurrentAlarm(ushort handle)
        {
            int ret = 0;
            AlarmHistoryDTO latestAlarm = new AlarmHistoryDTO();
            try
            {
                ret = FocasLib.cnc_stopophis(handle);

                ushort totalAlarms = 0;
                FocasLibBase.ODBAHIS5 obj = new FocasLibBase.ODBAHIS5();
                ret = FocasLib.cnc_rdalmhisno(handle, out totalAlarms);

                ret = FocasLib.cnc_rdalmhistry5(handle, totalAlarms, totalAlarms, (ushort)Marshal.SizeOf(obj), obj);
                if (ret == 0 && obj != null)
                {
                    latestAlarm.AlarmNo = obj.alm_his.data1.alm_no;
                    latestAlarm.AlarmGroupNo = obj.alm_his.data1.alm_grp;
                    latestAlarm.AlarmMessage = obj.alm_his.data1.alm_msg;
                    latestAlarm.AlarmTime = Utility.ConvertToDate(obj.alm_his.data1);
                }
            }
            catch (Exception ex)
            {
            }
            finally
            {
                ret = FocasLib.cnc_startophis(handle);
            }
            return latestAlarm;
        }

        public static DataTable ReadAlarmHistory(string machineID, string ipAddress, ushort portNo)
        {
            int ret = 0;
            ushort focasLibHandle = 0;
            DataTable alarms = Utility.get_table();
            try
            {

                ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
                if (ret == 0)
                {
                    ret = FocasLib.cnc_stopophis(focasLibHandle);
                    if (ret != 0)
                    {
                        Logger.WriteErrorLog("cnc_stopophis() failed. return value is = " + ret);
                    }
                    ushort totalAlarms = 0;
                    ret = FocasLib.cnc_rdalmhisno(focasLibHandle, out totalAlarms);
                    if (ret != 0)
                    {
                        Logger.WriteErrorLog("cnc_rdalmhisno() failed. return value is = " + ret);
                    }
                    if (totalAlarms == 0) return alarms;

                    //10 rows at a time               
                    DataRow row = default(DataRow);
                    FocasLibBase.ODBAHIS5 obj = new FocasLibBase.ODBAHIS5();
                    List<FocasLibBase.ALM_HIS5_data> focasAlarmsObj = new List<FocasLibBase.ALM_HIS5_data>();//almhisdata object list

                    ushort loop_count = (ushort)(totalAlarms / 10);
                    ushort remainder = (ushort)(totalAlarms % 10);
                    ushort i = 0, s_no = 0, e_no = 0;
                    //ushort length = 4 + 512 * 10;
                    ushort length = (ushort)Marshal.SizeOf(obj);
                    for (i = 0; i <= loop_count; i++)
                    {
                        if (i == loop_count)
                        {
                            if (remainder == 0) continue;
                            s_no = (ushort)(i * 10 + 1);
                            e_no = (ushort)(s_no + (remainder - 1));
                            ret = FocasLib.cnc_rdalmhistry5(focasLibHandle, s_no, e_no, length, obj);
                            if (ret != 0)
                            {
                                Logger.WriteErrorLog("cnc_rdalmhistry5() failed. return value is = " + ret);
                            }
                            focasAlarmsObj.Clear();
                            get_ahd_objects(ref focasAlarmsObj, obj);
                            for (int j = 0; j < remainder; j++)
                            {
                                row = alarms.NewRow();
                                Utility.get_datatable_row(focasAlarmsObj[j], ref row);
                                row["MachineID"] = machineID;
                                alarms.Rows.Add(row);
                            }
                        }
                        else
                        {
                            s_no = (ushort)(i * 10 + 1);
                            e_no = (ushort)(s_no + 9);
                            ret = FocasLib.cnc_rdalmhistry5(focasLibHandle, s_no, e_no, length, obj);
                            if (ret != 0)
                            {
                                Logger.WriteErrorLog("cnc_rdalmhistry5() failed. return value is = " + ret);
                            }
                            focasAlarmsObj.Clear();
                            get_ahd_objects(ref focasAlarmsObj, obj);//
                            for (int j = 0; j < 10; j++)
                            {
                                row = alarms.NewRow();
                                Utility.get_datatable_row(focasAlarmsObj[j], ref row);
                                row["MachineID"] = machineID;
                                alarms.Rows.Add(row);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (focasLibHandle > 0)
                {
                    ret = FocasLib.cnc_startophis(focasLibHandle);
                    FocasData.cnc_freelibhndl(focasLibHandle);
                }
            }

            return alarms;
        }

        public static DataTable ReadAlarmHistory18i(string machineID, string ipAddress, ushort portNo)
        {
            int ret = 0;
            ushort focasLibHandle = 0;
            DataTable alarms = Utility.get_table();
            try
            {

                ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
                if (ret == 0)
                {
                    ret = FocasLib.cnc_stopophis(focasLibHandle);
                    if (ret != 0)
                    {
                        Logger.WriteErrorLog("cnc_stopophis() failed. return value is = " + ret);
                    }
                    ushort totalAlarms = 0;
                    ret = FocasLib.cnc_rdalmhisno(focasLibHandle, out totalAlarms);
                    if (ret != 0)
                    {
                        Logger.WriteErrorLog("cnc_rdalmhisno() failed. return value is = " + ret);
                    }
                    if (totalAlarms == 0) return alarms;

                    //10 rows at a time               
                    DataRow row = default(DataRow);                   
                    FocasLibBase.ODBAHIS obj18i = new FocasLibBase.ODBAHIS();
                    List<FocasLibBase.ALM_HIS_data> focasAlarmsObjobj18i = new List<FocasLibBase.ALM_HIS_data>();

                    ushort loop_count = (ushort)(totalAlarms / 10);
                    ushort remainder = (ushort)(totalAlarms % 10);
                    ushort i = 0, s_no = 0, e_no = 0;                    
                    ushort length18i = (ushort)Marshal.SizeOf(obj18i);
                    for (i = 0; i <= loop_count; i++)
                    {                       
                        focasAlarmsObjobj18i.Clear();
                        if (i == loop_count)
                        {
                            if (remainder == 0) continue;
                            s_no = (ushort)(i * 10 + 1);
                            e_no = (ushort)(s_no + (remainder - 1));
                            ret = FocasLib.cnc_rdalmhistry(focasLibHandle, s_no, e_no, length18i, obj18i);
                            if (ret != 0)
                            {
                                Logger.WriteErrorLog("cnc_rdalmhistry() failed. return value is = " + ret);
                            }
                            get_ahd_objects(ref focasAlarmsObjobj18i, obj18i);
                            for (int j = 0; j < remainder; j++)
                            {
                                row = alarms.NewRow();
                                Utility.get_datatable_row(focasAlarmsObjobj18i[j], ref row);
                                row["MachineID"] = machineID;
                                alarms.Rows.Add(row);
                            }
                        }
                        else
                        {
                            s_no = (ushort)(i * 10 + 1);
                            e_no = (ushort)(s_no + 9);
                            ret = FocasLib.cnc_rdalmhistry(focasLibHandle, s_no, e_no, length18i, obj18i);
                            if (ret != 0)
                            {
                                Logger.WriteErrorLog("cnc_rdalmhistry() failed. return value is = " + ret);
                            }
                            get_ahd_objects(ref focasAlarmsObjobj18i, obj18i);
                            for (int j = 0; j < 10; j++)
                            {
                                row = alarms.NewRow();
                                Utility.get_datatable_row(focasAlarmsObjobj18i[j], ref row);
                                row["MachineID"] = machineID;
                                alarms.Rows.Add(row);
                            }                                                    
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (focasLibHandle > 0)
                {
                    ret = FocasLib.cnc_startophis(focasLibHandle);
                    FocasData.cnc_freelibhndl(focasLibHandle);
                }
            }
            return alarms;
        }

        private static void get_ahd_objects(ref List<FocasLibBase.ALM_HIS5_data> ahd_objects, FocasLibBase.ODBAHIS5 obj2)
        {
            try
            {
                ahd_objects.Add(obj2.alm_his.data1);
                ahd_objects.Add(obj2.alm_his.data2);
                ahd_objects.Add(obj2.alm_his.data3);
                ahd_objects.Add(obj2.alm_his.data4);
                ahd_objects.Add(obj2.alm_his.data5);
                ahd_objects.Add(obj2.alm_his.data6);
                ahd_objects.Add(obj2.alm_his.data7);
                ahd_objects.Add(obj2.alm_his.data8);
                ahd_objects.Add(obj2.alm_his.data9);
                ahd_objects.Add(obj2.alm_his.data10);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
        }

        private static void get_ahd_objects(ref List<FocasLibBase.ALM_HIS_data> ahd_objects, FocasLibBase.ODBAHIS obj2)
        {
            try
            {
                ahd_objects.Add(obj2.alm_his.data1);
                ahd_objects.Add(obj2.alm_his.data2);
                ahd_objects.Add(obj2.alm_his.data3);
                ahd_objects.Add(obj2.alm_his.data4);
                ahd_objects.Add(obj2.alm_his.data5);
                ahd_objects.Add(obj2.alm_his.data6);
                ahd_objects.Add(obj2.alm_his.data7);
                ahd_objects.Add(obj2.alm_his.data8);
                ahd_objects.Add(obj2.alm_his.data9);
                ahd_objects.Add(obj2.alm_his.data10);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
        }

        public static List<LiveAlarm> ReadLiveAlarms(ushort handle)
        {
            List<LiveAlarm> alarms = new List<LiveAlarm>();
            FocasLibBase.ODBALMMSG obj = new FocasLibBase.ODBALMMSG();
            short num = 10;
            short ret = FocasLib.cnc_rdalmmsg(handle, -1, ref num, obj);
            if (ret == 0)
            {
                if (obj.msg1.alm_no > 0)
                    alarms.Add(AssignLiveAlarm(obj.msg1));
                if (obj.msg2.alm_no > 0)
                    alarms.Add(AssignLiveAlarm(obj.msg2));
                if (obj.msg3.alm_no > 0)
                    alarms.Add(AssignLiveAlarm(obj.msg3));
                if (obj.msg4.alm_no > 0)
                    alarms.Add(AssignLiveAlarm(obj.msg4));
                if (obj.msg5.alm_no > 0)
                    alarms.Add(AssignLiveAlarm(obj.msg5));
                if (obj.msg6.alm_no > 0)
                    alarms.Add(AssignLiveAlarm(obj.msg6));
                if (obj.msg7.alm_no > 0)
                    alarms.Add(AssignLiveAlarm(obj.msg7));
                if (obj.msg8.alm_no > 0)
                    alarms.Add(AssignLiveAlarm(obj.msg8));
                if (obj.msg9.alm_no > 0)
                    alarms.Add(AssignLiveAlarm(obj.msg9));
                if (obj.msg10.alm_no > 0)
                    alarms.Add(AssignLiveAlarm(obj.msg10));
            }
            return alarms;
        }

        public static LiveAlarm AssignLiveAlarm(FocasLibBase.ODBALMMSG_data data)
        {
            LiveAlarm alarm = new LiveAlarm();
            alarm.AlarmNo = data.alm_no;
            alarm.Message = data.alm_msg;            
            return alarm;
        }
      
    }
}
