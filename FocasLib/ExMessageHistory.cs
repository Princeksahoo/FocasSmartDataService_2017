using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Linq;
using FocasLib;
using System.Runtime.InteropServices;
using DTO;

namespace FocasLibrary
{
    public static partial class FocasData
    {
        //CNC parameter '3112#2' must be 1. 
        public static List<OprMessageDTO> ReadExternalOperatorMessageHistory0i(string machineID, string ipAddress, ushort portNo)
        {
            int ret = 0;
            ushort focasLibHandle = 0;
            List<OprMessageDTO> oprMessagesObj = new List<OprMessageDTO>();
            try
            {
                ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_allclibhndl3() failed. return value is = " + ret);
                    return null;
                }

                ret = FocasLib.cnc_stopophis(focasLibHandle);
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_stopophis() failed. return value is = " + ret);
                    return null;
                }
                ushort totalMessages = 0;
                ret = FocasLib.cnc_rdomhisno(focasLibHandle, out totalMessages);
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_rdomhisno() failed. return value is = " + ret);
                    return null;
                }
                if (totalMessages == 0) return oprMessagesObj;

                //10 rows at a time               
               
                ushort loop_count = (ushort)(totalMessages / 10);
                ushort remainder = (ushort)(totalMessages % 10);
                ushort i = 0, s_no = 0, e_no = 0;
                //ushort length = 4 + 512 * 10;
               
                for (i = 0; i <= loop_count; i++)
                {
                    if (i == loop_count)
                    {
                        if (remainder > 0)
                        {
                            s_no = (ushort)(i * 10 + 1);
                            e_no = (ushort)(s_no + (remainder - 1));
                            FocasLibBase.ODBOMHIS2 obj = new FocasLibBase.ODBOMHIS2();
                            ushort length = (ushort)Marshal.SizeOf(obj);
                            ret = FocasLib.cnc_rdomhistry2(focasLibHandle, s_no, e_no, length, obj);                           
                            if (ret != 0)
                            {
                                Logger.WriteErrorLog("cnc_rdomhistry2() failed. return value is = " + ret);
                                continue;
                            }
                            AddMessagesToList(obj, oprMessagesObj);
                            try
                            {
                                oprMessagesObj.RemoveRange(totalMessages, oprMessagesObj.Count - totalMessages);
                            }
                            catch (Exception exxx)
                            {
                                Logger.WriteErrorLog(exxx.ToString());
                            }
                        }
                    }
                    else
                    {
                        s_no = (ushort)(i * 10 + 1);
                        e_no = (ushort)(s_no + 9);
                        FocasLibBase.ODBOMHIS2 obj = new FocasLibBase.ODBOMHIS2();
                        ushort length = (ushort)Marshal.SizeOf(obj);
                        ret = FocasLib.cnc_rdomhistry2(focasLibHandle, s_no, e_no, length, obj);
                        if (ret != 0)
                        {
                            Logger.WriteErrorLog("cnc_rdomhistry2() failed. return value is = " + ret);
                            continue;
                        }
                        AddMessagesToList(obj, oprMessagesObj);
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
            foreach (OprMessageDTO item in oprMessagesObj)
            {
                item.MachineID = machineID;                
            }
            return oprMessagesObj;
        }

        private static void AddMessagesToList(FocasLibBase.ODBOMHIS2 obj, List<OprMessageDTO> oprMessagesObj)
        {
           oprMessagesObj.Add(ConvertToMessage(obj.opm_his.data1));
           oprMessagesObj.Add(ConvertToMessage(obj.opm_his.data2));
           oprMessagesObj.Add(ConvertToMessage(obj.opm_his.data3));
           oprMessagesObj.Add(ConvertToMessage(obj.opm_his.data4));
           oprMessagesObj.Add(ConvertToMessage(obj.opm_his.data5));
           oprMessagesObj.Add(ConvertToMessage(obj.opm_his.data6));
           oprMessagesObj.Add(ConvertToMessage(obj.opm_his.data7));
           oprMessagesObj.Add(ConvertToMessage(obj.opm_his.data8));
           oprMessagesObj.Add(ConvertToMessage(obj.opm_his.data9));
           oprMessagesObj.Add(ConvertToMessage(obj.opm_his.data10));
        }

        private static OprMessageDTO ConvertToMessage(FocasLibBase.ODBOMHIS2_data data)
        {
            OprMessageDTO msg = new OprMessageDTO();
            msg.MessageNo = data.om_no;
            msg.Message = data.alm_msg;
            msg.MessageDate = new DateTime(data.year < 2000 ? data.year + 2000 : data.year, data.month, data.day, data.hour, data.minute, data.second);
            return msg;
        }

        private static void AddMessagesToList(FocasLibBase.ODBOMHIS obj, List<OprMessageDTO> oprMessagesObj)
        {
           oprMessagesObj.Add(ConvertToMessage(obj.omhis1));
           oprMessagesObj.Add(ConvertToMessage(obj.omhis2));
           oprMessagesObj.Add(ConvertToMessage(obj.omhis3));
           oprMessagesObj.Add(ConvertToMessage(obj.omhis4));
           oprMessagesObj.Add(ConvertToMessage(obj.omhis5));
           oprMessagesObj.Add(ConvertToMessage(obj.omhis6));
           oprMessagesObj.Add(ConvertToMessage(obj.omhis7));
           oprMessagesObj.Add(ConvertToMessage(obj.omhis8));
           oprMessagesObj.Add(ConvertToMessage(obj.omhis9));
           oprMessagesObj.Add(ConvertToMessage(obj.omhis10));
        }

        private static OprMessageDTO ConvertToMessage(FocasLibBase.ODBOMHIS_data data)
        {
            OprMessageDTO msg = new OprMessageDTO();
            msg.MessageNo = data.om_no;
            msg.Message = data.om_msg;
            msg.MessageDate = new DateTime(data.year < 2000 ? data.year + 2000 : data.year, data.month, data.day, data.hour, data.minute, data.second);
            return msg;
        }

        /*(9) CNC parameter error CNC parameter '3112#2' must be 1. */
        public static List<OprMessageDTO> ReadExternalOperatorMessageHistory18i(string machineID, string ipAddress, ushort portNo)
        {
            int ret = 0;
            ushort focasLibHandle = 0;
            List<OprMessageDTO> oprMessagesObj = new List<OprMessageDTO>();
            try
            {
                ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_allclibhndl3() failed. return value is = " + ret);
                    return null;
                }

                ret = FocasLib.cnc_stopomhis(focasLibHandle);
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_stopomhis() failed. return value is = " + ret);
                    return null;
                }
                FocasLibBase.ODBOMIF a = new FocasLibBase.ODBOMIF();
                ret = FocasLib.cnc_rdomhisinfo(focasLibHandle, a);
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_rdomhisinfo() failed. return value is = " + ret);
                    return null;
                }
                if (a.om_max == 0) return oprMessagesObj;
                //TODO - check it?
                int totalMessages = a.om_max;
                //10 rows at a time
                ushort loop_count = (ushort)(totalMessages / 10);
                ushort remainder = (ushort)(totalMessages % 10);
                ushort i = 0, s_no = 0, e_no = 0;
                //ushort length = 4 + 512 * 10;

                for (i = 0; i <= loop_count; i++)
                {
                    if (i == loop_count)
                    {
                        if (remainder > 0)
                        {
                            s_no = (ushort)(i * 10 + 1);
                            e_no = (ushort)(s_no + (remainder - 1));
                            FocasLibBase.ODBOMHIS obj = new FocasLibBase.ODBOMHIS();
                            ushort length = (ushort)Marshal.SizeOf(obj);
                            ret = FocasLib.cnc_rdomhistry(focasLibHandle, s_no, ref e_no, obj);
                            if (ret != 0)
                            {
                                Logger.WriteErrorLog("cnc_rdomhistry() failed. return value is = " + ret);
                                continue;
                            }                           
                            AddMessagesToList(obj, oprMessagesObj);
                            try
                            {
                                oprMessagesObj.RemoveRange(totalMessages, oprMessagesObj.Count - totalMessages);
                            }
                            catch (Exception exxx)
                            {
                                Logger.WriteErrorLog(exxx.ToString());
                            }
                        }
                    }
                    else
                    {
                        s_no = (ushort)(i * 10 + 1);
                        e_no = (ushort)(s_no + 9);
                        FocasLibBase.ODBOMHIS obj = new FocasLibBase.ODBOMHIS();
                        ushort length = (ushort)Marshal.SizeOf(obj);
                        ret = FocasLib.cnc_rdomhistry(focasLibHandle, s_no, ref e_no, obj);
                        if (ret != 0)
                        {
                            Logger.WriteErrorLog("cnc_rdomhistry() failed. return value is = " + ret);
                            continue;
                        }
                        //TODO - add the message based on s_no and e_no
                        AddMessagesToList(obj, oprMessagesObj);
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

            return oprMessagesObj;
        }
               
    }
}
