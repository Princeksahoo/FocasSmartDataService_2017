using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Data;
using System.Diagnostics;
using FocasLib;

namespace FocasLibrary
{
    public partial class FocasData
    {
        //TODO - complte it with datatable and database insert
        public static DataTable ReadOperationhistory18i(string machineID, string ipAddress, ushort portNo)
        {
            DataTable table = new DataTable();
            ushort handle = ushort.MinValue;
            int ret = 0;
            try
            {
                ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out handle);
                if (ret == 0)
                {
                    ret = FocasLib.cnc_stopophis(handle);
                    if (ret != 0)
                    {
                        Logger.WriteErrorLog("cnc_stopophis() failed. return value is = " + ret);                        
                    }
                    ushort totalHistoryRecords = 0;
                    ret = FocasLib.cnc_rdophisno(handle, out totalHistoryRecords);
                    if (ret != 0)
                    {
                        Logger.WriteErrorLog("cnc_rdophisno() failed. return value is = " + ret);
                    }
                    Logger.WriteDebugLog("Total number of rows found in Operation History = " + totalHistoryRecords);
                    if (totalHistoryRecords == 0) return table;
                    ushort returnRecordnumber = 0;
                    FocasLibBase.ODBHIS obj = new FocasLibBase.ODBHIS();
                    ushort objLength = (ushort)Marshal.SizeOf(obj);

                    int length = totalHistoryRecords;
                    returnRecordnumber = 1;
                    for (ushort i = 1; i <= length; i++)
                    {                        
                        returnRecordnumber = i;
                        obj = new FocasLibBase.ODBHIS();
                        ret = FocasLib.cnc_rdophistry(handle, i, returnRecordnumber, objLength, obj);                        
                        /*
                          rec_type  : Record type Description
                            0 : MDI key history 
                            1 : Signal history 
                            2 : Alarm history 
                            3 : Date history 
                            4 : Time history 
                            5 : MDI key history for SUB (only Series 160/180/210, Power Mate i) 
                            6 : Signal history for SUB (only Series 160/180/210, Power Mate i) 
                            7 : Alarm history for SUB (only Series 160/180/210, Power Mate i) 
                            10 : MDI key history for 3rd Path (only Series 160i) 
                            11 : Signal history for 3rd Path (only Series 160i) 
                            12 : Alarm history for 3rd Path (only Series 160i) 
                         */


                        if (obj.data.data1.rec_type == 3) //3 : Date history 
                        {
                            Debug.Print("YEAR = " + Convert.ToString(obj.data.data1.date_year.ToString()));
                            Debug.Print("MONTH = " + obj.data.data1.date_month.ToString());
                            Debug.Print("DAY = " + obj.data.data1.date_day.ToString());                           
                        }

                        if (obj.data.data1.rec_type == 4) //4 : Time history 
                        {
                            Debug.Print("hh = " + Convert.ToString(obj.data.data1.time_hour));
                            Debug.Print("MM = " + obj.data.data1.time_minute.ToString());
                            Debug.Print("SS = " + obj.data.data1.time_second.ToString());                           
                        }

                        if (obj.data.data1.rec_type == 1) // 1 : Signal history 
                        {
                            Debug.Print("sgn_sig_name = " + Convert.ToString(obj.data.data1.sgn_sig_name));
                            Debug.Print("sgn_sig_no = " + obj.data.data1.sgn_sig_no.ToString());
                            Debug.Print("nEW = " + obj.data.data1.sgn_sig_new.ToString());
                            Debug.Print("OLD " + obj.data.data1.sgn_sig_old.ToString());                            
                        }
                        if (obj.data.data1.rec_type == 2) // 2 : Alarm history 
                        {
                            Debug.Print("alm_alm_no = " + Convert.ToString(obj.data.data1.alm_alm_no));
                            Debug.Print("alm_alm_grp = " + obj.data.data1.alm_alm_grp);
                            Debug.Print("alm_axis_no = " + obj.data.data1.alm_axis_no);
                            Debug.Print("alm_dummy " + obj.data.data1.alm_dummy);                            
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
                if (handle != ushort.MinValue)
                {
                    FocasLib.cnc_startophis(handle);
                    FocasData.cnc_freelibhndl(handle);
                }
            }
            return table;
        }
        public static short ReadOperationhistory(ushort handle)
        {
            int ret = 0;
            try
            {

                FocasLib.cnc_stopophis(handle);

                ushort totalHistoryRecords = 0;
                FocasLib.cnc_rdophisno(handle,out totalHistoryRecords);


                ushort returnRecordnumber = 0;
                FocasLibBase.ODBOPHIS4_1 obj = new FocasLibBase.ODBOPHIS4_1();
                ushort objLength = (ushort)Marshal.SizeOf(obj);
                //ret = FocasLibHelper.cnc_rdophistry4(1, ref returnRecordnumber, ref objLength, obj);

                #region CommentedCode
                int length = totalHistoryRecords;
                returnRecordnumber = 1;
                for (ushort i = 1; i <= length; i++)
                {
                    objLength = (ushort)Marshal.SizeOf(obj);
                    returnRecordnumber = i;
                    ret = FocasLib.cnc_rdophistry4(handle,i, ref returnRecordnumber, ref objLength, obj);
                }
                #endregion
            }
            catch (Exception ex)
            {
                //TODO Log it
            }
            finally
            {
                FocasLib.cnc_startophis(handle);
            }
            return 0;
        }

        public static DataTable ReadOperationHistory(string machineId, string ipAddress, ushort portNo)
        {
            int ret = 0;
            ushort focasLibHandle = 0;
            DataTable operationHistoryTable = Utility.get_OperationTable();
            try
            {
                ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
                if (ret == 0)
                {
                    ret = FocasLib.cnc_stopophis(focasLibHandle);
                    ushort totalOperationHistoryRecords = 0;
                    ret = FocasLib.cnc_rdophisno(focasLibHandle, out totalOperationHistoryRecords);
                    if (totalOperationHistoryRecords == 0) return operationHistoryTable;
                    //Call C++ Dll to read the operation history details
                    //read one records at a time OR read a rage of records at a time OR Read all records at a time????
                    //C++ will write all the records to a file
                    //C# will process the file content and insert to database.

                }
            }
            catch (Exception ex)
            {
                //TODO
            }
            finally
            {
                if (focasLibHandle > 0)
                {
                    ret = FocasLib.cnc_startophis(focasLibHandle);
                }
            }

            return operationHistoryTable;
        }

        public static int ReadOperationHistoryALLCppDLL(string machineId, string ipAddress, ushort portNo, string filePath)
        {
            int ret = 0;
            //Call C++ Dll method, which will fetch data from CNC and write to a file, C# will process those files.

            return ret;
        }

        public static string ReadOperationHistoryRangeCppDLL(string machineId, string ipAddress, ushort portNo, string filePath)
        {
            int ret = 0;
            StringBuilder sb = new StringBuilder();
            ushort focasLibHandle = 0;
            //Call C++ Dll method, which will fetch range of data from CNC and return the data as string to C#
            //C# will parse the return string and create the DataTable for BulkCopy


            try
            {

                ret = FocasLib.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
                if (ret == 0)
                {
                    FocasLib.cnc_stopophis(focasLibHandle);

                    ushort totalHistoryRecords = 0;
                    FocasLib.cnc_rdophisno(focasLibHandle, out totalHistoryRecords);

                    for (ushort i = 1; i <= totalHistoryRecords; i++)
                    {
                        //TODO - Call C++ DLL function and appent the result to stringBuilder.
                        sb.Append("");
                    }
                }
            }
            catch (Exception ex)
            {
                //TODO Log it
            }
            finally
            {
                FocasLib.cnc_startophis(focasLibHandle);
            }
            return sb.ToString();
        }
    }
}
