using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text;
using System.Threading;
using FocasLib;
using FocasLibrary;
using System.Linq;
using DTO;
using System.Collections.Concurrent;
using System.Net.Sockets;
using System.Threading.Tasks;
using System.Collections;
using System.Text.RegularExpressions;
using MachineConnectLicenseDTO;
using TPMTrakSmartDataServiceForMazak;

namespace FocasSmartDataCollection
{
    public class CreateClient
    {
        private string ipAddress;
        private ushort portNo;
        private string machineId;
        private string interfaceId;
        private string MName;
        private string ControllerType;
        private int machineType;
        private int systemType;
        private int pingFailureCount = 0;
        private short AddressPartSCountFromMacro = 0;      
        private short _CompMacroLocation = 0;
        private short _OpnMacroLocation = 0;
        bool _isMacroStringEnabled = false;
        string _toolLifeStartLocation = string.Empty;

        private short InspectionDataReadFlag = 0;
        private bool enableSMSforProgramChange = false;
        public string MachineName
        {
            get { return machineId; }
        }
        MachineSetting setting = default(MachineSetting);
        MachineInfoDTO machineDTO = default(MachineInfoDTO);
        private static string appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        CncMachineType _cncMachineType = CncMachineType.cncUnknown;
        string _cncSeries = string.Empty;

        //LIC details
        bool _isCNCIdReadSuccessfully = false;//**********************************TODO - change to false before check-in
        private bool _isLicenseValid = false;

        private List<SpindleSpeedLoadDTO> _spindleInfoQueue = new List<SpindleSpeedLoadDTO>();
        private List<LiveDTO> _liveDTOQueue = new List<LiveDTO>();      
        private string _operationHistoryFolderPath = string.Empty;
        private double _timeDelayMainThread = 0;
        private string _programDownloadFolder = string.Empty;
        private int _DownloadFreq = 60;
        private Hashtable _AutoDownloadedSavedPrograms = new Hashtable();
        private bool _AutoDownloadEveryTimeIfNotSameAsMaster = true;
        private static DateTime _serviceStartedTimeStamp = DateTime.Now;
        private static DateTime _nextLicCheckedTimeStamp = _serviceStartedTimeStamp;
        bool _prevAlarmStatus = false;
        short cycleTimeMacroLocation = 0;
       

        private Timer _timerAlarmHistory = null;
        private Timer _timerOperationHistory = null;
        private Timer _timerSpindleLoadSpeed = null;
        private Timer _timerTPMTrakDataCollection = null;
        private Timer _timerPredictiveMaintenanceReader = null;
        private Timer _timerOffsetHistoryReader = null;
        private Timer _timerCycletimeReader = null;
        private Timer _timerToolLife = null;
        private Timer _timerProcessParameter = null;

        object _lockerAlarmHistory = new object();
        object _lockerOperationHistory = new object();       
        object _lockerTPMTrakDataCollection = new object();
        object _lockerOffsetCorrectionDataCollection = new object();
        object _lockerSpindleLoadSpeed = new object();
        object _lockerPredictiveMaintenance = new object();
        object _lockerOffsetHistory = new object();
        object _lockerCycletimeReader = new object();
        object _lockerProcessParameter = new object();
        object _lockerToolLife = new object();

        //static volatile object _lockerReleaseMemory = new object();
     
        //List<ushort> _focasHandles = new List<ushort>();
        bool _isOEMVersion = false;
        List<OffsetHistoryDTO> offsetHistoryList = new List<OffsetHistoryDTO>();
        List<LiveAlarm> _liveAlarmsGlobal = new List<LiveAlarm>();
        List<int> offsetHistoryRange = new List<int>();
        List<LiveAlarm> liveAlarmsLocal = new List<LiveAlarm>();
        short _ComponentMacro = 600;
        short _CycleEndDateTimeMacro = 606;
        short _DataReadflagMacro = 530;
        short _featureIDMacro = 539;
        short _DataReadStatusFlagMacro = 608;
        short _ComponentGroupIdMacro = 608;

        public CreateClient(MachineInfoDTO machine)
        {            
            this.ipAddress = machine.IpAddress;
            this.portNo = (ushort)machine.PortNo;
            this.machineId = machine.MachineId;
            this.MName = this.machineId;
            this.interfaceId = machine.InterfaceId;
            this.setting = machine.Settings;
            this.ControllerType = machine.ControllerType;
            this.systemType = machine.SystemType;
            this.machineType = machine.MachineType;
            this.machineDTO = machine;
            
            var appSettings = DatabaseAccess.GetServiceSettingsData();

            _operationHistoryFolderPath = appSettings.OperationHistoryFileDownloadPath;
            AddressPartSCountFromMacro = this.machineDTO.Settings.PartsCountUsingMacro;
            _programDownloadFolder = appSettings.ProgramDownloadPath;

        }

        public void GetClient()
        {
            Logger.WriteDebugLog(string.Format("Thread {0} started for data collection.", machineId));
            if (this.ControllerType.Equals("MITSUBISHI", StringComparison.OrdinalIgnoreCase))
            {
                short.TryParse(ConfigurationManager.AppSettings["Mtb_ComponentMacro"].ToString(), out _ComponentMacro);
                short.TryParse(ConfigurationManager.AppSettings["Mtb_CycleEndDateTimeMacro"].ToString(), out _CycleEndDateTimeMacro);
                short.TryParse(ConfigurationManager.AppSettings["Mtb_DataReadflagMacro"].ToString(), out _DataReadflagMacro);
                short.TryParse(ConfigurationManager.AppSettings["Mtb_featureIDMacro"].ToString(), out _featureIDMacro);
                short.TryParse(ConfigurationManager.AppSettings["Mtb_DataReadStatusFlagMacro"].ToString(), out _DataReadStatusFlagMacro);
                double.TryParse(ConfigurationManager.AppSettings["Mtb_InspectionReadDelayInSec"].ToString(), out _timeDelayMainThread);
                short.TryParse(ConfigurationManager.AppSettings["Mtb_ComponentGroupIdMacro"].ToString(), out _ComponentGroupIdMacro);
            }
            else if (this.ControllerType.Equals("FANUC", StringComparison.OrdinalIgnoreCase))
            {
                short.TryParse(ConfigurationManager.AppSettings["ComponentMacro"].ToString(), out _ComponentMacro);
                short.TryParse(ConfigurationManager.AppSettings["CycleEndDateTimeMacro"].ToString(), out _CycleEndDateTimeMacro);
                short.TryParse(ConfigurationManager.AppSettings["DataReadflagMacro"].ToString(), out _DataReadflagMacro);
                short.TryParse(ConfigurationManager.AppSettings["featureIDMacro"].ToString(), out _featureIDMacro);
                short.TryParse(ConfigurationManager.AppSettings["DataReadStatusFlagMacro"].ToString(), out _DataReadStatusFlagMacro);
                double.TryParse(ConfigurationManager.AppSettings["InspectionReadDelayInSec"].ToString(), out _timeDelayMainThread);
                short.TryParse(ConfigurationManager.AppSettings["ComponentGroupIdMacro"].ToString(), out _ComponentGroupIdMacro);
            }
            

            //int.TryParse(ConfigurationManager.AppSettings["SystemType"].ToString(), out systemType);
            //int.TryParse(ConfigurationManager.AppSettings["MachineType"].ToString(), out machineType);
            _timeDelayMainThread = TimeSpan.FromSeconds(_timeDelayMainThread).TotalMilliseconds;
            if (_timeDelayMainThread <= 500) _timeDelayMainThread = 500;

            while (true)
            {
                try
                {
                    #region stop_service                   
                    if (ServiceStop.stop_service == 1)
                    {
                        try
                        {
                            Logger.WriteDebugLog("stopping the service. coming out of main while loop.");
                            break;
                        }
                        catch (Exception ex)
                        {
                            Logger.WriteErrorLog(ex.Message);
                            break;
                        }
                    }
                    #endregion

                    try
                    {
                        if (Utility.CheckPingStatus(this.ipAddress))
                        {
                            #region Lic check
                            if (this.ControllerType.Equals("FANUC", StringComparison.OrdinalIgnoreCase) && !_isCNCIdReadSuccessfully)
                            {
                                string cncId = string.Empty;
                                List<string> cncIdList = FocasSmartDataService.licInfo.CNCData.Where(s => s.CNCdata1 != null).Select(s => s.CNCdata1).ToList();
                                _isLicenseValid = this.ValidateCNCSerialNo(this.machineId, this.ipAddress, this.portNo, cncIdList, out _isCNCIdReadSuccessfully, out cncId);

                                if (!_isLicenseValid)
                                {
                                    if (_isCNCIdReadSuccessfully)
                                    {
                                        Logger.WriteErrorLog("Lic Validation failed. Please contact AMIT/MMT.");
                                        break;
                                    }
                                    Thread.Sleep(TimeSpan.FromSeconds(10.0));
                                    continue;
                                }
                                //update table 
                                if (_isLicenseValid)
                                {
                                    this.SetCNCDateTime(this.machineId, this.ipAddress, this.portNo);
                                }
                            }
                            #endregion

                            //Inspection data read for Ashoka Leyland
                            if (this.ControllerType.Equals("FANUC",StringComparison.OrdinalIgnoreCase))
                                ReadInspectionDataAshokLeyland(this.machineId, this.ipAddress, this.portNo);
                            else if(this.ControllerType.Equals("MITSUBISHI", StringComparison.OrdinalIgnoreCase))
                                ReadInspectionDataAshokLeylandMitsubishi(this.machineId, this.ipAddress, this.portNo,this.systemType,this.machineType);
                        }
                        else
                        {
                            if (ServiceStop.stop_service == 1) break;
                            Thread.Sleep(1000);
                        }
                        if (_timeDelayMainThread > 0)
                        {
                            if (ServiceStop.stop_service == 1) break;
                            Thread.Sleep((int)_timeDelayMainThread);
                        }
                    }
                    catch (Exception e)
                    {
                        Logger.WriteErrorLog("Exception inside main while loop : " + e.ToString());
                        Thread.Sleep(1000 * 4);
                    }
                    finally
                    {

                    }
                }
                catch (Exception ex)
                {
                    Logger.WriteErrorLog("Exception from main while loop : " + ex.ToString());
                    Thread.Sleep(2000);
                }
            }
            this.CloseTimer();
            Logger.WriteDebugLog("End of while loop." + Environment.NewLine + "------------------------------------------");
        }

        private void ReadInspectionDataAshokLeylandMitsubishi(string machineId, string ipAddress, ushort portNo, int systemType, int machineType)
        {
            try
            {
                int ret = 0;
                ret = MitsubishiLib.Open2(ipAddress, portNo, systemType, machineType, 10);
                if (ret != 0)
                {
                    Logger.WriteErrorLog("MitsubishiLib.Open2() failed during ReadInspectionDataAshokLeylandMitsubishi() . return value is = " + ret);
                    Thread.Sleep(1000);
                    return;
                }

                //check new inspection data has been ready to read.
                int isDataChanged = Convert.ToInt32(MitsubishiLib.ReadMacroParameter(_DataReadflagMacro));
                //Logger.WriteErrorLog("Data Read Flag Value = " + isDataChanged);
                if (isDataChanged == 1)
                {
                    Logger.WriteDebugLog("Started reading Inspection data.");

                    int CompInterface = Convert.ToInt32(MitsubishiLib.ReadMacroParameter(_ComponentMacro));
                    int featureID = Convert.ToInt32(MitsubishiLib.ReadMacroParameter(_featureIDMacro));
                    int ComponentGroupIdMacroValue = 0;
                    if (_ComponentGroupIdMacro > 0)
                    {

                        ComponentGroupIdMacroValue = Convert.ToInt32(MitsubishiLib.ReadMacroParameter(_ComponentGroupIdMacro));
                    }
                    //Read CycleEnd datetime from macro
                    var CNCTimeStamp = MitsubishiLib.ReadCNCCycleEndTimeStamp(_CycleEndDateTimeMacro, (_CycleEndDateTimeMacro + 1));
                    //read inspection value from macro location
                    var featureName = DatabaseAccess.GetFeatures(featureID);
                    var inspectionData = DatabaseAccess.GetSPC_CharacteristicsForMCOAshokaLeyland(CompInterface.ToString());
                    Logger.WriteDebugLog(string.Format("Read data : Component Interface ID = {0}, Feature ID = {1}, Feature = {2}, Cycle End TS = {3}, ComponentGroup ID Value = {4}", CompInterface, featureID, featureName, CNCTimeStamp.ToString("yyyy-MM-dd HH:mm:ss"), ComponentGroupIdMacroValue));

                    var insListToRead = inspectionData.Where(item => item.FeatureID.Equals(featureName, StringComparison.OrdinalIgnoreCase)).ToList();
                    if (insListToRead.Count == 0)
                        Logger.WriteDebugLog(string.Format("Master Data not found for Feature {0} in table SPC_Characteristic for Component interface id {1}", featureName, CompInterface));
                    foreach (var item in insListToRead)
                    {
                        double inspectionValue = MitsubishiLib.ReadMacroParameter(item.MacroLocation);
                        if (inspectionValue != double.MaxValue)
                        {
                            item.DiamentionValue = inspectionValue;
                        }
                    }

                    //build type 37 string and insert to database
                    foreach (var item in insListToRead)
                    {
                        if (item.DiamentionValue == double.MaxValue || item.DiamentionValue == 0) continue;
                        var str = string.Format("START-[{0}]-[{1}]-[{2}]-[{3}]-[{4}]-[{5}]-[{6}]END", this.machineId, item.ComponentID, item.FeatureID, item.DiamentionId, item.DiamentionValue, CNCTimeStamp.ToString("yyyy-MM-dd HH:mm:ss"), ComponentGroupIdMacroValue);
                        SaveStringToTPMFile(str);
                        DatabaseAccess.SaveToDatabase(this.machineId, item.ComponentID, item.FeatureID, item.DiamentionId, item.DiamentionValue, CNCTimeStamp, ComponentGroupIdMacroValue);
                    }

                    //update all macro location value to '0'
                    //foreach (var item in inspectionData)
                    //{
                    //    FocasData.WriteMacro(focasLibHandle, item.MacroLocation, 0);
                    //}

                    //reset the data read flag to '0'
                    if (_DataReadStatusFlagMacro > 0)
                        MitsubishiLib.WriteMacroParameter(_DataReadStatusFlagMacro, 0);
                    MitsubishiLib.WriteMacroParameter(_DataReadflagMacro, 2);
                    
                    Logger.WriteDebugLog("Completed reading Inspection data. Updated 2 to Macro Location " + _DataReadflagMacro);
                }

                //if (focasLibHandle > 0)
                //{
                //    var r = FocasData.cnc_freelibhndl(focasLibHandle);
                //    //if (r != 0) _focasHandles.Add(focasLibHandle);
                //}
            }
            catch (Exception exx)
            {
                Logger.WriteErrorLog(exx.ToString());
            }
        }

        public void GetToolLifeData(Object stateObject)
        {
            if (!_isLicenseValid) return;
            if (Monitor.TryEnter(_lockerToolLife, 100))
            {
              
                try
                {
                    Thread.CurrentThread.Name = "ToolLifeData-" + this.machineId;
                    if (Utility.CheckPingStatus(this.ipAddress))
                    {                       
                        if (_isMacroStringEnabled == false)
                        {
                            ProcessToolLifeUsingFOCAS(this.machineId, this.ipAddress, this.portNo, setting, 2);
                        }
                        else
                        {
                            ProcessToolLifeUsingWinMach(this.machineId, this.ipAddress, this.portNo);
                        }

                        //if (this.machineDTO.MTB.Equals("ACE", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    Logger.WriteDebugLog("ACE : Reading Tool Life History data for control type." + _cncMachineType.ToString());
                        //    ProcessToolLifeUsingFOCAS(this.machineId, this.ipAddress, this.portNo, setting, 2);
                        //}
                        //else if (this.machineDTO.MTB.Equals("AMS", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    Logger.WriteDebugLog("AMS : Reading Tool Life History data for control type." + _cncMachineType.ToString());
                        //    ProcessToolLifeUsingDVariableAMS(this.machineId, this.ipAddress, this.portNo, _toolLifeDefaults);
                        //}
                        //else if (this.machineDTO.MTB.Equals("Kennametal", StringComparison.OrdinalIgnoreCase))
                        //{
                        //    Logger.WriteDebugLog("Kennametal : Reading Tool Life History data for control type." + _cncMachineType.ToString());
                        //    ProcessToolLifeUsingDVariableKennametal(this.machineId, this.ipAddress, this.portNo, _toolLifeDefaults);
                        //}
                        //else
                        //{
                        //    Logger.WriteDebugLog("ACE : Reading Tool Life History data for control type." + _cncMachineType.ToString());
                        //    ProcessToolLifeUsingFOCAS(this.machineId, this.ipAddress, this.portNo, setting, 2);
                        //}

                        
                    }
                }
                catch (Exception ex)
                {
                    Logger.WriteDebugLog(ex.ToString());
                }
                finally
                {                   
                    Monitor.Exit(_lockerToolLife);
                }
            }
        }

        private void ProcessToolLifeUsingFOCAS(string machineId, string ipAddress, ushort portNo, MachineSetting setting, int spindleType)
        {
            //Read n(12-12) macro memory location for tool target, Actual values and store to database.            
            //read every 5 minutes
            int ret = 0;
            ushort focasLibHandle = 0;
            ret = FocasData.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
            if (ret != 0)
            {
                Logger.WriteErrorLog("ProcessToolLife()--> cnc_allclibhndl3() failed. return value is = " + ret);
                return;
            }

            DateTime cNCTimeStamp = FocasData.ReadCNCTimeStamp(focasLibHandle);

            int component = FocasData.ReadMacro(focasLibHandle, _CompMacroLocation);
            int operation = FocasData.ReadMacro(focasLibHandle, _OpnMacroLocation);
            short programNo = FocasData.ReadMainProgram(focasLibHandle);
            var partsCount = FocasData.ReadPartsCount(focasLibHandle);

            Logger.WriteDebugLog(string.Format("Comp = {0}; Operation = {1} ; Program No = {2} ; PartsCount = {3}", component, operation, programNo, partsCount));
            //number of groups
            FocasLibBase.ODBTLIFE2 a = new FocasLibBase.ODBTLIFE2();
            ret = FocasLibrary.FocasLib.cnc_rdngrp(focasLibHandle, a);           
            if (ret != 0)
            {
                Logger.WriteErrorLog("Reading groups issue. return value is = " + ret);               
            }

            short NoOfGroups = (short)a.data;
            Logger.WriteDebugLog("No Of groups = " + NoOfGroups);

            //FocasLibBase.ODBUSEGRP aa = new FocasLibBase.ODBUSEGRP();
            //ret = FocasLibrary.FocasLib.cnc_rdtlusegrp(focasLibHandle, aa);
            //Logger.WriteDebugLog("cnc_rdtlusegrp = " + aa.use.ToString() + " --" + aa.next.ToString());

            //loop all the groups
            List<ToolLifeDO> toolife = new List<ToolLifeDO>();
            for (short groupNo = 1; groupNo <= NoOfGroups; groupNo++)
            {
                FocasLibBase.ODBTG c = new FocasLibBase.ODBTG();
                ret = FocasLibrary.FocasLib.cnc_rdtoolgrp(focasLibHandle, groupNo, (short)System.Runtime.InteropServices.Marshal.SizeOf(c), c);
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_rdtoolgrp issue. return value is = " + ret);
                }
                //for each tool in a group, create the ToolLifeDO object               
                List<ToolLifeDO> toolsByGroup = getEachToolData(c, component, operation, programNo, cNCTimeStamp, partsCount);

                toolife.AddRange(toolsByGroup);
            }

            DatabaseAccess.DeleteToolLifeTempRecords(this.machineId);
            DatabaseAccess.InsertBulkRows(toolife.ToDataTable<ToolLifeDO>(), "[dbo].[Focas_ToolLifeTemp]");
            DatabaseAccess.ProcessToolLifeTempToHistory(this.machineId);
            //DatabaseAccess.DeleteToolLifeTempRecords(this.machineId);

            if (focasLibHandle > 0)
            {
                FocasData.cnc_freelibhndl(focasLibHandle);
            }
        }

        private void ProcessToolLifeUsingWinMach(string machineId, string ipAddress, ushort portNo)
        {
            ushort focasLibHandle = 1;
            try
            {
                short ret = FocasData.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
                if (ret == 0)
                {

                    short offsetLocation = 0;
                    bool cc = short.TryParse(_toolLifeStartLocation, out offsetLocation);
                    if (offsetLocation > 0)
                    {
                        List<TPMString> list = new List<TPMString>();
                        //foreach (short current in new List(){offsetLocation })
                        {
                            int isDataReadyToread = FocasData.ReadMacro(focasLibHandle, offsetLocation);
                            if (isDataReadyToread > 0)
                            {
                                List<int> values = FocasData.ReadMacroRange(focasLibHandle, (short)(offsetLocation + 1), (short)(offsetLocation + 10));
                                TPMString tPMString = new TPMString();
                                tPMString.Seq = values[0];
                                var text = this.BuildStringOffset(values);
                                tPMString.TpmString = text;
                                this.SaveStringToTPMFile(text);
                                //tPMString.DateTime = this.GetDatetimeFromtpmString(values);

                                list.Add(tPMString);
                                FocasData.WriteMacro(focasLibHandle, offsetLocation, 0);

                                foreach (TPMString current2 in list.OrderBy(s => s.Seq))
                                {
                                    this.ProcessData(current2.TpmString, this.ipAddress, this.portNo.ToString(), this.machineId);
                                }
                                Logger.WriteDebugLog("Completed Tool Life History data.");
                            }
                            else
                            {
                                Logger.WriteDebugLog("Reading Tool Life data. Read flag is not high");
                            }
                        }
                    }
                }
                else
                {
                    Logger.WriteErrorLog("Not able to connect to CNC machine. ret value from fun cnc_allclibhndl3 = " + ret);
                }
            }
            catch(Exception exx)
            {

            }
            finally
            {
                if (focasLibHandle > 0)
                {
                    FocasData.cnc_freelibhndl(focasLibHandle);
                }
            }

        }

        private List<ToolLifeDO> getEachToolData(FocasLibBase.ODBTG c, int component, int operation, int programNo, DateTime cNCTimeStamp, int partsCount)
        {
            Logger.WriteDebugLog(string.Format("Group No = {0} , Tool Target = {1} , Tool Actual = {2}",c.grp_num,c.life,c.count));
            List<ToolLifeDO> toolife = new List<ToolLifeDO>();
            if (c.data.data1.tool_num > 0)
            {
                var tool = new ToolLifeDO()
                {
                    CNCTimeStamp = cNCTimeStamp,
                    ComponentID = component.ToString(),
                    MachineID = machineId,
                    OperationID = operation.ToString(),
                    ProgramNo = programNo,
                    ToolTarget = c.life,
                    ToolActual = c.count,
                    ToolNo = c.data.data1.tool_num.ToString(),
                    SpindleType = c.grp_num,
                    ToolUseOrderNumber = c.data.data1.tuse_num,
                    ToolInfo = c.data.data1.tinfo,
                    PartsCount = partsCount,
                };
                toolife.Add(tool);
            }
            if (c.data.data2.tool_num > 0)
            {
                var tool = new ToolLifeDO()
                {
                    CNCTimeStamp = cNCTimeStamp,
                    ComponentID = component.ToString(),
                    MachineID = machineId,
                    OperationID = operation.ToString(),
                    ProgramNo = programNo,
                    ToolTarget = c.life,
                    ToolActual = c.count,
                    ToolNo = c.data.data2.tool_num.ToString(),
                    SpindleType = c.grp_num,
                    ToolUseOrderNumber = c.data.data2.tuse_num,
                    ToolInfo = c.data.data2.tinfo,
                    PartsCount = partsCount,
                };
                toolife.Add(tool);
            }
            if (c.data.data3.tool_num > 0)
            {
                var tool = new ToolLifeDO()
                {
                    CNCTimeStamp = cNCTimeStamp,
                    ComponentID = component.ToString(),
                    MachineID = machineId,
                    OperationID = operation.ToString(),
                    ProgramNo = programNo,
                    ToolTarget = c.life,
                    ToolActual = c.count,
                    ToolNo = c.data.data3.tool_num.ToString(),
                    SpindleType = c.grp_num,
                    ToolUseOrderNumber = c.data.data3.tuse_num,
                    ToolInfo = c.data.data3.tinfo,
                    PartsCount = partsCount,
                };
                toolife.Add(tool);

            }
            if (c.data.data4.tool_num > 0)
            {
                var tool = new ToolLifeDO()
                {
                    CNCTimeStamp = cNCTimeStamp,
                    ComponentID = component.ToString(),
                    MachineID = machineId,
                    OperationID = operation.ToString(),
                    ProgramNo = programNo,
                    ToolTarget = c.life,
                    ToolActual = c.count,
                    ToolNo = c.data.data4.tool_num.ToString(),
                    SpindleType = c.grp_num,
                    ToolUseOrderNumber = c.data.data4.tuse_num,
                    ToolInfo = c.data.data4.tinfo,
                    PartsCount = partsCount,
                };
                toolife.Add(tool);
            }
            if (c.data.data5.tool_num > 0)
            {
                var tool = new ToolLifeDO()
                {
                    CNCTimeStamp = cNCTimeStamp,
                    ComponentID = component.ToString(),
                    MachineID = machineId,
                    OperationID = operation.ToString(),
                    ProgramNo = programNo,
                    ToolTarget = c.life,
                    ToolActual = c.count,
                    ToolNo = c.data.data5.tool_num.ToString(),
                    SpindleType = c.grp_num,
                    ToolUseOrderNumber = c.data.data5.tuse_num,
                    ToolInfo = c.data.data5.tinfo,
                    PartsCount = partsCount,
                };
                toolife.Add(tool);
            }



            if (c.data.data6.tool_num > 0)
            {
                var tool = new ToolLifeDO()
                {
                    CNCTimeStamp = cNCTimeStamp,
                    ComponentID = component.ToString(),
                    MachineID = machineId,
                    OperationID = operation.ToString(),
                    ProgramNo = programNo,
                    ToolTarget = c.life,
                    ToolActual = c.count,
                    ToolNo = c.data.data6.tool_num.ToString(),
                    SpindleType = c.grp_num,
                    ToolUseOrderNumber = c.data.data6.tuse_num,
                    ToolInfo = c.data.data6.tinfo,
                    PartsCount = partsCount,
                };
                toolife.Add(tool);
            }

            if (c.data.data7.tool_num > 0)
            {
                var tool = new ToolLifeDO()
                {
                    CNCTimeStamp = cNCTimeStamp,
                    ComponentID = component.ToString(),
                    MachineID = machineId,
                    OperationID = operation.ToString(),
                    ProgramNo = programNo,
                    ToolTarget = c.life,
                    ToolActual = c.count,
                    ToolNo = c.data.data7.tool_num.ToString(),
                    SpindleType = c.grp_num,
                    ToolUseOrderNumber = c.data.data7.tuse_num,
                    ToolInfo = c.data.data7.tinfo,
                    PartsCount = partsCount,
                };
                toolife.Add(tool);
            }

            if (c.data.data8.tool_num > 0)
            {
                var tool = new ToolLifeDO()
                {
                    CNCTimeStamp = cNCTimeStamp,
                    ComponentID = component.ToString(),
                    MachineID = machineId,
                    OperationID = operation.ToString(),
                    ProgramNo = programNo,
                    ToolTarget = c.life,
                    ToolActual = c.count,
                    ToolNo = c.data.data8.tool_num.ToString(),
                    SpindleType = c.grp_num,
                    ToolUseOrderNumber = c.data.data8.tuse_num,
                    ToolInfo = c.data.data8.tinfo,
                    PartsCount = partsCount,
                };
                toolife.Add(tool);
            }

            if (c.data.data9.tool_num > 0)
            {
                var tool = new ToolLifeDO()
                {
                    CNCTimeStamp = cNCTimeStamp,
                    ComponentID = component.ToString(),
                    MachineID = machineId,
                    OperationID = operation.ToString(),
                    ProgramNo = programNo,
                    ToolTarget = c.life,
                    ToolActual = c.count,
                    ToolNo = c.data.data9.tool_num.ToString(),
                    SpindleType = c.grp_num,
                    ToolUseOrderNumber = c.data.data9.tuse_num,
                    ToolInfo = c.data.data9.tinfo,
                    PartsCount = partsCount,
                };
                toolife.Add(tool);
            }

            if (c.data.data10.tool_num > 0)
            {
                var tool = new ToolLifeDO()
                {
                    CNCTimeStamp = cNCTimeStamp,
                    ComponentID = component.ToString(),
                    MachineID = machineId,
                    OperationID = operation.ToString(),
                    ProgramNo = programNo,
                    ToolTarget = c.life,
                    ToolActual = c.count,
                    ToolNo = c.data.data10.tool_num.ToString(),
                    SpindleType = c.grp_num,
                    ToolUseOrderNumber = c.data.data10.tuse_num,
                    ToolInfo = c.data.data10.tinfo,
                    PartsCount = partsCount,
                };
                toolife.Add(tool);
            }

            return toolife;
        }

        private void SetCNCDateTime(string machineId, string ipAddress, ushort port)
        {           
            Ping ping = null;
            ushort focasLibHandle = 0;
            try
            {
                ping = new Ping();
                PingReply pingReply = null;
                int count = 0;
                while (true)
                {
                    pingReply = ping.Send(ipAddress, 10000);
                    if (pingReply.Status != IPStatus.Success)
                    {
                        if (ServiceStop.stop_service == 1) break;
                        Logger.WriteErrorLog("Not able to ping. Ping status = " + pingReply.Status.ToString());
                        Thread.Sleep(2000);
                    }
                    else if (pingReply.Status == IPStatus.Success || ServiceStop.stop_service == 1 || count == 4)
                    {
                        break;
                    }
                    ++count;
                }
                if (pingReply.Status == IPStatus.Success)
                {
                    int num2 = FocasData.cnc_allclibhndl3(ipAddress, port, 10, out focasLibHandle);
                    if (num2 == 0)
                    {
                        FocasData.SetCNCDate(focasLibHandle, DateTime.Now);
                        FocasData.SetCNCTime(focasLibHandle, DateTime.Now);                                             
                    }
                    else
                    {
                        Logger.WriteErrorLog("Not able to connect to machine. cnc_allclibhndl3 status = " + num2.ToString());
                    }
                }
                else
                {
                    Logger.WriteErrorLog("Not able to ping. Ping status = " + pingReply.Status.ToString());
                }
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog(ex.ToString());
            }
            finally
            {
                if (ping != null)
                {
                    ping.Dispose();
                }
                if (focasLibHandle != 0)
                {
                    short num3 = FocasData.cnc_freelibhndl(focasLibHandle);                   
                }
            }            
        }

        private List<int> GetOffsetRange()
        {
            List<int> range = new List<int>();
            try
            {
                string offsetRange = ConfigurationManager.AppSettings["OffsetHistoryRange"].ToString();
                var splitRange = offsetRange.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string item in splitRange)
                {
                    var rangeArr = item.Split(new char[] { '-' }, StringSplitOptions.RemoveEmptyEntries);

                    for (int i = int.Parse(rangeArr[0]); i <= int.Parse(rangeArr[1]); i++)
                    {
                        if (!range.Exists(t => t == i))
                        {
                            range.Add(i);
                        }
                    }
                }
            }
            catch (Exception eee)
            {
                Logger.WriteErrorLog(eee.ToString());
            }
            return range;
        }       

        private void ReadPredictiveMaintenanceData(Object stateObject)
        {
           if (this.machineDTO.Settings.PredictiveMaintenanceSettings == null || this.machineDTO.Settings.PredictiveMaintenanceSettings.Count == 0)
                return;

            if (!_isLicenseValid) return;
            if (Monitor.TryEnter(this._lockerPredictiveMaintenance))
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                ushort focasLibHandle = 0;
                string text = string.Empty;              
                try
                {
                    Thread.CurrentThread.Name = "PredictiveMaintenanceDataCollation-" + this.machineId;

                    if (Utility.CheckPingStatus(this.ipAddress))
                    {
                        try
                        {
                            int ret = 0;
                            ret = FocasData.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
                            if (ret != 0)
                            {
                                Logger.WriteErrorLog("cnc_allclibhndl3() failed during ReadPredictiveMaintenanceData() . return value is = " + ret);
                                Thread.Sleep(1000);
                                return;
                            }

                            Logger.WriteDebugLog("Reading PredictiveMaintenance data....");
                            DateTime cncTimeStamp = FocasData.ReadCNCTimeStamp(focasLibHandle);
                            foreach (var item in this.machineDTO.Settings.PredictiveMaintenanceSettings)
                            {
                                item.MachineId = machineId;
                                item.TimeStamp = cncTimeStamp;
                                item.TargetValue = 0;
                                item.ActualValue = 0;
                                int targerValue = 0;
                                int currentValue = 0;

                                FocasData.GetPredictiveMaintenanceTargetCurrent(focasLibHandle, item.TargetDLocation, item.CurrentValueDLocation, out targerValue, out currentValue,this.machineDTO.MTB);
                                if (targerValue != int.MaxValue)
                                {
                                    item.TargetValue = targerValue;
                                }
                                if (currentValue != int.MaxValue)
                                {
                                    item.ActualValue = currentValue;
                                }
                            }
                            
                            //insert data to datbase......
                            DataTable dt = this.machineDTO.Settings.PredictiveMaintenanceSettings.ToDataTable<PredictiveMaintenanceDTO>();
                            dt.Columns.Remove("TargetDLocation"); dt.Columns.Remove("CurrentValueDLocation");
                            DatabaseAccess.InsertBulkRows(dt, "[dbo].[Focas_PredictiveMaintenanceTemp]");
                            //Update main table from temp table
                            DatabaseAccess.ProcessTempTableToMainTable(this.machineId, "Predictive");
                            DatabaseAccess.DeleteTempTableRecords(this.machineId, "Focas_PredictiveMaintenanceTemp");
                            Logger.WriteDebugLog("Completed reading PredictiveMaintenance data....");
                        }
                        catch (Exception exx)
                        {
                            Logger.WriteErrorLog(exx.ToString());
                        }
                    }
                }
                catch (Exception exx)
                {
                    Logger.WriteErrorLog(exx.ToString());
                }
                finally
                {
                    if (focasLibHandle > 0)
                    {
                      var r = FocasData.cnc_freelibhndl(focasLibHandle);                     
                    }                   
                    Monitor.Exit(_lockerPredictiveMaintenance);
                }
            }
        }

        private void ReadOffsetHistoryData(Object stateObject)
        {
            if (!_isLicenseValid) return;
            if (this.offsetHistoryList == null || this.offsetHistoryList.Count == 0)
                return;
           
            if (Monitor.TryEnter(this._lockerOffsetHistory))
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                ushort focasLibHandle = 0;
                string text = string.Empty;
                Ping ping = null;
                try
                {
                    Thread.CurrentThread.Name = string.IsNullOrWhiteSpace(Thread.CurrentThread.Name) ? "OffsetHistoryDataCollation-" + this.machineId : Thread.CurrentThread.Name;


                    if (Utility.CheckPingStatus(this.ipAddress))
                    {
                        try
                        {
                            int ret = 0;
                            ret = FocasData.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
                            if (ret != 0)
                            {
                                Logger.WriteErrorLog("cnc_allclibhndl3() failed during ReadOffsetHistoryData() . return value is = " + ret);
                                Thread.Sleep(1000);
                                return;
                            }

                             Logger.WriteDebugLog("Reading Offset values....");
                            DateTime cncTimeStamp = FocasData.ReadCNCTimeStamp(focasLibHandle);
                            string machineStatus;
                            var machineMode = FocasData.ReadMachineStatusMode(focasLibHandle, out machineStatus);
                            var dynamic_data = FocasData.cnc_rddynamic2(focasLibHandle);
                            foreach (var item in offsetHistoryList)
                            {
                                item.CNCTimeStamp = cncTimeStamp;
                                item.ProgramNo = "O" + dynamic_data.prgmnum.ToString();
                                item.MachineMode = machineMode;
                                item.ToolNo = FocasData.ReadToolNo(focasLibHandle);
                                item.WearOffsetX = FocasData.get_wear_offset(focasLibHandle, item.OffsetNo, 'x');
                                item.WearOffsetZ = FocasData.get_wear_offset(focasLibHandle, item.OffsetNo, 'z');
                                item.WearOffsetR = 0; // FocasData.get_wear_offset(focasLibHandle, item.OffsetNo, 'r');
                                item.WearOffsetT = 0;// FocasData.get_wear_offset(focasLibHandle, item.OffsetNo, 't');
                            }
                            //insert data to datbase......
                            DataTable dt = offsetHistoryList.ToDataTable<OffsetHistoryDTO>();
                            DatabaseAccess.InsertBulkRows(dt, "[dbo].[Focas_ToolOffsetHistoryTemp]");
                            
                            DatabaseAccess.ProcessTempTableToMainTable(this.machineId, "OffsetHistory");
                            DatabaseAccess.DeleteTempTableRecords(this.machineId, "Focas_ToolOffsetHistoryTemp");
                            Logger.WriteDebugLog("Completed reading Offset values....");
                        }
                        catch (Exception exx)
                        {
                            Logger.WriteErrorLog(exx.ToString());
                        }
                        finally
                        {
                            if (focasLibHandle > 0)
                            {
                                var r = FocasData.cnc_freelibhndl(focasLibHandle);
                                //if (r != 0) _focasHandles.Add(focasLibHandle);
                            }
                        }
                    }
                }
                catch (Exception exx)
                {
                    Logger.WriteErrorLog(exx.ToString());
                }
                finally
                {                  
                    Monitor.Exit(_lockerOffsetHistory);                   
                }
            }
        }
        private bool ValidateCNCSerialNo(string machineId, string ipAddress, ushort port, List<string> cncSerialnumbers, out bool isLicCheckedSucessfully, out string cncID)
        {
            bool result = false;
            isLicCheckedSucessfully = true;
            Ping ping = null;
            ushort focasLibHandle = 0;
            cncID = string.Empty;

            try
            {
                ping = new Ping();
                PingReply pingReply = null;
                while (true)
                {
                    pingReply = ping.Send(ipAddress, 10000);
                    if (pingReply.Status != IPStatus.Success)
                    {
                        if (ServiceStop.stop_service == 1) break;
                        Logger.WriteErrorLog("Not able to ping. Ping status = " + pingReply.Status.ToString());
                        Thread.Sleep(10000);
                    }
                    else if (pingReply.Status == IPStatus.Success || ServiceStop.stop_service == 1)
                    {
                        break;
                    }
                }
                if (pingReply.Status == IPStatus.Success)
                {
                    int num2 = FocasData.cnc_allclibhndl3(ipAddress, port, 10, out focasLibHandle);
                    if (num2 == 0)
                    {
                        string text = FocasData.ReadCNCId(focasLibHandle);
                        if (!string.IsNullOrEmpty(text))
                        {
                            if (cncSerialnumbers.Contains(text))
                            {
                                cncID = text;
                                result = true;
                            }
                        }
                        else
                        {
                            isLicCheckedSucessfully = false;
                        }
                    }
                    else
                    {
                        Logger.WriteErrorLog("Not able to connect to machine. cnc_allclibhndl3 status = " + num2.ToString());
                        isLicCheckedSucessfully = false;
                    }
                }
                else
                {
                    Logger.WriteErrorLog("Not able to ping. Ping status = " + pingReply.Status.ToString());
                    isLicCheckedSucessfully = false;
                }
            }
            catch (Exception ex)
            {
                isLicCheckedSucessfully = false;
                Logger.WriteDebugLog(ex.ToString());
            }
            finally
            {
                if (ping != null)
                {
                    ping.Dispose();
                }
                if (focasLibHandle != 0)
                {
                    short num3 = FocasData.cnc_freelibhndl(focasLibHandle);
                    //if (num3 != 0) _focasHandles.Add(focasLibHandle);
                }
            }
            return result;
        }

        private bool ValidateMachineModel(string machineId, string ipAddress, ushort port)
        {
            bool result = false;           
            Ping ping = null;
            ushort focasLibHandle = 0;
            try
            {
                ping = new Ping();
                PingReply pingReply = null;
                while (true)
                {
                    pingReply = ping.Send(ipAddress, 10000);
                    if (pingReply.Status != IPStatus.Success)
                    {
                        if (ServiceStop.stop_service == 1) break;
                        Logger.WriteErrorLog("Not able to ping. Ping status = " + pingReply.Status.ToString());
                        Thread.Sleep(10000);
                    }
                    else if (pingReply.Status == IPStatus.Success || ServiceStop.stop_service == 1)
                    {
                        break;
                    }
                }
                if (pingReply.Status == IPStatus.Success)
                {
                    int num2 = FocasData.cnc_allclibhndl3(ipAddress, port, 10, out focasLibHandle);
                    if (num2 == 0)
                    {
                        int mcModel = FocasData.ReadParameterInt(focasLibHandle, 4133);
                        int maxSpeedOnMotor = FocasData.ReadParameterInt(focasLibHandle, 4020);
                        int maxSpeedOnSpindle = FocasData.ReadParameterInt(focasLibHandle, 3741);
                        if (mcModel > 0)
                        {
                            DatabaseAccess.UpdateMachineModel(machineId, mcModel);
                        }
                    }
                    else
                    {
                        Logger.WriteErrorLog("Not able to connect to machine. cnc_allclibhndl3 status = " + num2.ToString());                       
                    }
                }
                else
                {
                    Logger.WriteErrorLog("Not able to ping. Ping status = " + pingReply.Status.ToString());                   
                }
            }
            catch (Exception ex)
            {              
                Logger.WriteDebugLog(ex.ToString());
            }
            finally
            {
                if (ping != null)
                {
                    ping.Dispose();
                }
                if (focasLibHandle != 0)
                {
                    short num3 = FocasData.cnc_freelibhndl(focasLibHandle);
                   // if (num3 != 0) _focasHandles.Add(focasLibHandle);
                }
            }
            return result;
        }


        //read comments-done, parse sub program-done, 
        //TODO - file naming in master file, file naming while saving to "AutoDownloadFolder"
        private int DownloadProgram(string machineId, string ipAddress, ushort port, string programNumber)
        {
            //download running program in temp folder
            bool result = false;
            int programNo = 0;
            programNumber = programNumber.TrimStart(new char[] { 'O' }).Trim();
            int.TryParse(programNumber, out programNo);
            if (programNo == 0) return 0;
            string mainProgramMasterStr = string.Empty;
            StringBuilder messageForSMS = new StringBuilder();
            Logger.WriteDebugLog("Downloading main program : " + programNo);
            string mainProgramCNCStr = FocasData.DownloadProgram(ipAddress, port, programNo, out result);
            if (!result) return -1;
            Hashtable hashSubPrograms = new Hashtable();
            //check if main program contains sub peogram(contains with M98P)
            string mainProgramCNCComment = FindProgramComment(mainProgramCNCStr);
            List<int> subProgramsCNC = FindSubPrograms(mainProgramCNCStr);
            List<int> subProgramsCNCTemp = new List<int>();
            if (subProgramsCNC.Count > 0)
            {
                //download sub programs starts with M98P
                foreach (var item in subProgramsCNC)
                {
                    Logger.WriteDebugLog("Downloading first level sub program : " + item);
                    string prgText = FocasData.DownloadProgram(ipAddress, port, item, out result);
                    if (result)
                    {
                        hashSubPrograms.Add(item, prgText);
                        //find second level sub programs
                        subProgramsCNCTemp.AddRange(FindSubPrograms(prgText));
                    }
                    //else
                    //    return -1;
                }
            }

            //download second level sub programs
            if (subProgramsCNCTemp.Count > 0)
            {
                //download sub programs starts with M98P
                foreach (var item in subProgramsCNCTemp.Distinct())
                {
                    Logger.WriteDebugLog("Downloading second level sub program : " + item);
                    string prgText = FocasData.DownloadProgram(ipAddress, port, item, out result);
                    if (result)
                    {
                        if (!hashSubPrograms.ContainsKey(item))
                        {
                            hashSubPrograms.Add(item, prgText);
                        }
                    }
                    //else
                    //    return -1;
                }
            }

            //compare main, sub program with Master folder program --> if not same, store the main and sub program with date time
            //O1234_yyyymmddhhmm.txt , O1234_567_yyyymmddhhmm.txt, O1234_678_yyyymmddhhmm.txt           
            CreateDirectory(_programDownloadFolder);     
           
            //compaire the containt of main and sub program from "Master" folder under machine folder??
            string masterProgramFolderPath = Path.Combine(_programDownloadFolder, "MasterPrograms", "O" + programNumber + mainProgramCNCComment);
            string autoDownloadedProgramPath = Path.Combine(_programDownloadFolder, "AutoDownloadedPrograms",DateTime.Now.ToString("dd-MMM-yyyy"));


            string masterProgramPath = Path.Combine(masterProgramFolderPath, "O" + programNumber + mainProgramCNCComment + ".txt");
            string autoDownloadMainProgramPath = Path.Combine(autoDownloadedProgramPath, "O" + programNumber + mainProgramCNCComment, "O" + programNumber + mainProgramCNCComment + DateTime.Now.ToString("_yyyyMMddHHmm")+ ".txt");
         
            if (Directory.Exists(masterProgramFolderPath) && File.Exists(masterProgramPath))
            {
                mainProgramMasterStr = ReadFileContent(masterProgramPath);
                //compare main CNC and main master program
                bool isMainProgramsSame = CompareContents(mainProgramCNCStr, mainProgramMasterStr);
                if (isMainProgramsSame == false)
                {
                    //save programs to autodownload folder
                    if (_AutoDownloadEveryTimeIfNotSameAsMaster)
                    {
                        Logger.WriteDebugLog("Main program not same, saving to autodownload folder");
                        CreateDirectory(Path.GetDirectoryName(autoDownloadMainProgramPath));                        
                        WriteFileContent(autoDownloadMainProgramPath, mainProgramCNCStr);
                        //TODO - downloaded program not same as Master Program - Log to MessageHistory Table
                        messageForSMS.AppendLine(" O" + programNumber + mainProgramCNCComment);
                        messageForSMS.AppendLine("File : " + Path.GetFileNameWithoutExtension(autoDownloadMainProgramPath));

                    }
                    else
                    {
                        //check the previous saved program
                        if (_AutoDownloadedSavedPrograms.ContainsKey(programNumber))
                        {                            
                            isMainProgramsSame = CompareContents(mainProgramCNCStr, _AutoDownloadedSavedPrograms[programNumber].ToString());
                            if (isMainProgramsSame == false)
                            {
                                Logger.WriteDebugLog("Main program not same with previous version of autodownload programs, saving to autodownload folder");
                                CreateDirectory(Path.GetDirectoryName(autoDownloadMainProgramPath));
                                WriteFileContent(autoDownloadMainProgramPath, mainProgramCNCStr);
                                _AutoDownloadedSavedPrograms[programNumber] = mainProgramCNCStr;
                                //TODO - downloaded program not same as Master Program - Log to MessageHistory Table
                                messageForSMS.AppendLine(" O" + programNumber + mainProgramCNCComment);
                                messageForSMS.AppendLine("File : " + Path.GetFileNameWithoutExtension(autoDownloadMainProgramPath));
                            }
                        }
                        else
                        {
                            Logger.WriteDebugLog("Main program not same with previous version of autodownload programs, saving to autodownload folder");
                            CreateDirectory(Path.GetDirectoryName(autoDownloadMainProgramPath));
                            _AutoDownloadedSavedPrograms.Add(programNumber, mainProgramCNCStr);
                            WriteFileContent(autoDownloadMainProgramPath, mainProgramCNCStr);
                            //TODO - downloaded program not same as Master Program - Log to MessageHistory Table
                            messageForSMS.AppendLine(" O" + programNumber + mainProgramCNCComment);
                            messageForSMS.AppendLine("File : " + Path.GetFileNameWithoutExtension(autoDownloadMainProgramPath));
                        }
                    }                    
                }
                else
                {
                    //save programs to autodownload folder
                    if (_AutoDownloadEveryTimeIfNotSameAsMaster)
                    {
                        Logger.WriteDebugLog("Program same as Master program, saving to autodownload folder");
                        CreateDirectory(Path.GetDirectoryName(autoDownloadMainProgramPath));
                        WriteFileContent(autoDownloadMainProgramPath, mainProgramCNCStr);                      

                    }
                }
                foreach (int item in hashSubPrograms.Keys)
                {                   
                    string subProgramMasterFile = Path.Combine(masterProgramFolderPath, "O" + programNumber + mainProgramCNCComment+ "_O" + item + ".txt");
                    string subProgramAutoDownloadFile = Path.Combine(autoDownloadedProgramPath, "O" + programNumber + mainProgramCNCComment, "O" + programNumber + mainProgramCNCComment + "_O" + item + DateTime.Now.ToString("_yyyyMMddHHmm") + ".txt");
                    if (File.Exists(subProgramMasterFile))
                    {
                        //compaire programs, if not match save it
                        bool isSubProgramSame = CompareContents(ReadFileContent(subProgramMasterFile), hashSubPrograms[item].ToString());
                        if (isSubProgramSame == false)
                        {
                            //check the previous saved program
                            if (_AutoDownloadEveryTimeIfNotSameAsMaster)
                            {
                                CreateDirectory(Path.GetDirectoryName(subProgramAutoDownloadFile));
                                WriteFileContent(subProgramAutoDownloadFile, hashSubPrograms[item].ToString());
                                //TODO - downloaded program not same as Master Program - Log to MessageHistory Table
                                messageForSMS.AppendLine(" " + Path.GetFileNameWithoutExtension(subProgramMasterFile));
                                messageForSMS.AppendLine("File : " + Path.GetFileNameWithoutExtension(subProgramAutoDownloadFile));
                            }
                            else
                            {
                                //check the previous saved program
                                if (_AutoDownloadedSavedPrograms.ContainsKey(item))
                                {
                                    isSubProgramSame = CompareContents(hashSubPrograms[item].ToString(), _AutoDownloadedSavedPrograms[item].ToString());
                                    if (isSubProgramSame == false)
                                    {
                                        CreateDirectory(Path.GetDirectoryName(subProgramAutoDownloadFile));
                                        WriteFileContent(subProgramAutoDownloadFile, hashSubPrograms[item].ToString());
                                        _AutoDownloadedSavedPrograms[item] = hashSubPrograms[item].ToString();
                                        //TODO - downloaded program not same as Master Program - Log to MessageHistory Table
                                        messageForSMS.AppendLine(" " + Path.GetFileNameWithoutExtension(subProgramMasterFile));
                                        messageForSMS.AppendLine("File : " + Path.GetFileNameWithoutExtension(subProgramAutoDownloadFile));
                                    }
                                }
                                else
                                {
                                    CreateDirectory(Path.GetDirectoryName(subProgramAutoDownloadFile));
                                    _AutoDownloadedSavedPrograms.Add(item, hashSubPrograms[item].ToString());
                                    WriteFileContent(subProgramAutoDownloadFile, hashSubPrograms[item].ToString());
                                    //TODO - downloaded program not same as Master Program - Log to MessageHistory Table
                                    messageForSMS.AppendLine(" " + Path.GetFileNameWithoutExtension(subProgramMasterFile));
                                    messageForSMS.AppendLine("File : " + Path.GetFileNameWithoutExtension(subProgramAutoDownloadFile));
                                }
                            }
                        }
                    }
                    else
                    {
                        //save to Master folder
                        Logger.WriteDebugLog("Master sub program created for : " + item);
                        WriteFileContent(subProgramMasterFile, hashSubPrograms[item].ToString());                        
                    }
                }
            }
            else
            {
                //main program not exists, save all programs to master folder  
                Logger.WriteDebugLog(string.Format( "Main program {0} not exists, save master and sub programs to master folder",programNumber));
                CreateDirectory(masterProgramFolderPath);                
                //write the programs to folder if containt not same...
                masterProgramPath = Path.Combine(masterProgramFolderPath, "O" + programNumber + mainProgramCNCComment + ".txt");
                WriteFileContent(masterProgramPath, mainProgramCNCStr);
                foreach (int item in hashSubPrograms.Keys)
                {
                    string subProgramMasterFile = Path.Combine(masterProgramFolderPath, "O" + programNumber + mainProgramCNCComment + "_O" + item + ".txt");                   
                    WriteFileContent(subProgramMasterFile, hashSubPrograms[item].ToString());             
                }
            }
            if (enableSMSforProgramChange && messageForSMS.Length > 0)
            {
                //messageForSMS.Insert(0, "Program Change Alert : " + this.machineId);
                Logger.WriteDebugLog(messageForSMS.ToString());              
                DatabaseAccess.InsertAlertNotificationHistory(this.machineId, messageForSMS.ToString());
            }
          
            return 0;
        }

        private int DownloadProgramDASCNC(string machineId, string ipAddress, ushort port, string programNumber)
        {
            //download running program in temp folder
            bool result = false;
            int programNo = 0;
            programNumber = programNumber.TrimStart(new char[] { 'O' }).Trim();
            int.TryParse(programNumber, out programNo);
            if (programNo == 0) return 0;

            //TODO - get program full path
            string programFolderPath = string.Empty;
            string programFolderFullPath_Path2 = string.Empty;
            string programFolderFullPath = FocasData.ReadFullProgramPathRunningProgram(ipAddress, port);
            if (!string.IsNullOrEmpty(programFolderFullPath))
            {
                programFolderFullPath_Path2 = programFolderFullPath.Replace("PATH1", "PATH2");
                programFolderPath = Directory.GetParent(programFolderFullPath).Name;
            }
            bool isProgramFolderSupports = string.IsNullOrEmpty(programFolderFullPath) ? false : true;
            string programFolderPath_PATH2 = "PATH2";


            string mainProgramMasterStr = string.Empty;
            StringBuilder messageForSMS = new StringBuilder();
            Logger.WriteDebugLog("Downloading main program : O" + programNo);
            string mainProgramCNCStr = FocasData.DownloadProgram(ipAddress, port, programNo, out result, programFolderFullPath, isProgramFolderSupports);
            if (!result) return -1;
            Hashtable hashSubPrograms = new Hashtable();
            //check if main program contains sub peogram(contains with M98P)
            string mainProgramCNCComment = FindProgramComment(mainProgramCNCStr);
            //List<int> subProgramsCNC = FindSubPrograms(mainProgramCNCStr);

            List<int> subProgramsCNC = FindSubProgramsDASCNC(mainProgramCNCStr);
            List<int> subProgramsCNCTemp = new List<int>();
            if (subProgramsCNC.Count > 0)
            {
                //download sub programs starts with M98P
                foreach (var item in subProgramsCNC)
                {
                    Logger.WriteDebugLog("Downloading first level sub program : " + item);//TODO - PATH2
                    string prgText = FocasData.DownloadProgram(ipAddress, port, item, out result, programFolderFullPath_Path2, isProgramFolderSupports);
                    if (result)
                    {
                        hashSubPrograms.Add(item, prgText);//TODO - PATH2
                        //find second level sub programs
                        subProgramsCNCTemp.AddRange(FindSubProgramsDASCNC(prgText));
                    }
                    //else
                    //    return -1;
                }
            }

            //download second level sub programs
            if (subProgramsCNCTemp.Count > 0)
            {
                //download sub programs starts with M98P
                foreach (var item in subProgramsCNCTemp.Distinct())
                {
                    Logger.WriteDebugLog("Downloading second level sub program : " + item);//TODO - PATH2
                    string prgText = FocasData.DownloadProgram(ipAddress, port, item, out result, programFolderFullPath_Path2, isProgramFolderSupports);
                    if (result)
                    {
                        if (!hashSubPrograms.ContainsKey(item))//TODO - PATH2
                        {
                            hashSubPrograms.Add(item, prgText);//TODO - PATH2
                        }
                    }
                    //else
                    //    return -1;
                }
            }

            //compare main, sub program with Master folder program --> if not same, store the main and sub program with date time
            //O1234_yyyymmddhhmm.txt , O1234_567_yyyymmddhhmm.txt, O1234_678_yyyymmddhhmm.txt           
            CreateDirectory(_programDownloadFolder);



            //compaire the containt of main and sub program from "Master" folder under machine folder??
            string masterProgramFolderPath = Path.Combine(_programDownloadFolder, "MasterPrograms", "O" + programNumber + mainProgramCNCComment);
            string autoDownloadedProgramPath = Path.Combine(_programDownloadFolder, "AutoDownloadedPrograms", DateTime.Now.ToString("dd-MMM-yyyy"));


            string masterProgramPath = Path.Combine(masterProgramFolderPath, programFolderPath + "_O" + programNumber + mainProgramCNCComment + ".txt");
            string autoDownloadMainProgramPath = Path.Combine(autoDownloadedProgramPath, "O" + programNumber + mainProgramCNCComment,programFolderPath + "_O" + programNumber + mainProgramCNCComment + DateTime.Now.ToString("_yyyyMMddHHmm") + ".txt");

            if (Directory.Exists(masterProgramFolderPath) && File.Exists(masterProgramPath))
            {
                mainProgramMasterStr = ReadFileContent(masterProgramPath);
                //compare main CNC and main master program
                bool isMainProgramsSame = CompareContents(mainProgramCNCStr, mainProgramMasterStr);
                if (isMainProgramsSame == false)
                {
                    //save programs to autodownload folder
                    if (_AutoDownloadEveryTimeIfNotSameAsMaster)
                    {
                        Logger.WriteDebugLog("Main program not same, saving to autodownload folder");
                        CreateDirectory(Path.GetDirectoryName(autoDownloadMainProgramPath));
                        WriteFileContent(autoDownloadMainProgramPath, mainProgramCNCStr);
                        //TODO - downloaded program not same as Master Program - Log to MessageHistory Table
                        messageForSMS.AppendLine(programFolderPath + "_O" + programNumber + mainProgramCNCComment);
                        messageForSMS.AppendLine("File : " + Path.GetFileNameWithoutExtension(autoDownloadMainProgramPath));

                    }
                    else
                    {
                        //check the previous saved program
                        if (_AutoDownloadedSavedPrograms.ContainsKey(programFolderPath + programNumber))
                        {
                            isMainProgramsSame = CompareContents(mainProgramCNCStr, _AutoDownloadedSavedPrograms[programFolderPath + programNumber].ToString());
                            if (isMainProgramsSame == false)
                            {
                                Logger.WriteDebugLog("Main program not same with previous version of autodownload programs, saving to autodownload folder");
                                CreateDirectory(Path.GetDirectoryName(autoDownloadMainProgramPath));
                                WriteFileContent(autoDownloadMainProgramPath, mainProgramCNCStr);
                                _AutoDownloadedSavedPrograms[programFolderPath + programNumber] = mainProgramCNCStr;
                                //TODO - downloaded program not same as Master Program - Log to MessageHistory Table
                                messageForSMS.AppendLine(" " + programFolderPath + "_O" + programNumber + mainProgramCNCComment);
                                messageForSMS.AppendLine("File : " + Path.GetFileNameWithoutExtension(autoDownloadMainProgramPath));
                            }
                        }
                        else
                        {
                            Logger.WriteDebugLog("Main program not same with previous version of autodownload programs, saving to autodownload folder");
                            CreateDirectory(Path.GetDirectoryName(autoDownloadMainProgramPath));
                            _AutoDownloadedSavedPrograms.Add(programFolderPath + programNumber, mainProgramCNCStr);
                            WriteFileContent(autoDownloadMainProgramPath, mainProgramCNCStr);
                            //TODO - downloaded program not same as Master Program - Log to MessageHistory Table
                            messageForSMS.AppendLine(" " + programFolderPath + "_O" + programNumber + mainProgramCNCComment);
                            messageForSMS.AppendLine("File : " + Path.GetFileNameWithoutExtension(autoDownloadMainProgramPath));
                        }
                    }
                }
                foreach (int item in hashSubPrograms.Keys)
                {
                    string subProgramMasterFile = Path.Combine(masterProgramFolderPath, "PATH2_O" + programNumber + mainProgramCNCComment + "_O" + item + ".txt");
                    string subProgramAutoDownloadFile = Path.Combine(autoDownloadedProgramPath, "O" + programNumber + mainProgramCNCComment, "PATH2_O" + programNumber + mainProgramCNCComment + "_O" + item + DateTime.Now.ToString("_yyyyMMddHHmm") + ".txt");
                    if (File.Exists(subProgramMasterFile))
                    {
                        //compaire programs, if not match save it
                        bool isSubProgramSame = CompareContents(ReadFileContent(subProgramMasterFile), hashSubPrograms[item].ToString());
                        if (isSubProgramSame == false)
                        {
                            //check the previous saved program
                            if (_AutoDownloadEveryTimeIfNotSameAsMaster)
                            {
                                CreateDirectory(Path.GetDirectoryName(subProgramAutoDownloadFile));
                                WriteFileContent(subProgramAutoDownloadFile, hashSubPrograms[item].ToString());
                                //TODO - downloaded program not same as Master Program - Log to MessageHistory Table
                                messageForSMS.AppendLine(Path.GetFileNameWithoutExtension(subProgramMasterFile));
                                messageForSMS.AppendLine("File : " + Path.GetFileNameWithoutExtension(subProgramAutoDownloadFile));
                            }
                            else
                            {
                                //check the previous saved program
                                if (_AutoDownloadedSavedPrograms.ContainsKey("PATH2" + item))
                                {
                                    isSubProgramSame = CompareContents(hashSubPrograms[item].ToString(), _AutoDownloadedSavedPrograms["PATH2" + item].ToString());
                                    if (isSubProgramSame == false)
                                    {
                                        CreateDirectory(Path.GetDirectoryName(subProgramAutoDownloadFile));
                                        WriteFileContent(subProgramAutoDownloadFile, hashSubPrograms[item].ToString());
                                        _AutoDownloadedSavedPrograms["PATH2" + item] = hashSubPrograms[item].ToString();
                                        //TODO - downloaded program not same as Master Program - Log to MessageHistory Table
                                        messageForSMS.AppendLine(Path.GetFileNameWithoutExtension(subProgramMasterFile));
                                        messageForSMS.AppendLine("File : " + Path.GetFileNameWithoutExtension(subProgramAutoDownloadFile));
                                    }
                                }
                                else
                                {
                                    CreateDirectory(Path.GetDirectoryName(subProgramAutoDownloadFile));
                                    _AutoDownloadedSavedPrograms.Add("PATH2" + item, hashSubPrograms[item].ToString());
                                    WriteFileContent(subProgramAutoDownloadFile, hashSubPrograms[item].ToString());
                                    //TODO - downloaded program not same as Master Program - Log to MessageHistory Table
                                    messageForSMS.AppendLine(" " + Path.GetFileNameWithoutExtension(subProgramMasterFile));
                                    messageForSMS.AppendLine("File : " + Path.GetFileNameWithoutExtension(subProgramAutoDownloadFile));
                                }
                            }
                        }
                    }
                    else
                    {
                        //save to Master folder
                        Logger.WriteDebugLog("Master sub program created for : " + item);
                        WriteFileContent(subProgramMasterFile, hashSubPrograms[item].ToString());
                    }
                }
            }
            else
            {
                //main program not exists, save all programs to master folder  
                Logger.WriteDebugLog(string.Format("Main program {0} not exists, save master and sub programs to master folder", programNumber));
                CreateDirectory(masterProgramFolderPath);
                //write the programs to folder if containt not same...
                masterProgramPath = Path.Combine(masterProgramFolderPath, programFolderPath + "_O" + programNumber + mainProgramCNCComment + ".txt");
                WriteFileContent(masterProgramPath, mainProgramCNCStr);
                foreach (int item in hashSubPrograms.Keys)
                {
                    string subProgramMasterFile = Path.Combine(masterProgramFolderPath, "PATH2_O" + programNumber + mainProgramCNCComment + "_O" + item + ".txt");
                    WriteFileContent(subProgramMasterFile, hashSubPrograms[item].ToString());
                }
            }
            if (enableSMSforProgramChange && messageForSMS.Length > 0)
            {
                //messageForSMS.Insert(0, "Program Change Alert : " + this.machineId);
                Logger.WriteDebugLog("Message For SMS = " + messageForSMS.ToString());
                DatabaseAccess.InsertAlertNotificationHistory(this.machineId, messageForSMS.ToString());
            }

            return 0;
        }

        private static List<int> FindSubPrograms(string programText)
        {
            List<int> programs = new List<int>();
            if (programText.Contains("M98P"))
            {
                string[] lines = programText.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                //parse the file to findout sub-programs
                foreach (var line in lines)
                {
                    if (line.Contains("M98P"))
                    {
                        string prg = line.Remove(0, line.IndexOf("M98P") + 4);
                        Regex rgx = new Regex("[a-zA-Z ]"); //Regex.Replace(prg,"[^0-9 ]","");                       
                        prg = rgx.Replace(prg, "");
                        int p;
                        if (Int32.TryParse(prg, out p))
                        {
                            if (!programs.Contains(p))
                            {
                                programs.Add(p);
                            }
                        }
                    }
                }
            }
            return programs;
        }

 		private static List<int> FindSubProgramsDASCNC(string programText)
        {
            List<int> programs = new List<int>();
            if (programText.Contains("M90"))
            {
                string[] lines = programText.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                //parse the file to findout sub-programs
                foreach (var line in lines)
                {
                    if (line.Contains("M90"))
                    {
                        string prg = line.Remove(0, line.IndexOf("M90") + 3);
                        //Regex rgx = new Regex("[a-zA-Z )]"); //Regex.Replace(prg,"[^0-9 ]","");      
                        Regex rgx = new Regex("[^0-9 ]");
                        prg = rgx.Replace(prg, "");
                        int p;
                        if (Int32.TryParse(prg, out p))
                        {
                            if (!programs.Contains(p))
                            {
                                programs.Add(p);
                            }
                        }
                    }
                }
            }
            return programs;
        }
        private static string FindProgramComment(string programText)
        {
            string comment = "(";
            string[] lines = programText.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines.ToList().Take(2))
            {
                if (line.Contains("(") && line.Contains(")"))
                {
                    comment += line.Substring(line.IndexOf("(") + 1, line.IndexOf(")") - line.IndexOf("(") - 1);
                    break;
                }
            }
            return Utility.SafeFileName(comment + ")");
        }

        private static string FindProgramNumberAndComment(string programText, out int programNumber)
        {
            string comment = "(";
            programNumber = 0;
            string[] lines = programText.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines.ToList().Take(2))
            {
                if (line.Contains("O"))
                {
                    string prog = line;
                    if (line.Contains("("))
                    {
                        prog = prog.Substring(prog.IndexOf("O") + 1, prog.IndexOf("(") - 1);
                    }
                    else
                    {
                        Regex rgx = new Regex("[a-zA-Z() ]"); //Regex.Replace(prg,"[^0-9 ]","");                       
                        prog = rgx.Replace(prog, "");
                    }
                    int p;
                    if (Int32.TryParse(prog, out p))
                    {
                        programNumber = p;
                    }
                    break;
                }
            }
            foreach (var line in lines.ToList().Take(2))
            {
                if (line.Contains("(") && line.Contains(")"))
                {
                    comment += line.Substring(line.IndexOf("(") + 1, line.IndexOf(")") - line.IndexOf("(") - 1);
                    break;
                }
            }
            return Utility.SafeFileName(comment + ")");
        }


        private static bool CompareContents(string str1, string str2)
        {
            if (str1.Equals(str2, StringComparison.OrdinalIgnoreCase))
                return true;

            return false;
        }

        private static string ReadFileContent(string filePath)
        {
            try
            {
                return File.ReadAllText((filePath));
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }

            return string.Empty;
        }

        private static bool WriteFileContent(string filePath, string str)
        {
            try
            {
 				if (!Directory.Exists(Path.GetDirectoryName(filePath)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(filePath));
                }
                File.WriteAllText((filePath), str);
                return true;
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }

            return false;
        }
       
        public static string SafePathName(string name)
        {
            StringBuilder str = new StringBuilder( name);           

            foreach (char c in System.IO.Path.GetInvalidPathChars())
            {
                str = str.Replace(c, '_');
            }
            return str.ToString();
        }

        public static bool CreateDirectory(string masterProgramFolderPath)
        {
            var safeMasterProgramFolderPath = SafePathName(masterProgramFolderPath);
             if (!Directory.Exists(safeMasterProgramFolderPath))
            {
                try
                {
                    Directory.CreateDirectory(safeMasterProgramFolderPath);
                }
                catch(Exception ex)
                {
                    Logger.WriteErrorLog(ex.ToString());
                    return false;
                }
            }
            return true;
        }

        private int get_alarm_type(int n)
        {
            int i, res = 0;
            for (i = 0; i < 32; i++)
            {
                int n1 = n;

                res = (int)(n1 & (1 << i));
                if (res != 0)
                {
                    return (i);
                }
            }
            if (i == 32)
            {
                return -1;
            }
            return -1;
        }

        public void GetAlarmsData(Object stateObject)
        {
            if (!_isLicenseValid) return;
            if (Monitor.TryEnter(_lockerAlarmHistory, 100))
            {               
                try
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    Thread.CurrentThread.Name = "AlarmsHistory-" + Utility.SafeFileName(this.machineId);

                    if (Utility.CheckPingStatus(this.ipAddress))
                    {
                        CheckMachineType();
                        Logger.WriteDebugLog("Reading Alarms History data for control type." + _cncMachineType.ToString());
                        if (_cncMachineType == CncMachineType.cncUnknown) return;
                        DataTable dt = default(DataTable);
                        if (_cncMachineType == CncMachineType.Series300i ||
                            _cncMachineType == CncMachineType.Series310i ||
                            _cncMachineType == CncMachineType.Series320i ||
                            _cncMachineType == CncMachineType.Series0i)
                        {
                            dt = FocasData.ReadAlarmHistory(machineId, ipAddress, portNo);
                        }
                        else
                        {
                            //oimc,210i
                            dt = FocasData.ReadAlarmHistory18i(machineId, ipAddress, portNo);
                        }
                        DatabaseAccess.InsertAlarms(dt, machineId);
                        Logger.WriteDebugLog("Completed reading Alarms History data.");
                    }
                }
                catch (Exception ex)
                {
                    Logger.WriteDebugLog(ex.ToString());
                }
                finally
                {                   
                    Monitor.Exit(_lockerAlarmHistory);
                               
                }
            }

        }

        public void GetAlarmsDataforEndTimeUpdate()
        {           
            try
            {

                if (Utility.CheckPingStatus(this.ipAddress))
                {
                    CheckMachineType();
                    Logger.WriteDebugLog("Reading Alarms History data to update the ALARM END TIME for control type." + _cncMachineType.ToString());
                    if (_cncMachineType == CncMachineType.cncUnknown) return;
                    DataTable dt = default(DataTable);
                    if (_cncMachineType == CncMachineType.Series300i ||
                        _cncMachineType == CncMachineType.Series310i ||
                        _cncMachineType == CncMachineType.Series320i ||
                        _cncMachineType == CncMachineType.Series0i)
                    {
                        dt = FocasData.ReadAlarmHistory(machineId, ipAddress, portNo);
                    }
                    else
                    {
                        //oimc,210i
                        dt = FocasData.ReadAlarmHistory18i(machineId, ipAddress, portNo);
                    }
                    DatabaseAccess.InsertAlarms(dt, machineId);
                    Logger.WriteDebugLog("Completed reading Alarms History data.");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteDebugLog(ex.ToString());
            }
            finally
            {              
               
            }

        }

        public void GetOperationHistoryData(Object stateObject)
        {
            if (!_isLicenseValid) return;
            if (Monitor.TryEnter(_lockerOperationHistory, 100))
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
              
                try
                {
                    Thread.CurrentThread.Name = "OperationHistory-" + Utility.SafeFileName(this.machineId);

                    if (Utility.CheckPingStatus(this.ipAddress))
                    {
                        Logger.WriteDebugLog("Reading Operation History data for control type." + _cncMachineType.ToString());
                        string FilePath = Path.Combine(_operationHistoryFolderPath, this.machineId, DateTime.Now.ToString("yyyy-MM-dd"));
                        if (!Directory.Exists(FilePath))
                        {
                            try
                            {
                                Directory.CreateDirectory(FilePath);
                            }
                            catch { }
                        }
                        string fileName = Path.Combine(FilePath, DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt");
                        short OperationHistoryFlagLocation = 0;
                        short dprint_flagLocation = 0;
                        int dprint_flagValue = 0;
                        if (!short.TryParse(ConfigurationManager.AppSettings["OperationHistory_FlagLocation"].ToString(), out OperationHistoryFlagLocation))
                        {
                            OperationHistoryFlagLocation = 0;
                        }
                        if (!short.TryParse(ConfigurationManager.AppSettings["DPRINT_FlagLocation"].ToString(), out dprint_flagLocation))
                        {
                            dprint_flagLocation = 0;
                        }


                        try
                        {
                            if (OperationHistoryFlagLocation > 0)
                            {
                                FocasData.UpdateOperatinHistoryMacroLocation(this.ipAddress, this.portNo, OperationHistoryFlagLocation, 1);
                            }
                            if (dprint_flagLocation > 0)
                            {
                                dprint_flagValue = FocasData.ReadOperatinHistoryDPrintLocation(this.ipAddress, this.portNo, dprint_flagLocation);
                            }
                            if (dprint_flagValue == 0)
                            {
                                FocasData.DownloadOperationHistory(this.ipAddress, this.portNo, fileName);
                            }
                            if (OperationHistoryFlagLocation > 0)
                            {
                                FocasData.UpdateOperatinHistoryMacroLocation(this.ipAddress, this.portNo, OperationHistoryFlagLocation, 0);
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.WriteErrorLog(ex.ToString());
                        }
                    }                   
                }
                catch (Exception ex)
                {
                    Logger.WriteDebugLog(ex.ToString());
                }
                finally
                {                   
                    Monitor.Exit(_lockerOperationHistory);                                    
                }
            }
        }


        public void GetTPMTrakStringData(object stateObject)
        {
            if (!_isLicenseValid) return;
            if (Monitor.TryEnter(this._lockerTPMTrakDataCollection, 100))
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                ushort focasLibHandle = 0;
                string text = string.Empty;               
                try
                {
                    Thread.CurrentThread.Name = "TPMTrakDataCollation-" + this.machineId;
                   
                    if ( Utility.CheckPingStatus(this.ipAddress))
                    {
                        Logger.WriteDebugLog("Reading Production data. ");
                        int ret = (int)FocasData.cnc_allclibhndl3(this.ipAddress, this.portNo, 10, out focasLibHandle);
                        if (ret == 0)
                        {
                            //string mode = FocasData.ReadMachineMode(focasLibHandle);
                            //if (mode.Equals("MEM", StringComparison.OrdinalIgnoreCase))
                            {
                                List<TPMString> list = new List<TPMString>();
                                foreach (TPMMacroLocation current in this.setting.TPMDataMacroLocations)
                                {
                                    int isDataReadyToread = FocasData.ReadMacro(focasLibHandle, current.StatusMacro);
                                    if (isDataReadyToread > 0)
                                    {
                                        Logger.WriteDebugLog("Reading Production data. Data read Macro location is high. = " + current.StatusMacro);
                                        List<int> values = FocasData.ReadMacroRange(focasLibHandle, current.StartLocation, current.EndLocation);
                                        TPMString tPMString = new TPMString();
                                        tPMString.Seq = values[0];
                                        text = this.BuildString(values);
                                        tPMString.TpmString = text;
                                        this.SaveStringToTPMFile(text);
                                        //tPMString.DateTime = this.GetDatetimeFromtpmString(values);

                                        list.Add(tPMString);
                                        FocasData.WriteMacro(focasLibHandle, current.StatusMacro, 0);
                                    }
                                }
                                foreach (TPMString current2 in list.OrderBy(s => s.Seq))
                                {
                                    this.ProcessData(current2.TpmString, this.ipAddress, this.portNo.ToString(), this.machineId);
                                }
                            }
                        }
                        else
                        {
                            Logger.WriteErrorLog("Not able to connect to CNC machine. ret value from fun cnc_allclibhndl3 = " + ret);
                        }
                    }                    
                }
                catch (Exception ex)
                {
                    Logger.WriteDebugLog(ex.ToString());
                }
                finally
                {
                    if (focasLibHandle != 0)
                    {
                        var r = FocasData.cnc_freelibhndl(focasLibHandle);
                        //if (r != 0) _focasHandles.Add(focasLibHandle);
                    }                    
                    Monitor.Exit(this._lockerTPMTrakDataCollection);

                }
            }
        }      

        public void GetCycleTimeData(object stateObject)
        {
            if (!_isLicenseValid) return;
            if (Monitor.TryEnter(this._lockerCycletimeReader, 1000))
            {
                ushort focasLibHandle = 0;
                string text = string.Empty;
              
                try
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    Thread.CurrentThread.Name = "TPMTrakCycleTimeDataCollation-" + this.machineId;

                    if (Utility.CheckPingStatus(this.ipAddress))
                    {
                        int ret = (int)FocasData.cnc_allclibhndl3(this.ipAddress, this.portNo, 10, out focasLibHandle);
                        if (ret == 0)
                        {
                            
                                int isDataReadyToread = FocasData.ReadMacro(focasLibHandle, (short) (cycleTimeMacroLocation + 5));
                                if (isDataReadyToread > 0)
                                {
                                    List<int> values = FocasData.ReadMacroRange(focasLibHandle, cycleTimeMacroLocation,(short) (cycleTimeMacroLocation + 4));
                                    FocasData.WriteMacro(focasLibHandle, (short)(cycleTimeMacroLocation + 5), 0);
                                    DatabaseAccess.InsertCycleTimeData(values, this.machineId);
                                }                               
                        }
                        else
                        {
                            Logger.WriteErrorLog("Not able to connect to CNC machine. ret value from fun cnc_allclibhndl3 = " + ret);
                        }
                    }                   
                }
                catch (Exception ex)
                {
                    Logger.WriteDebugLog(ex.ToString());
                }
                finally
                {
                    if (focasLibHandle != 0)
                    {
                        var r = FocasData.cnc_freelibhndl(focasLibHandle);                      
                    }                  
                    Monitor.Exit(this._lockerCycletimeReader);

                }
            }
        }

        private DateTime GetDatetimeFromtpmString(List<int> values)
        {
            string[] formats = new string[]
			{
				"yyyyMMdd HHmmss"
			};
            DateTime minValue = DateTime.MinValue;
            var date = values[values.Count - 2];
            var time = values[values.Count - 1];
            if (!DateTime.TryParseExact(date + " " + time, formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out minValue))
            {
                string strDate = Utility.get_actual_date(date);
                string strTime = Utility.get_actual_time(time);
                DateTime.TryParse(strDate + " " + strTime, out minValue);   
            }
            return minValue;
        }
        
        private string BuildString(List<int> values)
        {
            if (values[1] == 11 || values[1] == 1 || values[1]==22 || values[1] == 2)
            {
                return string.Format("START-{0}-{1}-{2}-{3}-{4}-{5}-{6}-{7}-{8}-{9}-END-{10}", values[1], values[2],
                                           values[3], values[4], values[5], values[6], values[7], values[8], values[9], values[10], values[0]);               
            }
            else
            {
                return string.Format("START-{0}-{1}-{2}-{3}-{4}-{5}-{6}-{7}-{8}-{9}-{10}-{11}-END-{12}", values[1], values[2],
                                                values[3], values[4], values[5], values[6], values[7], values[8], values[9], values[10], values[11], values[12], values[0]);
            }            
        }

        private string BuildStringOffset(List<int> values)
        {
            return string.Format("START-{0}-{1}-{2}-{3}-{4}-{5}-{6}-{7}-{8}-{9}-END", values[0], values[1], values[2],
                                            values[3], values[4], values[5], values[6], values[7], values[8], values[9].ToString("000000"));

            //START-Data type-MachineID-PartID-Operation-Tool no-Edge no-Target-Actual-Date-Time-END
        }

        private string BuildInspection37String(string mc, string comp, string opn, SPCCharacteristics spc, DateTime cncTime)
        {
           //START-37-MC-COMP-OPRN-Featureid-DIMENSIONid-<VALUE>-DATE-TIME-END 
            return string.Format("START-37-{0}-{1}-{2}-{3}-{4}-@{5}/-{6}-{7}-END", mc, comp,opn, spc.FeatureID, spc.DiamentionId, spc.DiamentionValue, cncTime.ToString("yyyyMMdd"), cncTime.ToString("HHmmss") );
        }

        private void SaveStringToTPMFile(string str)
        {
            string progTime = String.Format("_{0:yyyyMMdd}", DateTime.Now);

            StreamWriter writer = default(StreamWriter);
            try
            {
                writer = new StreamWriter(appPath + "\\TPMFiles\\F-" + Utility.SafeFileName(Thread.CurrentThread.Name + progTime) + ".tpm", true);
                writer.WriteLine(str);
                writer.Flush();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (writer != null)
                {
                    writer.Close();
                    writer.Dispose();
                }
            }
        }

        public void WriteInToFileDBInsert(string str)
        {
            string progTime = String.Format("_{0:yyyyMMdd}", DateTime.Now);
            string location = appPath + "\\Logs\\DBInsert-" + Utility.SafeFileName( MName + progTime) + ".txt";

            StreamWriter writer = default(StreamWriter);
            try
            {
                writer = new StreamWriter(location, true, Encoding.Default, 8195);
                writer.WriteLine(str);
                writer.Flush();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (writer != null)
                {
                    writer.Close();
                    writer.Dispose();
                }
            }
        }

        public void ProcessData(string InputStr, string IP, string PortNo, string MName)
        {
            try
            {
                string ValidString = FilterInvalids(InputStr);
                WriteInToFileDBInsert(string.Format("{0} : Start Insert Record - {1} ; IP = {2}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:FFF"), ValidString, IP));
                InsertDataUsingSP(ValidString, IP, PortNo);
                WriteInToFileDBInsert(string.Format("{0} : Stop Insert - {1}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:FFF"), IP));
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("ProcessFile() :" + ex.ToString());
            }
            return;
        }

        public static string FilterInvalids(string DataString)
        {
            string FilterString = string.Empty;
            try
            {
                for (int i = 0; i < DataString.Length; i++)
                {
                    byte[] asciiBytes = Encoding.ASCII.GetBytes(DataString.Substring(i, 1));

                    if (asciiBytes[0] >= Encoding.ASCII.GetBytes("#")[0] && asciiBytes[0] <= Encoding.ASCII.GetBytes("}")[0])  //to handle STR   -1-0111-000000001-1-0002-1-20110713-175258914-20110713-175847898-END more than 2 spaces in string
                    {
                        FilterString = FilterString + DataString.Substring(i, 1);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            return FilterString;
        }

        public static int InsertDataUsingSP(string DataString, string IP, string PortNo)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("s_GetProcessDataString", Con);
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.Add("@datastring", SqlDbType.NVarChar).Value = DataString;
            cmd.Parameters.Add("@IpAddress", SqlDbType.NVarChar).Value = IP;
            cmd.Parameters.Add("@OutputPara", SqlDbType.Int).Value = 0;
            cmd.Parameters.Add("@LogicalPortNo", SqlDbType.SmallInt).Value = PortNo;
            int OutPut = 0;
            try
            {
                OutPut = cmd.ExecuteNonQuery();
                if (OutPut < 0)
                {
                    Logger.WriteErrorLog(string.Format("InsertDataUsingSP() - ExecuteNonQuery returns < 0 value : {0} :- {1}", IP, DataString));
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("InsertDataUsingSP():" + ex.Message);
            }
            finally
            {
                if (Con != null) Con.Close();
                cmd = null;
                Con = null;
            }
            return OutPut;
        }          
          
        public void CloseTimer()
        {
            if (_timerAlarmHistory != null) _timerAlarmHistory.Dispose();
            if (_timerOperationHistory != null) _timerOperationHistory.Dispose();
            if (_timerSpindleLoadSpeed != null) _timerSpindleLoadSpeed.Dispose();
            if (_timerPredictiveMaintenanceReader != null) _timerPredictiveMaintenanceReader.Dispose();
            if (this._timerTPMTrakDataCollection != null) this._timerTPMTrakDataCollection.Dispose();
            if (this._timerOffsetHistoryReader != null) this._timerOffsetHistoryReader.Dispose();
            if (this._timerCycletimeReader != null) this._timerCycletimeReader.Dispose();
            if (_timerToolLife != null) this._timerToolLife.Dispose();
        }   

        public void CheckMachineType()
        {
            if (_cncSeries.Equals(string.Empty))
            {
                ushort focasLibHandle = ushort.MinValue;
                short ret = FocasData.cnc_allclibhndl3(ipAddress, portNo, 4, out focasLibHandle);
                if (ret == 0)
                {
                    if (FocasData.GetFanucMachineType(focasLibHandle, ref _cncMachineType, out _cncSeries) != 0)
                    {
                        Logger.WriteErrorLog("Failed to get system info. method failed cnc_sysinfo()");
                    }
                    Logger.WriteDebugLog("CNC control type  = " + _cncMachineType.ToString() + " , " + _cncSeries);
                }
                ret = FocasData.cnc_freelibhndl(focasLibHandle);
                //if (ret != 0) _focasHandles.Add(focasLibHandle);
            }
        }       

        public short WriteWearOffsetToCNC(decimal value, short offsetLocation, ushort focasLibHandle)
        {
            try
            {
                return FocasData.WriteWearOffset2(focasLibHandle, offsetLocation, value);

            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            return 0;
        }

        public double ReadWearOffsetFromCNC(short offsetLocation, ushort focasLibHandle)
        {
            double offsetValue = double.NaN;
            try
            {
                //wear offset read write( only 3 decimal places) cnc_wrtofs( h, tidx, 0, 8, offset ) ;
                offsetValue = FocasData.ReadWearOffset2(focasLibHandle, offsetLocation);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            return offsetValue;
        }

        public short WriteWearOffsetToCNC_TEST(decimal value, short offsetLocation, ushort focasLibHandle)
        {
            try
            {
                return FocasData.WriteWearOffset_TEST(focasLibHandle, offsetLocation, value);

            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            return 0;
        }

        public void GetSpindleLoadSpeedData(Object stateObject)
        {
            if (!_isLicenseValid) return;
            if (Monitor.TryEnter(_lockerSpindleLoadSpeed, 100))
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
               
                try
                {
                    Thread.CurrentThread.Name = "SpindleLoadSpeed-" + this.machineId;

                    if (Utility.CheckPingStatus(this.ipAddress))
                    {
                        CheckMachineType();
                        Logger.WriteDebugLog("Reading Spindle Load Speed History data for control type." + _cncMachineType.ToString());
                        if (_cncMachineType == CncMachineType.cncUnknown) return;
                        //ProcessSpindleLoadSpeed(this.machineId, this.ipAddress, this.portNo, setting.SpeedLocationStart, setting.SpeedLocationEnd,setting.LoadLocationStart,setting.LoadLocationEnd);
                        ProcessSpindleLoadSpeed(this.machineId, this.ipAddress, this.portNo);
                        Logger.WriteDebugLog("Completed reading Spindle Load Speed History data.");
                    }
                }
                catch (Exception ex)
                {
                    Logger.WriteDebugLog(ex.ToString());
                }
                finally
                {                   
                    Monitor.Exit(_lockerSpindleLoadSpeed);
                }
            }

        }

        private void ProcessSpindleLoadSpeed(string machineId, string ipAddress, ushort portNo)
        {
            try
            {
                int ret = 0;
                ushort focasLibHandle = 0;

                ret = FocasData.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_allclibhndl3() failed. return value is = " + ret);
                    Thread.Sleep(1000);
                    return;
                }

                short AXIS = setting.SpindleAxisNumber == 0 ? (short)1 : setting.SpindleAxisNumber;
                if (AXIS == 1)
                {
                    var spindle = new SpindleSpeedLoadDTO();
                    spindle.CNCTimeStamp = FocasData.ReadCNCTimeStamp(focasLibHandle);
                    spindle.MachineId = machineId;
                    spindle.SpindleLoad = FocasData.ReadSpindleLoad(focasLibHandle);
                    spindle.Temperature = FocasData.ReadSpindleMotorTemp(focasLibHandle);
                    FocasLibBase.ODBDY2_1 dynamic_data = FocasData.cnc_rddynamic2(focasLibHandle);
                    spindle.ProgramNo = "O" + dynamic_data.prgmnum.ToString();
                    spindle.ToolNo = (FocasData.ReadToolNo(focasLibHandle) / 100).ToString();

                    spindle.SpindleSpeed = dynamic_data.acts;
                    spindle.FeedRate = dynamic_data.actf;
                    spindle.SpindleTarque = 0;
                    spindle.AxisNo = "X";
                    _spindleInfoQueue.Add(spindle);
                    if (_spindleInfoQueue.LongCount() > 4)
                    {
                        DataTable dt = _spindleInfoQueue.ToDataTable<SpindleSpeedLoadDTO>();
                        Thread thread = new Thread(() =>
                        DatabaseAccess.InsertBulkRows(dt, "[dbo].[Focas_SpindleInfo]"));
                        thread.Start();
                        _spindleInfoQueue.Clear();
                    }

                    if (focasLibHandle > 0)
                    {
                        var r = FocasData.cnc_freelibhndl(focasLibHandle);
                       // if (r != 0) _focasHandles.Add(focasLibHandle);
                    }
                }
                else
                {
                    FocasLibBase.ODBDY2_1 dynamic_data = FocasData.cnc_rddynamic2(focasLibHandle);
                    var programNo = "O" + dynamic_data.prgmnum.ToString();
                    var toolNo = (FocasData.ReadToolNo(focasLibHandle) / 100).ToString();

                    //FocasLibBase.ODBAXISNAME axisName = new FocasLibBase.ODBAXISNAME();
                    //short axisNumbers = 3;
                    //ret = FocasLibrary.FocasLib.cnc_rdaxisname(focasLibHandle, ref axisNumbers, axisName);
                    //char ccc = (char) axisName.data2.name;
                    List<ServoLoad> load = FocasData.ReadServoMotorLoad(focasLibHandle, 5);

                    for (short i = 1; i <= AXIS; i++)
                    {
                        var spindle = new SpindleSpeedLoadDTO();
                        spindle.CNCTimeStamp = FocasData.ReadCNCTimeStamp(focasLibHandle);
                        spindle.MachineId = machineId;
                        spindle.SpindleLoad = load.Where(l => l.AxisNumber == i).Select(a => a.Load).FirstOrDefault();
                        spindle.Temperature = FocasData.ReadServoMotorTemp(focasLibHandle, i);
                        spindle.ProgramNo = programNo;
                        spindle.ToolNo = toolNo;
                        spindle.SpindleSpeed = dynamic_data.acts;
                        spindle.FeedRate = dynamic_data.actf;
                        spindle.SpindleTarque = 0;
                        string axisName = load.Where(l => l.AxisNumber == i).Select(a => a.Axis).FirstOrDefault();
                        spindle.AxisNo = axisName != string.Empty ? axisName : i.ToString();
                        _spindleInfoQueue.Add(spindle);
                    }
                    if (_spindleInfoQueue.LongCount() > 4)
                    {
                        DataTable dt = _spindleInfoQueue.ToDataTable<SpindleSpeedLoadDTO>();
                        Thread thread = new Thread(() =>
                        DatabaseAccess.InsertBulkRows(dt, "[dbo].[Focas_SpindleInfo]"));
                        thread.Start();
                        _spindleInfoQueue.Clear();
                    }

                    if (focasLibHandle > 0)
                    {
                        var r = FocasData.cnc_freelibhndl(focasLibHandle);
                        //if (r != 0) _focasHandles.Add(focasLibHandle);
                    }
                }

            }
            catch (Exception exx)
            {
                Logger.WriteErrorLog(exx.ToString());
            }
        }

        private void ReadInspectionData(string machineId, string ipAddress, ushort portNo)
        {
            try
            {
                int ret = 0;
                ushort focasLibHandle = 0;

                ret = FocasData.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_allclibhndl3() failed during ReadInspectionData() . return value is = " + ret);
                    Thread.Sleep(1000);
                    return;
                }               

                //check new inspection data has been ready to read.
                int isDataChanged = FocasData.ReadMacro(focasLibHandle, 530);
                if (isDataChanged == 1)
                {
                    Logger.WriteDebugLog("Started reading Inspection data.");
                    int CompInterface = FocasData.ReadMacro(focasLibHandle, _CompMacroLocation);
                    int OpnInterface = 10;
                    Logger.WriteDebugLog(string.Format("Read spc data for Comp = {0} and Opn = {1}.",CompInterface,OpnInterface));
                    var CNCTimeStamp = FocasData.ReadCNCTimeStamp(focasLibHandle);
                    //read inspection value from macro location
                    var inspectionData = DatabaseAccess.GetSPC_CharacteristicsForMCO(this.interfaceId, CompInterface.ToString(), OpnInterface.ToString());
                    foreach (var item in inspectionData)
                    {
                        double inspectionValue = FocasData.ReadMacroDouble2(focasLibHandle, item.MacroLocation);
                        if (inspectionValue != double.MaxValue)
                        {
                            item.DiamentionValue = inspectionValue;
                        }
                    }                    
                   
                    //build type 37 string and insert to database
                    foreach (var item in inspectionData)
                    {
                        if (item.DiamentionValue == double.MaxValue) continue;
                        var type37String = BuildInspection37String(this.machineId, CompInterface.ToString(), "", item, CNCTimeStamp);
                        SaveStringToTPMFile(type37String);
                        ProcessData(type37String, ipAddress, portNo.ToString(), machineId);
                    }

                    //update all macro location value to '0'
                    foreach (var item in inspectionData)
                    {
                        FocasData.WriteMacro(focasLibHandle, item.MacroLocation, 0);
                    }

                    //reset the data read flag to '0'
                    FocasData.WriteMacro(focasLibHandle, 530, 2);
                    Logger.WriteDebugLog("Completed reading Inspection data.");
                }               

                if (focasLibHandle > 0)
                {
                    var r = FocasData.cnc_freelibhndl(focasLibHandle);
                    //if (r != 0) _focasHandles.Add(focasLibHandle);
                }
            }
            catch (Exception exx)
            {
                Logger.WriteErrorLog(exx.ToString());
            }
        }

        private void ReadInspectionDataAshokLeyland(string machineId, string ipAddress, ushort portNo)
        {
            try
            {
                int ret = 0;
                ushort focasLibHandle = 0;

                ret = FocasData.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
                if (ret != 0)
                {
                    Logger.WriteErrorLog("cnc_allclibhndl3() failed during ReadInspectionDataAshokLeyland() . return value is = " + ret);
                    Thread.Sleep(1000);
                    return;
                }

                //check new inspection data has been ready to read.
                int isDataChanged = FocasData.ReadMacro(focasLibHandle, _DataReadflagMacro);
                Logger.WriteErrorLog("Data Read Flag Value = " + isDataChanged);
                if (isDataChanged == 1)
                {
                    Logger.WriteDebugLog("Started reading Inspection data.");
                    int CompInterface = FocasData.ReadMacro(focasLibHandle, _ComponentMacro);
                    int featureID = FocasData.ReadMacro(focasLibHandle, _featureIDMacro);
                    int ComponentGroupIdMacroValue = 0;
                    if (_ComponentGroupIdMacro > 0)
                    {
                        ComponentGroupIdMacroValue = FocasData.ReadMacro(focasLibHandle, _ComponentGroupIdMacro);
                    }
                    //Read CycleEnd datetime from macro
                    var CNCTimeStamp = FocasData.ReadCNCTimeStamp(focasLibHandle, _CycleEndDateTimeMacro, (short)(_CycleEndDateTimeMacro +  1));
                    //read inspection value from macro location
                    var featureName = DatabaseAccess.GetFeatures(featureID);
                    var inspectionData = DatabaseAccess.GetSPC_CharacteristicsForMCOAshokaLeyland(CompInterface.ToString());
                    Logger.WriteDebugLog(string.Format("Read data : Component Interface ID = {0}, Feature ID = {1}, Feature = {2}, Cycle End TS = {3}, ComponentGroup ID Value = {4}", CompInterface, featureID, featureName, CNCTimeStamp.ToString("yyyy-MM-dd HH:mm:ss"), ComponentGroupIdMacroValue));

                    var insListToRead = inspectionData.Where(item => item.FeatureID.Equals(featureName, StringComparison.OrdinalIgnoreCase)).ToList();
                    if (insListToRead.Count == 0)
                        Logger.WriteDebugLog(string.Format("Master Data not found for Feature {0} in table SPC_Characteristic for Component interface id {1}", featureName , CompInterface));
                    foreach (var item in insListToRead)
                    {
                        double inspectionValue = FocasData.ReadMacroDouble2(focasLibHandle, item.MacroLocation);
                        if (inspectionValue != double.MaxValue)
                        {
                            item.DiamentionValue = inspectionValue;                            
                        }
                    }

                    //build type 37 string and insert to database
                    foreach (var item in insListToRead)
                    {
                        if (item.DiamentionValue == double.MaxValue || item.DiamentionValue == 0) continue;
                        var str = string.Format("START-[{0}]-[{1}]-[{2}]-[{3}]-[{4}]-[{5}]-[{6}]END", this.machineId, item.ComponentID, item.FeatureID, item.DiamentionId, item.DiamentionValue, CNCTimeStamp.ToString("yyyy-MM-dd HH:mm:ss"), ComponentGroupIdMacroValue);
                        SaveStringToTPMFile(str);
                        DatabaseAccess.SaveToDatabase(this.machineId, item.ComponentID, item.FeatureID, item.DiamentionId, item.DiamentionValue, CNCTimeStamp, ComponentGroupIdMacroValue);
                    }

                    //update all macro location value to '0'
                    //foreach (var item in inspectionData)
                    //{
                    //    FocasData.WriteMacro(focasLibHandle, item.MacroLocation, 0);
                    //}

                    //reset the data read flag to '0'
                    if(_DataReadStatusFlagMacro > 0)
                        FocasData.WriteMacro(focasLibHandle, _DataReadStatusFlagMacro, 0);
                    FocasData.WriteMacro(focasLibHandle, _DataReadflagMacro, 2);
                    Logger.WriteDebugLog("Completed reading Inspection data. Updated 2 to Macro Location " + _DataReadflagMacro);
                }

                if (focasLibHandle > 0)
                {
                    var r = FocasData.cnc_freelibhndl(focasLibHandle);
                    //if (r != 0) _focasHandles.Add(focasLibHandle);
                }
            }
            catch (Exception exx)
            {
                Logger.WriteErrorLog(exx.ToString());
            }
        }

        public void GetProcessParameterData(Object stateObject)
        {
            if (!_isLicenseValid) return;
            if (Monitor.TryEnter(_lockerProcessParameter, 100))
            {               
                try
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    Thread.CurrentThread.Name = "ProcessParameterData-" + Utility.SafeFileName(this.machineId);

                    if (Utility.CheckPingStatus(this.ipAddress))
                    {
                        ReadProcessParameterData(this.machineId, this.ipAddress, this.portNo);
                    }
                }
                catch (Exception ex)
                {
                    Logger.WriteDebugLog(ex.ToString());
                }
                finally
                {                   
                    Monitor.Exit(_lockerProcessParameter);                  
                }
            }
        }

        private void ReadProcessParameterData(string machineId, string ipAddress, ushort portNo)
        {
            try
            {
                int ret = 0;
                ushort focasLibHandle = 0;

                ret = FocasData.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
                if (ret != 0)
                {
                    Logger.WriteErrorLog("ReadProcessParameterData => cnc_allclibhndl3() failed. return value is = " + ret);
                    Thread.Sleep(1000);
                    return;
                }
              
                var CNCTimeStamp = FocasData.ReadCNCTimeStamp(focasLibHandle);

                if (this.setting.ProcessParameterSettings.Count == 0)
                {
                    Logger.WriteDebugLog(string.Format("master data not found in \"[ProcessParameterMaster_MGTL]\" table for input parameters"));
                }               

                List<ProcessParameterDTO> inputs = new List<ProcessParameterDTO>();
                foreach (var item in this.setting.ProcessParameterSettings)
                {
                    if (item.RLocation <= 0) continue;
                 
                    var result = FocasData.ReadPMCOneWord(focasLibHandle, 5, item.RLocation, (ushort)(item.RLocation + 2));
                    if (result != short.MinValue)
                    {
                        ProcessParameterDTO obj = new ProcessParameterDTO()
                        {
                            UpdatedtimeStamp = CNCTimeStamp,
                            MachineID = this.machineId,
                            ParameterID = item.ParameterID,                   
                            ParameterBitValue = result
                        };
                        inputs.Add(obj);
                    }
                }
                DatabaseAccess.InsertBulkRows(inputs.ToDataTable<ProcessParameterDTO>(), "[dbo].[ProcessParameterTransaction_MGTL]");
                Logger.WriteDebugLog("Completed reading ProcessParameter data.");


                if (focasLibHandle > 0)
                {
                    var r = FocasData.cnc_freelibhndl(focasLibHandle);
                    //if (r != 0) _focasHandles.Add(focasLibHandle);
                }
            }
            catch (Exception exx)
            {
                Logger.WriteErrorLog(exx.ToString());
            }
        }

        private void Test(ushort handle)
        {

            int ret = 0;
            ushort focasLibHandle = 0;

            /*MGTL****************************************
               short axisNo = 5;
            FocasLibBase.ODBSVLOAD aa = new FocasLibBase.ODBSVLOAD();
            ret = FocasLibrary.FocasLib.cnc_rdsvmeter(focasLibHandle, ref axisNo, aa);

            FocasLibBase.ODBSPEED b = new FocasLibBase.ODBSPEED();
            FocasLibBase.ODBACT a = new FocasLibBase.ODBACT();
            ret = FocasLibrary.FocasLib.cnc_rdspeed(focasLibHandle, -1, b);
            ret = FocasLibrary.FocasLib.cnc_actf(focasLibHandle, a);
            ***********************/


            ret = FocasData.cnc_allclibhndl3(ipAddress, portNo, 10, out focasLibHandle);
            ReadProcessParameterData(this.machineId, this.ipAddress, this.portNo);
            //FocasLibBase.ODBPTIME oDBPTIME = new FocasLibBase.ODBPTIME();
            //FocasLibrary.FocasLib.cnc_rdproctime(focasLibHandle, oDBPTIME);

            //int a = 0;
            //short b = 0;
            //FocasLibBase.PRGDIRTM c = new FocasLibBase.PRGDIRTM();
            //FocasLibrary.FocasLib.cnc_rdprgdirtime(focasLibHandle, ref a, ref b, c);

            FocasLibrary.FocasLibBase.ODBSINFO sinfo = new FocasLibBase.ODBSINFO();
            ret = FocasLibrary.FocasLib.cnc_getsraminfo(focasLibHandle, sinfo); 

            ret = FocasLibrary.FocasLib.cnc_sramgetstart(focasLibHandle, sinfo.info.info1.sramname);

            short a1 = 0;
            int length = sinfo.info.info1.sramsize;
            
            string fileName = sinfo.info.info1.fname1.Trim();
            BinaryWriter writer = new BinaryWriter(File.Open(fileName, FileMode.Create));
            while (true)
            {
                char[] buf = new char[length];
                int c1 = buf.Length;
                ret = FocasLibrary.FocasLib.cnc_sramget(focasLibHandle, out a1, buf, ref c1);
                if( ret == 0)
                 writer.Write(buf,0,c1);
                if (a1 == 0) break; //|| ret != 0
            }
            if (writer != null)
            {
                writer.Close();
                writer.Dispose();
            }
           
            //File.WriteAllText(fileName, programStr.ToString());
            ret = FocasLibrary.FocasLib.cnc_sramgetend(focasLibHandle); 


            int oo = 00;
            //this.ValidateMachineModel(this.machineId, this.ipAddress, this.portNo);


            //cnc_rdproctime  

            //cnc_rdprgdirtime

            ////insert data to datbase......
            //DataTable dt = this.machineDTO.Settings.PredictiveMaintenanceSettings.ToDataTable<PredictiveMaintenanceDTO>();
            //dt.Columns.Remove("TargetDLocation"); dt.Columns.Remove("CurrentValueDLocation");
            //Thread thread = new Thread(() =>
            //DatabaseAccess.InsertBulkRows(dt, "[dbo].[Focas_PredictiveMaintenance]"));
            //thread.Start();


            //LiveDTO dto = new LiveDTO 
            //                        {   
            //                            MachineID = this.machineId,
            //                            MachineStatus = "In Cycle",
            //                            MachineUpDownStatus = 1,
            //                            CNCTimeStamp = DateTime.Now,
            //                            MachineUpDownBatchTS = DateTime.Now,                                        
            //                            BatchTS = DateTime.Now
                                    
            //                        };
            //_liveDTOQueue.Add(dto);
            //DataTable dt = _liveDTOQueue.ToDataTable<LiveDTO>();
            //DatabaseAccess.InsertBulkRows(dt, "[dbo].[Focas_LiveData]");

            //var cool = new CoolentLubOilLevelDTO()
            //{
            //    CNCTimeStamp = DateTime.Now,
            //    MachineId = this.machineId,
            //    CoolentLevel = 3000,
            //    LubOilLevel = 2000,
            //    PrevCoolentLevel = _prevCoolentLevel,
            //    PrevLubOilLevel = _prevLubOilLevel
            //};

            //DatabaseAccess.InsertCoolentLubOilInfo(cool);
            //_prevCoolentLevel = cool.CoolentLevel;
            //_prevLubOilLevel = cool.LubOilLevel;           

            //FocasData.ReadAllPrograms(this.ipAddress, this.portNo);
            // FocasData.DownloadProgram(this.ipAddress, this.portNo,9999);
            //FocasData.UploadProgram(this.ipAddress, this.portNo, "");
            //FocasData.SearchPrograms(this.ipAddress, this.portNo, 9999);
            //FocasData.DeletePrograms(this.ipAddress, this.portNo, 9999);
            //FocasData.DownloadOperationHistory(this.ipAddress, this.portNo, 9999);
            //Thread.Sleep(1000 * 60);
            //return;

            ////////1. - read external operator messages and insert to database
            //////List<OprMessageDTO> oprMessages = FocasData.ReadExternalOperatorMessageHistory0i(this.machineId, this.ipAddress, this.portNo);            
            //////DataTable dataTable = oprMessages.ToDataTable<OprMessageDTO>();
            //////DatabaseAccess.InsertBulkRows(dataTable, "[dbo].[Focas_ExOperatorMessageTemp]");
            //////DatabaseAccess.ProcessExOperatorMessageToHistory(machineId);
            //////DatabaseAccess.DeleteExOperatorMessageTempRecords(machineId);

            ////////2.- Read D-parameter table value - used to read 10 memory location for speed, Load
            ////////database field - id, machine id, cnc time stamp, speed, load
            ////////read every 1 second

            //////DateTime CNCTimeStamp = FocasData.ReadCNCTimeStamp(handle);
            //////List<int> speedList = FocasData.ReadPMCDataTableInt(handle, 1500, 1539);
            //////List<int> loadList = FocasData.ReadPMCDataTableInt(handle, 1600, 1639);
            //////List<SpindleSpeedLoad> speedLoad = new List<SpindleSpeedLoad>();
            //////for (int i = 0; i < speedList.Count; i++)
            //////{
            //////    var spindleInfo = new SpindleSpeedLoad { CNCTimeStamp = CNCTimeStamp, MachineId = this.machineId, SpindleSpeed = speedList[i]};
            //////    if (i < loadList.Count) spindleInfo.SpindleLoad = loadList[i];
            //////    speedLoad.Add(spindleInfo);
            //////}           
            //////DatabaseAccess.InsertBulkRows(speedLoad.ToDataTable<SpindleSpeedLoad>(), "[dbo].[Focas_SpindleInfo]");


            ////////3.- Read D-parameter table value - used to read 1 memory location for lub,coolent
            ////////database field - id, machine id, cnc time stamp, coolent level, lub oil level
            ////////read every 5 minutes
            //////DateTime CNCTimeStamp1 = FocasData.ReadCNCTimeStamp(handle);
            //////List<int> coolentLub = FocasData.ReadPMCDataTableInt(handle, 1500, 1508);
            //////var cool = new CoolentLubOilLevel() { CNCTimeStamp = CNCTimeStamp1, MachineId = this.machineId, CoolentLevel = coolentLub[0], LubOilLevel = coolentLub[1] };
            //////DatabaseAccess.InsertCoolentLubOilInfo(cool);

            ////////FocasData.ClearOperationHistory(handle);

            ////////4. - Read n(12-12) macro memory location for tool target, Actual values and store to database. 
            ////////database fields - machine id, program number, component, operation, toolNo - macro location,target , actual,cnc time stamp, 
            ////////read every 5 minutes
            //////DateTime CNCTimeStamp2 = FocasData.ReadCNCTimeStamp(handle);
            //////int component = FocasData.ReadMacro(handle, 1234);
            //////int operation = FocasData.ReadMacro(handle, 1235);
            //////short programNo = FocasData.ReadMainProgram(handle);
            //////List<int> toolTarget = FocasData.ReadMacroRange(handle, 100, 120);
            //////List<int> toolActual = FocasData.ReadMacroRange(handle, 200, 220);
            //////List<ToolLife> toolife = new List<ToolLife>();
            //////for (int i = 0; i < toolTarget.Count; i++)
            //////{
            //////    var tool = new ToolLife() {CNCTimeStamp=CNCTimeStamp2,ComponentID=component,MachineID=machineId,OperationID=operation,ProgramNo=programNo,
            //////    ToolTraget=toolTarget[i],ToolNoLocation= 100 + (i * 4)};
            //////    if (i < toolActual.Count) tool.ToolActual = toolActual[i];                
            //////}
            //////DatabaseAccess.InsertBulkRows(toolife.ToDataTable<ToolLife>(), "[dbo].[Focas_ToolActTrg]");
            //////Thread.Sleep(1000);

            /*
            FocasData.ClearOperationHistory(handle);
            //Read PMC data table           
            FocasData.ReadPMCDataTable(handle);

            DateTime date;
            FocasData.GetCNCDate(handle, out date);
            TimeSpan time;
            FocasData.GetCNCTime(handle, out time);

            //set the date and time on CNC
            FocasData.SetCNCDate(handle, DateTime.Now.AddDays(1));
            FocasData.SetCNCTime(handle, DateTime.Now.AddHours(12));

            //reset the date & time
            FocasData.SetCNCDate(handle, DateTime.Now);
            FocasData.SetCNCTime(handle, DateTime.Now);
            */
            //Reads the external operator's message history data
            //List<OprMessageDTO> oprMessages = FocasData.ReadExternalOperatorMessageHistory300i(this.machineId, this.ipAddress, this.portNo);
            //DataTable dataTable = Utility.ToDataTable<OprMessageDTO>(oprMessages);
            //DataTable dataTable1 = Utility.ConvertToDataTable<OprMessageDTO>(oprMessages);
            //0i
            //FocasData.ReadExternalOperatorMessageHistory18i(this.machineId, this.ipAddress, this.portNo);

            //FocasData.ReadAllPrograms(handle);

            /*
            //Reads the contents of the operator's message in CNC
            FocasData.ReadPMCAlarmMessage(handle);
            FocasData.ReadPMCMessages(handle);
            //read operation history
            FocasData.ReadOperationhistory18i(this.machineId,this.ipAddress,this.portNo);
            * 
             */

            //List<ToolLifeDO> toolife = new List<ToolLifeDO>();
            //for (int i = 0; i < 4; i++)
            //{
            //    var tool = new ToolLifeDO()
            //    {
            //        CNCTimeStamp = DateTime.Now,
            //        ComponentID = "100" + i.ToString(),
            //        MachineID = "vantage",
            //        OperationID = "10",
            //        ProgramNo = 2222,
            //        ToolTarget = i,
            //        ToolNo = "T" + (i + 1),
            //        SpindleType = 1,
            //    };
            //    toolife.Add(tool);
            //    // if (i < toolActual.Count) tool.ToolActual = toolActual[i];
            //}
            //DatabaseAccess.InsertBulkRows(toolife.ToDataTable<ToolLifeDO>(), "[dbo].[Focas_ToolLife]");

            //FocasData.GetTPMTrakFlagStatus(_focasLibHandle);
            //var dd = FocasData.GetCoolantLubricantLevel(_focasLibHandle);
            //Logger.WriteDebugLog(FocasData.GetCoolantLubricantLevel(_focasLibHandle));

           


        }
       
    }
}

