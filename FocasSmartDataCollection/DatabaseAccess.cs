using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data.SqlClient;
using System.Data;
using System.Reflection;
using System.Threading;
using FocasLib;
using System.Configuration;
using System.Linq;
using DTO;
using MachineConnectLicenseDTO;

namespace FocasSmartDataCollection
{
    public static class DatabaseAccess
    {
        public static string appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        public static int DeleteToolLifeTempRecords(string machineName)
        {
            int recordAffected = 0;
            string cmdStr = String.Format("delete from Focas_ToolLifeTemp where MachineId ='{0}'", machineName);
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            SqlCommand command = new SqlCommand(cmdStr, sqlConn);
            command.CommandType = System.Data.CommandType.Text;
            try
            {
                recordAffected = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sqlConn != null)
                    sqlConn.Close();
            }
            return recordAffected;
        }

        public static int ProcessToolLifeTempToHistory(string machineName)
        {
            int recordAffected = 0;
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            SqlCommand command = new SqlCommand("Focas_InsertToolLifeDetails", sqlConn);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.CommandTimeout = 360;
            command.Parameters.AddWithValue("@MachineID", machineName.Trim());
            try
            {
                recordAffected = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sqlConn != null)
                    sqlConn.Close();
            }
            return recordAffected;
        }

        public static List<MachineInfoDTO> GetTPMTrakMachine()
        {
            List<MachineInfoDTO> machines = new List<MachineInfoDTO>();
            string query = @"select machineid,DNCIP,DNCIPPortNo,Interfaceid,EnablePartCountByMacro,SpindleAxisNumber, MachineMTB,ControllerType,DeviceType,MachineType from MachineInformation where DNCTransferEnabled = 1 AND (isnull(ControllerType,'FANUC') = 'FANUC' OR ControllerType = 'MITSUBISHI') ";
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(query, conn);
            SqlDataReader reader = default(SqlDataReader);
            try
            {
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                while (reader.Read())
                {
                    MachineInfoDTO machine = new MachineInfoDTO();
                    machine.MachineId = reader["machineid"].ToString();
                    machine.IpAddress = reader["DNCIP"].ToString();
                    machine.PortNo = Int32.Parse(reader["DNCIPPortNo"].ToString());                  
                    machine.InterfaceId = reader["Interfaceid"].ToString();
                    machine.MTB = reader["MachineMTB"].ToString();
                    machine.ControllerType = reader["ControllerType"] is DBNull ? "" : reader["ControllerType"].ToString();
                    machine.SystemType = reader["DeviceType"] is DBNull ? 0 : Convert.ToInt32(reader["DeviceType"].ToString());
                    machine.MachineType = reader["MachineType"] is DBNull ? 0 : Convert.ToInt32(reader["MachineType"].ToString());

                    machine.Settings = new MachineSetting
                                           {
                                               LocationActualStart = 801,
                                               LocationActualEnd = 812,
                                               LocationTargetStart = 901,
                                               LocationTargetEnd = 912,

                                               LocationActualStartSubSpindle = 851,
                                               LocationActualEndSubSpindle = 862,
                                               LocationTargetStartSubSpindle = 951,
                                               LocationTargetEndSubSpindle = 962,

                                               CoolantOilLocationStart = 660,
                                               CoolantOilLocationEnd = 663,                                              

                                               ComponentMLocation = 581,
                                               OperationMLocation = 582,
                                               TPMDataMacroLocations = new List<TPMMacroLocation>(),                                             

                                           };
                    //machine.Settings.PartsCountUsingMacro =  string.IsNullOrEmpty(reader["EnablePartCountByMacro"].ToString()) ? (short) 0 : short.Parse(reader["EnablePartCountByMacro"].ToString());
                    //machine.Settings.SpindleAxisNumber = string.IsNullOrEmpty(reader["SpindleAxisNumber"].ToString()) ? (short) 1 : short.Parse(reader["SpindleAxisNumber"].ToString());
                    //machine.Settings.TPMDataMacroLocations = DatabaseAccess.GetTPMMacroLocations(machine.MachineId);                    
                    //machine.Settings.PredictiveMaintenanceSettings = DatabaseAccess.GetPredictiveMaintenanceSettings(machine.MTB);
                    //machine.Settings.ProcessParameterSettings = DatabaseAccess.GetProcessParameters_MGTL(machine.MachineId, machine.MTB);
                    machines.Add(machine);
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (reader != null) reader.Close();
                if (conn != null) conn.Close();
            }
            return machines;
        }
        //Vasavi
        public static List<string> GetTPMTrakEnabledMachines()
        {
            List<string> machines = new List<string>();
            string sqlQuery = "Select  machineid   from machineinformation  where DNCTransferEnabled = 1";
            SqlDataReader reader = default(SqlDataReader);
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(sqlQuery, Con);
            try
            {
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    machines.Add(reader.GetString(0));
                }

            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (Con != null)
                    Con.Close();
            }
            return machines;
        }

        public static int GetTPMTrakMachineCount()
        {
            SqlConnection conn = ConnectionManager.GetConnection();
            string query = "select count(*) from MachineInformation  where DNCTransferEnabled = 1"; //"select count(*) from MachineInformation where TPMTrakEnabled = 1";
            SqlCommand cmd = new SqlCommand(query, conn);

            object macCount = null;
            try
            {
                macCount = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (conn != null) conn.Close();
            }
            if (macCount == null)
            {
                return 0;
            }
            return int.Parse(macCount.ToString());
        }

        public static int GetLoghistorydays()
        {
            return 10;
        }

        public static void DeleteFromOnlineMachineList(string mid, string miid)
        {
            string qry = "delete from onlinemachinelist where machineid=@mid";
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(qry, conn);
            cmd.Parameters.Add("@mid", SqlDbType.NVarChar).Value = mid;
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                cmd.Dispose();
                conn.Close();
                conn.Dispose();
            }
        }

        public static void InsertLive(LiveDTO live)
        {
            SqlConnection conn = null;
            try
            {
                conn = ConnectionManager.GetConnection();
                string qry = @"insert into Focas_LiveData(MachineID,CutTime,PowerOnTime,LiveTime,CNCTimeStamp,OperatingTime) 
                values(@mid,@ct,@pot,getdate(),@CNCTimeStamp,@OperatingTime)";
                SqlCommand cmd = new SqlCommand(qry, conn);

                cmd.Parameters.Add("@mid", SqlDbType.NVarChar).Value = live.MachineID;
                cmd.Parameters.Add("@ct", SqlDbType.Float).Value = live.CutTime;
                cmd.Parameters.Add("@pot", SqlDbType.Float).Value = live.PowerOnTime;
                cmd.Parameters.Add("@CNCTimeStamp", SqlDbType.DateTime).Value = live.CNCTimeStamp;
                cmd.Parameters.Add("@OperatingTime", SqlDbType.Float).Value = live.OperatingTime;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
            return;
        }

        public static void InsertAlarms(DataTable alarms, string machineName)
        {
            if (alarms == null || alarms.Rows.Count == 0) return;

            SqlBulkCopy bulkCopy = default(SqlBulkCopy);
            string conString = ConfigurationManager.ConnectionStrings["ConnectionString"].ToString();
            try
            {
                bulkCopy = new SqlBulkCopy(conString);
                bulkCopy.BulkCopyTimeout = 300;

                bulkCopy.DestinationTableName = "[dbo].[Focas_AlarmTemp]";
                bulkCopy.ColumnMappings.Add("AlarmNo", "AlarmNo");
                bulkCopy.ColumnMappings.Add("AlarmGroupNo", "AlarmGroupNo");
                bulkCopy.ColumnMappings.Add("AlarmMSG", "AlarmMSG");
                bulkCopy.ColumnMappings.Add("AlarmAxisNo", "AlarmAxisNo");
                bulkCopy.ColumnMappings.Add("AlarmTotAxisNo", "AlarmTotAxisNo");
                bulkCopy.ColumnMappings.Add("AlarmGCode", "AlarmGCode");
                bulkCopy.ColumnMappings.Add("AlarmOtherCode", "AlarmOtherCode");
                bulkCopy.ColumnMappings.Add("AlarmMPos", "AlarmMPos");
                bulkCopy.ColumnMappings.Add("AlarmAPos", "AlarmAPos");
                bulkCopy.ColumnMappings.Add("AlarmTime", "AlarmTime");
                bulkCopy.ColumnMappings.Add("MachineID", "MachineId");

                bulkCopy.NotifyAfter = 20;
                bulkCopy.SqlRowsCopied += delegate(object sender, SqlRowsCopiedEventArgs e)
                {
                    Logger.WriteDebugLog(string.Format("Row insertion Notifed : {0} rows copied to Table dbo.AlarmTemp .", e.RowsCopied));
                };

                Logger.WriteDebugLog("Started importing ALARMS data.");
                bulkCopy.WriteToServer(alarms);
                Logger.WriteDebugLog("Completed importing ALARMS data.");
                if (bulkCopy != null) bulkCopy.Close();

                Logger.WriteDebugLog(string.Format("Stored Proc S_GetPushAlarmTempToHistory called for machine {0}", machineName));
                ProcessAlarmTempToHistory(machineName);

                //delete the records from temp table
                Logger.WriteDebugLog(string.Format("Deleting the records from AlarmTemp table for machine {0}", machineName));
                DeleteAlarmTempRecords(machineName);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(string.Format("Exception in  ProcessAlarmFile() method. Message :{0}", ex.ToString()));
            }
            finally
            {
                if (bulkCopy != null) bulkCopy.Close();

            }
        }

        public static int ProcessAlarmTempToHistory(string machineName)
        {
            int recordAffected = 0;
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            SqlCommand command = new SqlCommand("Focas_PushAlarmTempToHistory", sqlConn);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.CommandTimeout = 360;
            command.Parameters.AddWithValue("@machineid", machineName.Trim());
            try
            {
                recordAffected = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sqlConn != null)
                    sqlConn.Close();
            }
            return recordAffected;
        }

        public static int DeleteAlarmTempRecords(string machineName)
        {
            int recordAffected = 0;
            string cmdStr = String.Format("delete from Focas_AlarmTemp where MachineId ='{0}'", machineName);
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            SqlCommand command = new SqlCommand(cmdStr, sqlConn);
            command.CommandType = System.Data.CommandType.Text;
            try
            {
                recordAffected = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sqlConn != null)
                    sqlConn.Close();
            }
            return recordAffected;
        }

        public static void InsertBulkRows(DataTable dataTable, string destinationTableName)
        {
            if (dataTable == null || dataTable.Rows.Count == 0) return;

            SqlBulkCopy bulkCopy = default(SqlBulkCopy);
            string conString = ConfigurationManager.ConnectionStrings["ConnectionString"].ToString();
            try
            {
                bulkCopy = new SqlBulkCopy(conString);
                bulkCopy.BulkCopyTimeout = 300;

                bulkCopy.DestinationTableName = destinationTableName;

                var columnMappings = from x in dataTable.Columns.Cast<DataColumn>()
                                     select new SqlBulkCopyColumnMapping(x.ColumnName, x.ColumnName);

                foreach (var mapping in columnMappings)
                {
                    if (mapping.DestinationColumn.Contains("RLocation")) continue;
                    bulkCopy.ColumnMappings.Add(mapping);
                }
                //Logger.WriteDebugLog("Started importing data to " + destinationTableName);
                bulkCopy.WriteToServer(dataTable);
                //Logger.WriteDebugLog("Completed importing data to " + destinationTableName);
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(string.Format("Exception in inserting recods to table {0} in method InsertBulkRows . Message :{1}", destinationTableName, ex.ToString()));
            }
            finally
            {
                if (bulkCopy != null) bulkCopy.Close();
            }
        }
        
        public static OffserCorrectionDTO GetOffsetCorrectionValue(string machineInterfaceId)
        {
            //Focas_WearOffsetCorrectionValue  
            OffserCorrectionDTO OffserCorrection = new OffserCorrectionDTO();
            OffserCorrection.Result = int.MinValue;
            SqlConnection conn = ConnectionManager.GetConnection();           
            SqlCommand cmd = new SqlCommand("Focas_WearOffsetCorrectionValue", conn);
            cmd.Parameters.AddWithValue("@machineID", machineInterfaceId);
            cmd.CommandType = CommandType.StoredProcedure;

            SqlDataReader reader = null;
            try
            {
                reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    OffserCorrection.FeatureID = reader["featureID"].ToString();
                    OffserCorrection.MeasuredValue = decimal.Parse(reader["MeasuredValue"].ToString());
                    OffserCorrection.OffsetCorrectionMasterID = int.Parse(reader["OffsetMasterDataID"].ToString());
                    OffserCorrection.OffsetCorrectionValue = decimal.Parse(reader["WearOffsetCorrection"].ToString());
                    OffserCorrection.OffsetLocation = short.Parse(reader["OffsetLocation"].ToString());
                    OffserCorrection.ProgramID = reader["Component"].ToString();
                    OffserCorrection.SampleID = int.Parse(reader["SampleId"].ToString());
                    OffserCorrection.Result = int.Parse(reader["Result"].ToString());
                    OffserCorrection.ResultText = reader["ResultText"].ToString(); 
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (reader != null) reader.Close();
                if (conn != null) conn.Close();
            }

            return OffserCorrection;

        }

        public static void InsertNewOffsetVal(int offsetId, decimal measuredDimension, decimal newOffsetVal, decimal wearOffsetVal,string resultText)
        {            
            SqlConnection sqlConn = null;
            try
            {
                sqlConn = ConnectionManager.GetConnection();
                SqlCommand cmd = new SqlCommand(@"Focas_WearOffsetCorrectionView", sqlConn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@param", "insert");
                cmd.Parameters.AddWithValue("@measureddimension", measuredDimension);
                cmd.Parameters.AddWithValue("@Focus_wearOffsetCorrectionid", offsetId);
                cmd.Parameters.AddWithValue("@newwearoffsetvalue", newOffsetVal);
                cmd.Parameters.AddWithValue("@WearoffsetValue", wearOffsetVal);
                cmd.Parameters.AddWithValue("@Result", resultText);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());

            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }            
        }

        public static void UpdateInspectionAutodata(int sampleID, int result, string mcInterfaceId)
        {
            SqlConnection sqlConn = null;
            try
            {
                sqlConn = ConnectionManager.GetConnection();
                SqlCommand cmd = new SqlCommand(@"update InspectionAutodata set IsProcessed = @Result where SampleID = @SampleID and MC = @MC", sqlConn);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.AddWithValue("@SampleID", sampleID);
                cmd.Parameters.AddWithValue("@MC", mcInterfaceId);
                cmd.Parameters.AddWithValue("@Result", result);                
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());

            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }
        }

        public static void GetPartsCountAndBatchTS(string machineID, out string programNumber, out int partsCount, out DateTime batchTS, out DateTime CNCTimeStamp, out int MachineUpDownStatus, out DateTime MachineUpDownBatchTS)
        {
            programNumber = string.Empty;
            batchTS = DateTime.MinValue;
            partsCount = 0;
            CNCTimeStamp = DateTime.MinValue;
            MachineUpDownStatus = 0;
            MachineUpDownBatchTS = DateTime.MinValue;

            SqlConnection conn = ConnectionManager.GetConnection();
            string query = "select top 1 [PartsCount], [ProgramNo], [BatchTS],CNCTimeStamp, isnull(MachineUpDownStatus,-1) as MachineUpDownStatus, MachineUpDownBatchTS from [dbo].[Focas_LiveData] where [MachineID] = @MachineID order by id desc ";
            SqlCommand cmd = new SqlCommand(query, conn);
            cmd.Parameters.AddWithValue("@MachineID",machineID);

            try
            {
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    programNumber = reader["ProgramNo"].ToString();
                    if(!Convert.IsDBNull(reader["BatchTS"]))
                    {
                        DateTime.TryParse(reader["BatchTS"].ToString(),out batchTS);
                    }
                    
                    partsCount = Int32.Parse(reader["PartsCount"].ToString());

                    if(!Convert.IsDBNull(reader["CNCTimeStamp"]))
                    {
                        DateTime.TryParse(reader["CNCTimeStamp"].ToString(),out CNCTimeStamp);
                    }
                    if (!Convert.IsDBNull(reader["MachineUpDownBatchTS"]))
                    {
                        DateTime.TryParse(reader["MachineUpDownBatchTS"].ToString(), out MachineUpDownBatchTS);
                    }
                    MachineUpDownStatus =  Int32.Parse(reader["MachineUpDownStatus"].ToString());                   
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {               
                if (conn != null) conn.Close();
            }           
        }      

        public static string GetPlantIDForMachine(string machineID)
        {
            string plantID = "PLANT";
            SqlConnection conn = ConnectionManager.GetConnection();
            string query = "select PlantId from PlantMachine where MachineId = @MachineID";
            SqlCommand cmd = new SqlCommand(query, conn);
            cmd.Parameters.AddWithValue("@MachineID", machineID);

            try
            {
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    plantID = reader["PlantId"].ToString();                    
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (conn != null) conn.Close();
            }
            return plantID;
        }

        public static List<SPCCharacteristics> GetSPC_CharacteristicsForMCO(string machineInterface, string componentInterface, string operationInterface)
        {
            List<SPCCharacteristics> list = new List<SPCCharacteristics>();            
            SqlConnection conn = ConnectionManager.GetConnection();
            string query = @"select M.InterfaceID as MachineInterface, CI.InterfaceID as CompInterface, CO.interfaceid as OpnInterface, SP.CharacteristicID, SP.MacroLocation, '1' as FeatureId
                            from SPC_Characteristic SP
                            inner join machineinformation M on M.MachineID=SP.machineid
                            inner join Componentinformation CI on CI.ComponentID=SP.Componentid
                            inner join Componentoperationpricing CO on CO.machineid=SP.machineid and CO.componentid=SP.Componentid
                            and CO.operationno=SP.operationno
                            where M.interfaceid=@MC and CI.interfaceid=@COMP and CO.interfaceid=@OPN 
                            Order by SP.CharacteristicID
                            ";
            SqlCommand cmd = new SqlCommand(query, conn);
            cmd.Parameters.AddWithValue("@MC", machineInterface);
            cmd.Parameters.AddWithValue("@COMP", componentInterface);
            cmd.Parameters.AddWithValue("@OPN", operationInterface);

            try
            {
                SqlDataReader reader = cmd.ExecuteReader();
                SPCCharacteristics sPCCharacteristics = null;
                while (reader.Read())
                {                   
                    if(!Convert.IsDBNull(reader["MacroLocation"]))
                    {
                        sPCCharacteristics = new SPCCharacteristics();
                        sPCCharacteristics.MacroLocation = Int16.Parse(reader["MacroLocation"].ToString());
                        sPCCharacteristics.DiamentionId = Int16.Parse(reader["CharacteristicID"].ToString());
                        sPCCharacteristics.FeatureID = Int16.Parse(reader["FeatureId"].ToString());
                        list.Add(sPCCharacteristics);
                    }
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {                
                if (conn != null) conn.Close();
            }
            return list;
        }

        public static List<SPCCharacteristicsAshoka> GetSPC_CharacteristicsForMCOAshokaLeyland(string componentInterface)
        {
            List<SPCCharacteristicsAshoka> list = new List<SPCCharacteristicsAshoka>();
            SqlConnection conn = ConnectionManager.GetConnection();
            string query = @"select[CharacteristicCode],[CharacteristicID],[MacroLocation],[FeatureName],CI.Componentid from SPC_Characteristic SP           
                            inner join Componentinformation CI on CI.Componentid  =SP.Componentid
                            where  CI.interfaceid=@COMP  Order by SP.[MacroLocation]
                            ";
            SqlCommand cmd = new SqlCommand(query, conn);          
            cmd.Parameters.AddWithValue("@COMP", componentInterface);          

            try
            {
                SqlDataReader reader = cmd.ExecuteReader();
                SPCCharacteristicsAshoka sPCCharacteristics = null;
                while (reader.Read())
                {
                    if (!Convert.IsDBNull(reader["MacroLocation"]))
                    {
                        sPCCharacteristics = new SPCCharacteristicsAshoka();
                        sPCCharacteristics.MacroLocation = Int16.Parse(reader["MacroLocation"].ToString());
                        sPCCharacteristics.DiamentionId = reader["CharacteristicID"].ToString();
                        sPCCharacteristics.FeatureID = reader["FeatureName"].ToString();
                        sPCCharacteristics.ComponentID = reader["Componentid"].ToString();
                        list.Add(sPCCharacteristics);
                    }
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (conn != null) conn.Close();
            }
            return list;
        }

        internal static void UpdateServiceRuntime(int runtime)
        {           
            SqlConnection sqlConn = null;
            try
            {
                sqlConn = ConnectionManager.GetConnection();
                SqlCommand cmd = new SqlCommand(@"UPDATE Focas_Defaults SET ValueInText = @Runtime WHERE Parameter = '" + Utility.Base64Encode("ServiceRuntime") + "'", sqlConn);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.AddWithValue("@Runtime",Utility.Base64Encode(runtime.ToString()));                
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sqlConn != null) sqlConn.Close();
            }
        }

        internal static int GetServiceRuntime()
        {
            int Totalruntime = 0;;
            SqlConnection conn = ConnectionManager.GetConnection();
            string query = "Select ValueInText From Focas_Defaults WHERE Parameter = '" + Utility.Base64Encode("ServiceRuntime") + "'";
            SqlCommand cmd = new SqlCommand(query, conn);
            try
            {
                var runtime = cmd.ExecuteScalar();
                if (runtime != null)
                {
                    int.TryParse(Utility.Base64Decode(runtime.ToString()), out Totalruntime);
                }
                else
                {
                    Totalruntime = int.MaxValue;
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (conn != null) conn.Close();
            }
            return Totalruntime;
        }

        public static List<TPMMacroLocation> GetTPMMacroLocations(string machineId)
        {
            List<TPMMacroLocation> list = new List<TPMMacroLocation>();
            SqlConnection conn = ConnectionManager.GetConnection();
            string query = @"select * from dbo.Focas_Defaults where parameter like 'TPMMacroLocation%'";
            SqlCommand cmd = new SqlCommand(query, conn);
            try
            {
                SqlDataReader reader = cmd.ExecuteReader();
                TPMMacroLocation tPMMacroLocation = null;
                while (reader.Read())
                {
                    if (!Convert.IsDBNull(reader["ValueInText"]))
                    {
                        tPMMacroLocation = new TPMMacroLocation();
                        tPMMacroLocation.StatusMacro = Int16.Parse(reader["ValueInText"].ToString());
                        tPMMacroLocation.StartLocation = Int16.Parse(reader["ValueInText2"].ToString());
                        tPMMacroLocation.EndLocation = Int16.Parse(reader["ValueInInt"].ToString());
                        list.Add(tPMMacroLocation);
                    }
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (conn != null) conn.Close();
            }
            return list;
        }

        public static void InsertToolOffset(OffsetHistoryDTO dto)
        {
            SqlConnection conn = null;
            try
            {
                conn = ConnectionManager.GetConnection();
                string qry = @"insert into Focas_ToolOffsetHistory 
                               (MachineID,MachineTimeStamp,ProgramNumber,ToolNo,CuttingTime,ToolUsageTime,OffsetNo,
                                WearOffsetX,WearOffsetZ,WearOffsetR,WearOffsetT,GeometryOffsetX,GeometryOffsetZ,GeometryOffsetR,GeometryOffsetT)  
                                values(@v1,@v2,@v3,@v4,@v5,@v6,@v7,@v8,@v9,@v10,@v11,@v12,@v13,@v14,@v15)";
                SqlCommand cmd = new SqlCommand(qry, conn);
                cmd.Parameters.Add("@v1", SqlDbType.NVarChar).Value = dto.MachineID;
                cmd.Parameters.Add("@v2", SqlDbType.DateTime).Value = dto.CNCTimeStamp;
                cmd.Parameters.Add("@v3", SqlDbType.Int).Value = dto.ProgramNo;
                //cmd.Parameters.Add("@v4", SqlDbType.Int).Value = dto.ToolNo;
                //cmd.Parameters.Add("@v5", SqlDbType.Float).Value = dto.CuttingTime;
                //cmd.Parameters.Add("@v6", SqlDbType.Float).Value = dto.tool_usage_time;
                cmd.Parameters.Add("@v7", SqlDbType.Int).Value = dto.OffsetNo;
                cmd.Parameters.Add("@v8", SqlDbType.Float).Value = dto.WearOffsetX;
                cmd.Parameters.Add("@v9", SqlDbType.Float).Value = dto.WearOffsetZ;
                cmd.Parameters.Add("@v10", SqlDbType.Float).Value = dto.WearOffsetR;
                cmd.Parameters.Add("@v11", SqlDbType.Float).Value = dto.WearOffsetT;
                //cmd.Parameters.Add("@v12", SqlDbType.Float).Value = dto.G_Offset_X;
                //cmd.Parameters.Add("@v13", SqlDbType.Float).Value = dto.G_Offset_Y;
                //cmd.Parameters.Add("@v14", SqlDbType.Float).Value = dto.G_Offset_R;
                //cmd.Parameters.Add("@v15", SqlDbType.Float).Value = dto.G_Offset_T;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
            return;
        }

        public static List<PredictiveMaintenanceDTO> GetPredictiveMaintenanceSettings(string MTB)
        {
            List<PredictiveMaintenanceDTO> list = new List<PredictiveMaintenanceDTO>();
            SqlConnection conn = ConnectionManager.GetConnection();
            string query = @"select distinct * from dbo.Focas_PredictiveMaintenanceMaster where MTB = @MTB and IsEnabled = 1";
            SqlCommand cmd = new SqlCommand(query, conn);            
            cmd.Parameters.AddWithValue("@MTB", MTB);
            try
            {
                SqlDataReader reader = cmd.ExecuteReader();
                PredictiveMaintenanceDTO predictiveMaintenance = null;
                while (reader.Read())
                {
                    predictiveMaintenance = new PredictiveMaintenanceDTO 
                    {
                        AlarmNo = Convert.ToInt32(reader["AlarmNo"].ToString()),
                        TargetDLocation = Convert.ToUInt16(reader["TargetDLocation"].ToString()),
                        CurrentValueDLocation = Convert.ToUInt16(reader["CurrentValueDLocation"].ToString()),
                    };
                    if (MTB.Equals("ACE", StringComparison.OrdinalIgnoreCase))
                    {
                        predictiveMaintenance.CurrentValueDLocation++;
                    }
                    list.Add(predictiveMaintenance);
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (conn != null) conn.Close();
            }
            return list;
        }

        //Vasavi 04/may/2016
        public static List<SpindleData> GetSpindleInfo(string Machine, string StartTime, String EndTime)
        {
            List<SpindleData> list = new List<SpindleData>();
            SqlDataReader rdr = null;
            string sqlQuery = " select SpindleSpeed,SpindleLoad,CNCTimeStamp,Temperature,AxisNo from Focas_SpindleInfo where ( CNCTimeStamp between @StartTime and @EndTime) and MachineID =@machineid order by CNCTimeStamp";
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand(sqlQuery, Con);
            cmd.Parameters.AddWithValue("@machineid", Machine.Trim());
            cmd.Parameters.AddWithValue("@StartTime", StartTime);
            cmd.Parameters.AddWithValue("@EndTime", EndTime);
             SqlDataReader reader = cmd.ExecuteReader();
            try
            {               
                while (reader.Read())
                {
                     SpindleData SD = new SpindleData();
                     SD.ts = reader.GetDateTime(2);// Convert.ToDateTime(reader["CNCTimeStamp"]);
                     SD.ss = (double) reader.GetDecimal(0);//Convert.ToDouble(reader["SpindleSpeed"]);
                     SD.st = (double)reader.GetDecimal(3);// Convert.ToDouble(reader["Temperature"]);
                     SD.sl = (double)reader.GetDecimal(1);// Double.Parse(reader["SpindleLoad"].ToString());
                     SD.ax = reader.GetString(4);
                    
                    list.Add(SD);                 
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                 reader.Close();

                if (Con != null) Con.Close();
                  
            }

            return list;
          
        }

        internal static ServiceSettingsVals GetServiceSettingsData()
        {         
            SqlDataReader sdr = null;
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            ServiceSettingsVals val = new ServiceSettingsVals
                                    {
                                        AlarmDataInterval = 5,
                                        LiveDataInterval = 10,
                                        SpindleDataInterval = 10,
                                        OperationHistoryInterval = 30,
                                        ProgramDownloadPath = @"C:\tpmdnc\OperationHistory",
                                        OperationHistoryFileDownloadPath = @"C:\tpmdnc\Programs"
                                    };
            string sqlQuery = string.Empty;
            try
            {
                sqlQuery = "select * from Focas_defaults where parameter = 'ServiceData' OR parameter = 'FocasAppSettings'";
                cmd = new SqlCommand(sqlQuery, conn);
                cmd.CommandType = System.Data.CommandType.Text;
               
                cmd.CommandTimeout = 120;
                sdr = cmd.ExecuteReader();

                while (sdr.Read())
                {
                    if ((sdr["ValueInText"]).Equals("SpindleDataInterval"))
                    {
                        val.SpindleDataInterval = Convert.ToInt32(sdr["ValueInText2"].ToString());
                    }
                    else if ((sdr["ValueInText"]).Equals("LiveDataInterval"))
                    {
                        val.LiveDataInterval = Convert.ToInt32(sdr["ValueInText2"].ToString());
                    }
                    else if ((sdr["ValueInText"]).Equals("AlarmDataInterval"))
                    {
                        val.AlarmDataInterval = Convert.ToInt32(sdr["ValueInText2"].ToString());
                    }
                    else if ((sdr["ValueInText"]).Equals("ProgramsPath"))
                    {
                        val.ProgramDownloadPath = sdr["ValueInText2"].ToString();
                    }//
                    else if ((sdr["ValueInText"]).Equals("OperationHistoryPath"))
                    {
                        val.OperationHistoryFileDownloadPath =sdr["ValueInText2"].ToString();
                    }
                    else if ((sdr["ValueInText"]).Equals("OperationHistoryDataInterval"))
                    {
                        val.OperationHistoryInterval = Convert.ToInt32(sdr["ValueInText2"].ToString());
                    }
                }
               
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sdr != null) sdr.Close();
                if (conn != null) conn.Close();
            }
            return val;
        }

        public static void DeleteTableData(int daysToKeepData,string table)
        {
            string sqlQuery = "";
            if (table.Equals("Focas_LiveData"))
            {
               sqlQuery = "Delete FROM Focas_livedata WHERE CNCTimestamp <= dateadd(day," + -daysToKeepData + ",DATEADD(dd, 0, DATEDIFF(dd, 0, (select max(CNCTimestamp) from Focas_livedata))))";
            }
            else if (table.Equals("Focas_SpindleInfo"))
            {
                sqlQuery = "Delete FROM Focas_SpindleInfo WHERE CNCTimestamp <= dateadd(day," + -daysToKeepData + ",DATEADD(dd, 0, DATEDIFF(dd, 0, (select max(CNCTimestamp) from Focas_SpindleInfo))))";
            }
            else if (table.Equals("Focas_PredictiveMaintenance"))
            {
                sqlQuery = "Delete FROM Focas_PredictiveMaintenance WHERE TimeStamp <= dateadd(day," + -daysToKeepData + ",DATEADD(dd, 0, DATEDIFF(dd, 0, (select max(TimeStamp) from Focas_PredictiveMaintenance))))";
            }
            else if (table.Equals("Focas_ToolOffsetHistory"))
            {
                sqlQuery = "Delete FROM Focas_ToolOffsetHistory WHERE MachineTimeStamp <= dateadd(day," + -daysToKeepData + ",DATEADD(dd, 0, DATEDIFF(dd, 0, (select max(MachineTimeStamp) from Focas_ToolOffsetHistory))))";
            }
            else if (table.Equals("CompressData"))
            {
                sqlQuery = "Delete FROM CompressData WHERE Date <= dateadd(day," + -daysToKeepData + ",DATEADD(dd, 0, DATEDIFF(dd, 0, (select max(Date) from CompressData))))";
            }
            if (sqlQuery == "") return;
            SqlConnection conn = ConnectionManager.GetConnection();         
            SqlCommand cmd = new SqlCommand(sqlQuery, conn);
            int recordAffected = 0;
            try
            {
                cmd.CommandTimeout = 60 * 10;
                recordAffected = cmd.ExecuteNonQuery();
                Logger.WriteDebugLog(string.Format("{0} rows have been deleted from Table {1}", recordAffected, table));

            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("Exception : " + ex);
            }
            finally
            {
                if (cmd != null) cmd.Dispose();
                if (conn != null) conn.Close();
            }
        }

        public static int ProcessTempTableToMainTable(string machineName, string tableName)
        {
            int recordAffected = 0;
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            SqlCommand command = new SqlCommand("Focas_PushTempToHistory", sqlConn);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.CommandTimeout = 360;
            command.Parameters.AddWithValue("@Machineid", machineName.Trim());
            command.Parameters.AddWithValue("@Parameter", tableName.Trim());
            try
            {
                recordAffected = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sqlConn != null)
                    sqlConn.Close();
            }
            return recordAffected;
        }

        public static int DeleteTempTableRecords(string machineName, string tableName)
        {
            int recordAffected = 0;
            string cmdStr = String.Format("delete from {0} where MachineId ='{1}'", tableName,machineName);
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            SqlCommand command = new SqlCommand(cmdStr, sqlConn);
            command.CommandType = System.Data.CommandType.Text;
            try
            {
                recordAffected = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sqlConn != null)
                    sqlConn.Close();
            }
            return recordAffected;
        }

        public static int UpdateMachineInfo(string machineName, string licType, string licExpDate, CNCData cncData)
        {
            int recordAffected = 0;
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            string sql = @"if exists(select * from Focas_info where Machineid=@machineid)
                            BEGIN
                            update Focas_info set [CNCData1]=@CNCData1 ,[LicType]=@LicType,[ExpDate]=@ExpDate,[IsOEM]=@IsOEM where Machineid=@machineid
                            END
                            else
                            BEGIN
                            INSERT INTO [dbo].[Focas_info] ([MachineId],[CNCData1],[LicType],[ExpDate],[IsOEM])
                                 VALUES (@machineid,@CNCData1,@LicType,@ExpDate,@IsOEM )
                            END";
            SqlCommand command = new SqlCommand(sql, sqlConn);
            command.CommandType = System.Data.CommandType.Text;
            command.CommandTimeout = 360;
            command.Parameters.AddWithValue("@machineid", machineName.Trim());
            command.Parameters.AddWithValue("@CNCData1", cncData.CNCdata1);
            command.Parameters.AddWithValue("@LicType", licType);
            command.Parameters.AddWithValue("@ExpDate", licExpDate);
            command.Parameters.AddWithValue("@IsOEM", cncData.IsOEM);
           
            try
            {
                recordAffected = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sqlConn != null)
                    sqlConn.Close();
            }
            return recordAffected;
        }

        public static int insertintoCompressData(string Machine, string StartTime, byte[] bytearr)
        {
            int recordAffected = 0;
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            string sql = @"if exists(select * from compressData where Machine=@Machine and Date=@Date)
                            BEGIN
                            update compressData set [SpindleData]=@file1 where Machine=@Machine and Date=@Date
                            END
                            else
                            BEGIN
                            Insert into CompressData (Date,Machine,SpindleData) values(@date,@Machine,@file1)
                            END";
            SqlCommand command = new SqlCommand(sql, sqlConn);
            command.CommandType = System.Data.CommandType.Text;
            command.Parameters.AddWithValue("@file1", bytearr);
            command.Parameters.AddWithValue("@date", StartTime);
            command.Parameters.AddWithValue("@Machine", Machine);

            try
            {
                recordAffected = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sqlConn != null)
                    sqlConn.Close();
            }
            return recordAffected;
        }       

        public static DateTime GetLastRunforTheDay()
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("Select ValueInText from Focas_Defaults where Parameter = 'CompressData_LastRunForTheDay'", Con);

            object LastRunforTheDay = null;
            try
            {
                LastRunforTheDay = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            if (LastRunforTheDay == null || Convert.IsDBNull(LastRunforTheDay))
            {
                return DateTime.Now.AddDays(-4);
            }
            return DateTime.Parse((string)LastRunforTheDay);
        }
        
        public static string GetLogicalDayEnd(string LRunDay)
        {           
            object SEDate = null;
            SqlConnection Con = ConnectionManager.GetConnection();
            try
            {
                SqlCommand cmd = new SqlCommand("SELECT dbo.f_GetLogicalDayEnd( '" + string.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Parse(LRunDay).AddSeconds(1)) + "')", Con);
                cmd.CommandTimeout = 360;
                SEDate = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("GENERATED ERROR : \n" + ex.ToString());
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            if (SEDate == null || Convert.IsDBNull(SEDate))
            {
                return DateTime.Now.Date.AddDays(1).ToString("yyyy-MM-dd 06:00:00");
            }
            return string.Format("{0:yyyy-MM-dd HH:mm:ss}", Convert.ToDateTime(SEDate));
        }

        public static string GetLogicalDayStart(string LRunDay)
        {
            object SEDate = null;
            SqlConnection Con = ConnectionManager.GetConnection();
            try
            {
                SqlCommand cmd = new SqlCommand("SELECT dbo.f_GetLogicalDayStart( '" + string.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Parse(LRunDay).AddSeconds(1)) + "')", Con);
                cmd.CommandTimeout = 360;
                SEDate = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("GENERATED ERROR : \n" + ex.ToString());
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            if (SEDate == null || Convert.IsDBNull(SEDate))
            {
                return DateTime.Now.Date.ToString("yyyy-MM-dd 06:00:00");
            }
            return string.Format("{0:yyyy-MM-dd HH:mm:ss}", Convert.ToDateTime(SEDate));
        }

        public static int UpdateLRunDay(string LRunDay)
        {


            int recordAffected = 0;
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            string sql = @"if exists(select * from Focas_Defaults where parameter='CompressData_LastRunForTheDay')
                            BEGIN
                            Update Focas_Defaults set ValueInText = @ValueInText where parameter = 'CompressData_LastRunForTheDay'
                            END
                            else
                            BEGIN
                            Insert into Focas_Defaults (parameter,ValueInText) values('CompressData_LastRunForTheDay',@ValueInText)
                            END";
            SqlCommand command = new SqlCommand(sql, sqlConn);
            command.CommandType = System.Data.CommandType.Text;
           // SqlCommand cmd = new SqlCommand("Update Focas_Defaults set ValueInText = '" + string.Format("{0:yyyy-MMM-dd HH:mm:ss}", DateTime.Parse(LRunDay)) + "' where parameter = 'CompressData_LastRunForTheDay'", Con);
           try
            {
               
               command.Parameters.AddWithValue("@ValueInText",string.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Parse(LRunDay)));
                recordAffected = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sqlConn != null)
                    sqlConn.Close();
            }
            return recordAffected;
           
        }

        public static string UpdateMachineModel(string machineName, int mcModel)
        {
            int recordAffected = 0;
            string machineModel = "UNKNOWN";
            if (mcModel == 299)
                machineModel = "Jobber";
            else if (mcModel == 341)
                machineModel = "Super Jobber";

            if (machineModel == "UNKNOWN") return machineModel;
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            string sql = @"update machineinformation set [MachineModel]= @MachineModel where Machineid=@machineid";
            SqlCommand command = new SqlCommand(sql, sqlConn);
            command.CommandType = System.Data.CommandType.Text;
            command.CommandTimeout = 360;
            command.Parameters.AddWithValue("@machineid", machineName.Trim());
            command.Parameters.AddWithValue("@MachineModel", machineModel);
            try
            {
                recordAffected = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sqlConn != null)
                    sqlConn.Close();
            }
            return machineModel;
        }

        internal static DateTime lastAggDate()
        {
            object SEDate = null;
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("select  isnull(max([Date]), (select min(cnctimestamp) from Focas_LiveData)) as MaxAggDate from FocasWeb_ShiftwiseSummary", conn);
            cmd.CommandTimeout = 180;
            //DateTime lastAggDate = (DateTime)cmd.ExecuteScalar();

            try
            {
                SEDate = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("GENERATED ERROR : \n" + ex.ToString());
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }
            }
            if (SEDate == null || Convert.IsDBNull(SEDate))
            {
                return DateTime.Now.AddDays(4);
            }
            return Convert.ToDateTime(SEDate);
        }

        public static void ExecuteProc(DateTime lastaggDate)
        {
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("[FocasWeb_InsertShift&HourwiseSummary]", conn);
            cmd.Parameters.AddWithValue("@StartDate", lastaggDate.ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandTimeout = 600;
            try
            {
                Logger.WriteDebugLog("Executing proc \"FocasWeb_InsertShift&HourwiseSummary\" " + " for Date = " + lastaggDate.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.ExecuteNonQuery();
                Logger.WriteDebugLog("Completed Executing proc \"FocasWeb_InsertShift&HourwiseSummary\" ");
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
            finally
            {
                if (conn != null) conn.Close();
            }

        }

        public static DateTime GetLogicalDayStart(DateTime currentTime)
        {
            SqlConnection Con = ConnectionManager.GetConnection();
            SqlCommand cmd = new SqlCommand("SELECT dbo.f_GetLogicalDayStart( '" + currentTime.ToString("yyyy-MM-dd HH:mm:ss") + "')", Con);
            cmd.CommandTimeout = 360;
            object SEDate = null;
            try
            {
                SEDate = cmd.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog("GENERATED ERROR : \n" + ex.ToString());
            }
            finally
            {
                if (Con != null)
                {
                    Con.Close();
                }
            }
            if (SEDate == null || Convert.IsDBNull(SEDate))
            {
                return DateTime.Now.AddDays(4);
            }
            return Convert.ToDateTime(SEDate);
        }


        internal static bool UpdateAlarmEndTime(string machineId, DateTime endtime, int alarmNo)
        {
            bool updatedAlarmHistory = false;
            SqlConnection conn = null;
            string query = @"update Focas_AlarmHistory set Endtime=getdate()  from 
                            (select max(alarmTime) as StartTime,Machineid,AlarmNo from Focas_AlarmHistory
                            group by MachineID,AlarmNo)T inner join Focas_AlarmHistory FAH on T.StartTime=FAH.AlarmTime 
                            and T.Machineid=FAH.MachineID and T.AlarmNo =FAH.AlarmNo
                            where (T.StartTime=FAH.AlarmTime and T.Machineid=FAH.MachineID and T.AlarmNo =FAH.AlarmNo) AND T.Machineid= @MachineID and T.AlarmNo =@AlarmNo";
           
            try
            {
                conn = ConnectionManager.GetConnection();
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@MachineID", machineId);
                cmd.Parameters.AddWithValue("@EndTime", endtime.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@AlarmNo", alarmNo);
                int recordsEffected = cmd.ExecuteNonQuery();
                updatedAlarmHistory = recordsEffected > 0;
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
            return updatedAlarmHistory;
        }


        public static BusinessRule GetBusinessRule(string Track = "Program Change Alert")
        {
            BusinessRule businessobj = null;
            SqlDataReader sdr = null;
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("SELECT top 1 * FROM BusinessRules Where [Track] = @Track  AND isnull(mobile,1) = 1", conn);
                cmd.Parameters.AddWithValue("@Track", Track);
                sdr = cmd.ExecuteReader();
                if (sdr.HasRows)
                {
                    sdr.Read();
                    businessobj = new BusinessRule();
                    //float value = 0;                       
                    businessobj.SlNo = Convert.ToInt32(sdr["slno"]);
                    //businessobj.RuleAppliesTo = sdr["RuleAppliesTo"].ToString();
                    businessobj.Machine = sdr["Resource"].ToString();
                    businessobj.RuleID = sdr["Track"].ToString();
                    //businessobj.Condition = sdr["Condition"].ToString();
                    //if (float.TryParse(sdr["TrackValue"].ToString(), out value))
                    //    businessobj.TrackValue = value * 60;
                    businessobj.Message = sdr["Message"].ToString();
                    //businessobj.Mobile = sdr["Mobile"] != DBNull.Value ? Convert.ToInt32(sdr["Mobile"].ToString()) : 0;
                    businessobj.MobileNo = sdr["MobileNo"].ToString();
                    //if (float.TryParse(sdr["MaxTrackValue"].ToString(), out value))
                    //    businessobj.MaxTrackValue = value;
                    businessobj.MsgFormat = sdr["MsgFormat"].ToString();
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sdr != null) sdr.Close();
                if (conn != null) conn.Close();
            }
            return businessobj;
        }

        public static void InsertAlertNotificationHistory(string machineId, string message )
        {
            BusinessRule rule = GetBusinessRule();
            if (rule == null || string.IsNullOrEmpty(rule.MobileNo)) return;
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            try
            {
                string query = @"if not exists (select * from Alert_Notification_History where RuleID = @RuleID AND MachineID = @MachineID AND AlertType = @AlertType and MobileNo = @MobileNo and AlertStartTS = @AlertStartTS)
                                    Begin
		                                    INSERT INTO Alert_Notification_History (RuleID,MachineID,AlertType,CreatedTime,BodyMessage,MobileNo,AlertStartTS,Subject,Status,RetryCount)
                                            VALUES (@RuleID,@MachineID,@AlertType,@CreatedTime,@BodyMessage,@MobileNo,@AlertStartTS,@Subject,@Status,@RetryCount)
                                    End ";
                cmd = new SqlCommand(query, conn);
                var CreatedTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                cmd.Parameters.AddWithValue("@RuleID", rule.RuleID);
                cmd.Parameters.AddWithValue("@MachineID", machineId);
                cmd.Parameters.AddWithValue("@AlertType", "SMS");
                cmd.Parameters.AddWithValue("@CreatedTime", CreatedTime);
                cmd.Parameters.AddWithValue("@Subject", string.Format(rule.Message,machineId));
                cmd.Parameters.AddWithValue("@BodyMessage", message);
                cmd.Parameters.AddWithValue("@MobileNo", rule.MobileNo);
                cmd.Parameters.AddWithValue("@RetryCount", 0);
                cmd.Parameters.AddWithValue("@Status", 0);
                cmd.Parameters.AddWithValue("@AlertStartTS", CreatedTime);
                int rowAffected = cmd.ExecuteNonQuery();
                if (rowAffected > 0)
                {
                    Logger.WriteDebugLog("Message inserted to Alert_Notification_History : " + message);//TODO
                }
                cmd.Parameters.Clear();

            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (conn != null) conn.Close();
            }
        }

        public static void InsertCycleTimeData(List<int> values, string machineId)
        {

            int recordAffected = 0;
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            try
            {
                TimeSpan cycleTime = TimeSpan.FromSeconds(0);
                DateTime st = Utility.GetDatetimeFromtpmString(values[1], values[2]);
                DateTime nd = Utility.GetDatetimeFromtpmString(values[3], values[4]);
                if (st != DateTime.MinValue && nd != DateTime.MinValue)
                    cycleTime = nd - st;
                string sql = @"INSERT INTO [dbo].[Focas_CycleDetails] ([MachineID],[ProgramNo],[CycleTime],[CNCTimeStamp]) VALUES (@MachineID,@ProgramNo,@CycleTime,@CNCTimeStamp)";
                SqlCommand command = new SqlCommand(sql, sqlConn);
                command.CommandType = System.Data.CommandType.Text;
                command.Parameters.AddWithValue("@MachineID", machineId);
                command.Parameters.AddWithValue("@ProgramNo", "O" + values[0]);
                command.Parameters.AddWithValue("@CycleTime", (int)cycleTime.TotalSeconds);
                if(st != DateTime.MinValue)
                    command.Parameters.AddWithValue("@CNCTimeStamp",  st.ToString("yyyy-MM-dd HH:mm:ss")) ;
                else
                   command.Parameters.AddWithValue("@CNCTimeStamp",  DBNull.Value);


                recordAffected = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sqlConn != null)
                    sqlConn.Close();
            }
        }

        internal static List<ProcessParameterDTO> GetProcessParameters_MGTL(string machineID,string mtb)
        {
            SqlDataReader sdr = null;
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            List<ProcessParameterDTO> settings = new List<ProcessParameterDTO>();

            try
            {
                cmd = new SqlCommand("select * from [ProcessParameterMaster_MGTL] where IsVisible='true' and MTB = @MTB", conn);
                cmd.Parameters.AddWithValue("@MTB", mtb);
                sdr = cmd.ExecuteReader();
                while (sdr.Read())
                {
                    ProcessParameterDTO dto = new ProcessParameterDTO();
                    dto.MachineID = machineID;
                    dto.ParameterID = Int32.Parse(sdr["ParameterID"].ToString());
                    dto.RLocation = ushort.Parse(sdr["RedBit"].ToString());
                    dto.UpdatedtimeStamp = DateTime.Now;
                    settings.Add(dto);
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sdr != null) sdr.Close();
                if (conn != null) conn.Close();
            }
            return settings;
        }


        internal static void InsertSProcessParameters_MGTL(string machineid, string parameterid, string parameterbit, string parameterbitcolumn)
        {
            int recordAffected = 0;
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            try
            {
                string sql = @"INSERT INTO [dbo].[ProcessParameterTransaction_MGTL] ([MachineID],[ParameterID],[ParameterBitType],[ParameterBitColumn],[UpdatedtimeStamp]) VALUES (@MachineID,@ParameterID,@ParameterBitType, @ParameterBitColumn, @UpdatedtimeStamp)";
                SqlCommand command = new SqlCommand(sql, sqlConn);
                command.CommandType = System.Data.CommandType.Text;
                command.Parameters.AddWithValue("@MachineID", machineid);
                command.Parameters.AddWithValue("@ParameterID", parameterid);
                command.Parameters.AddWithValue("@ParameterBitType", parameterbit);
                command.Parameters.AddWithValue("@ParameterBitColumn", parameterbitcolumn);
                command.Parameters.AddWithValue("@UpdatedtimeStamp", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                recordAffected = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sqlConn != null)
                    sqlConn.Close();
            }
        }

        internal static void SaveToDatabase(string machineId, string componentID, string featureID, string diamentionId, double diamentionValue, DateTime InsertedTS, int ComponentGroupIdMacro)
        {
            int recordAffected = 0;
            SqlConnection sqlConn = ConnectionManager.GetConnection();
            try
            {
                string sql = @"INSERT INTO [dbo].[AshokLeyland_DimensionTransaction] ([MachineID] ,[ComponentID],[Feature],[DimensionID],[DimensionValue],[InsertedTS],ComponentGroupIdMacro)
                                VALUES (@MachineID, @ComponentID, @Feature, @DimensionID, @DimensionValue ,@InsertedTS,@ComponentGroupIdMacro)";
                SqlCommand command = new SqlCommand(sql, sqlConn);
                command.CommandType = System.Data.CommandType.Text;
                command.Parameters.AddWithValue("@MachineID", machineId);
                command.Parameters.AddWithValue("@ComponentID", componentID);
                command.Parameters.AddWithValue("@Feature", featureID);
                command.Parameters.AddWithValue("@DimensionID", diamentionId);
                command.Parameters.AddWithValue("@DimensionValue", diamentionValue);
                command.Parameters.AddWithValue("@InsertedTS", InsertedTS.ToString("yyyy-MM-dd HH:mm:ss"));
                command.Parameters.AddWithValue("@ComponentGroupIdMacro", ComponentGroupIdMacro);
                recordAffected = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sqlConn != null)
                    sqlConn.Close();
            }
        }

        internal static string GetFeatures(int FeatureID)
        {
            SqlDataReader sdr = null;
            SqlConnection conn = ConnectionManager.GetConnection();
            SqlCommand cmd = null;
            string featureName = "NOT FOUND";

            try
            {
                cmd = new SqlCommand("select * from AshokLeyland_FeatureMaster where FeatureID = @FeatureID", conn);
                cmd.Parameters.AddWithValue("@FeatureID", FeatureID);
                sdr = cmd.ExecuteReader();
                if (sdr.Read())
                {
                    featureName = sdr["FeatureName"].ToString();
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.ToString());
            }
            finally
            {
                if (sdr != null) sdr.Close();
                if (conn != null) conn.Close();
            }
            return featureName;
        }
    }
}
