using System;
using System.Collections.Generic;
using System.Text;

namespace DTO
{
    [Serializable]
    public class LiveDTO : DTOBase, IEquatable<LiveDTO>
    {        
        public string MachineID { get; set; }
        public string MachineStatus { get; set; }
        public string MachineMode { get; set; }
        public string ProgramNo { get; set; }
        public int ToolNo { get; set; }
        public int OffsetNo { get; set; }
        public string SpindleStatus { get; set; }
        public long SpindleSpeed { get; set; }
        public decimal SpindleLoad { get; set; }
        public decimal Temperature { get; set; }
        public decimal SpindleTarque { get; set; }
        public decimal FeedRate { get; set; }
        public int AlarmNo { get; set; }
        public double PowerOnTime { get; set; }
        public double OperatingTime { get; set; }
        public double CutTime { get; set; }
        public string ServoLoad_XYZ { get; set; }
        public string  ProgramBlock { get; set; }
        public string AxisPosition { get; set; }
        public int PartsCount { get; set; }
        public DateTime BatchTS { get; set; }
        public System.DateTime CNCTimeStamp { get; set; }
        public int MachineUpDownStatus { get; set; }
        public DateTime MachineUpDownBatchTS { get; set; }

        public void Clear()
        {
            MachineID = string.Empty;            
            MachineStatus = string.Empty;
            MachineMode = string.Empty;
            ProgramNo = string.Empty;
            ToolNo = 0;
            OffsetNo = 0;
            SpindleStatus = string.Empty;
            SpindleLoad = 0;
            SpindleLoad = 0;
            Temperature = 0;
            SpindleTarque = 0;
            FeedRate = 0;
            AlarmNo = 0;
            PowerOnTime = 0;
            OperatingTime = 0;
            CutTime = 0;
            ServoLoad_XYZ = "0,0,0,0";
            CNCTimeStamp = DateTime.Now;
            PartsCount = 0;
            MachineUpDownStatus = 0;
            MachineUpDownBatchTS = DateTime.MinValue;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as LiveDTO);
        }

        public bool Equals(LiveDTO obj)
        {
            if ((object)obj == null)
            {
                return false;
            }
            return false;
        }
        //TODO do it in a correct way
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
