using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;


namespace DTO
{
    [Serializable]
    public class OffsetHistoryDTO : DTOBase
    {
        public string MachineID { get; set; }
        public string ProgramNo { get; set; }
        public int OffsetNo { get; set; }
        public double WearOffsetX { get; set; }
        public double WearOffsetZ { get; set; }
        public double WearOffsetR { get; set; }
        public double WearOffsetT { get; set; }
        public DateTime CNCTimeStamp { get; set; }
        public string MachineMode { get; set; }
        public int ToolNo { get; set; }
       
        public void Clear()
        {
            MachineID = string.Empty;
            OffsetNo = 0;
            ProgramNo = "";
            WearOffsetX = WearOffsetZ = WearOffsetR = WearOffsetT = 0;
        }             
    }
}
