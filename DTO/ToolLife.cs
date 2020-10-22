using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DTO
{
    [Serializable]
    public class ToolLifeDO : DTOBase, IEquatable<ToolLifeDO>
    {
        public string MachineID { get; set; }
        public string ComponentID { get; set; }
        public string OperationID { get; set; }        
        public string ToolNo { get; set; }
        public int ToolActual { get; set; }
        public int ToolTarget { get; set; }
        public int SpindleType { get; set; }
        public int ProgramNo { get; set; }
        public DateTime CNCTimeStamp { get; set; }
        public int ToolUseOrderNumber { get; set; }
        public int ToolInfo { get; set; }
        public int PartsCount { get; set; }

        public override bool Equals(object obj)
        {
            return Equals(obj as ToolLifeDO);
        }

        public bool Equals(ToolLifeDO obj)
        {
            if ((object)obj == null)
            {
                return false;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
