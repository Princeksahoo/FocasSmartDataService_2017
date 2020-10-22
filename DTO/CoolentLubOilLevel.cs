using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DTO
{
     [Serializable]
    public class CoolentLubOilLevelDTO : DTOBase, IEquatable<CoolentLubOilLevelDTO>
    {
        public string MachineId { get; set; }
        public DateTime CNCTimeStamp { get; set; }
        public decimal CoolentLevel { get; set; }
        public decimal LubOilLevel { get; set; }
        public decimal PrevCoolentLevel { get; set; }
        public decimal PrevLubOilLevel { get; set; }
        //public bool CoolentLevelStatus { get; set; }
        //public bool LubOilLevelStatus { get; set; }
        public override bool Equals(object obj)
        {
            return Equals(obj as CoolentLubOilLevelDTO);
        }

        public bool Equals(CoolentLubOilLevelDTO obj)
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
