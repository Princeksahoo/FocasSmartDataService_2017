using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DTO
{
     [Serializable]
    public class SpindleSpeedLoadDTO : DTOBase, IEquatable<SpindleSpeedLoadDTO>
    {
        public string MachineId { get; set; }
        public DateTime CNCTimeStamp { get; set; }
        public double SpindleSpeed { get; set; }
        public double SpindleLoad { get; set; }
        public double Temperature { get; set; }
        public double FeedRate { get; set; }
        public string ProgramNo { get; set; }
        public string ToolNo { get; set; }
        public double SpindleTarque { get; set; }
        public string AxisNo { get; set; }  
      
        public override bool Equals(object obj)
        {
            return Equals(obj as SpindleSpeedLoadDTO);
        }

        public bool Equals(SpindleSpeedLoadDTO obj)
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

     public class ServoLoad
     {
         public int AxisNumber { get; set; }
         public int Load { get; set; }
         public string Axis { get; set; }
     }
}
