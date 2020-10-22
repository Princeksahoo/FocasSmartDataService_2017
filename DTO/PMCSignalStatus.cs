using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DTO
{
    public class PMCSignalStatus
    {
        public string MachineID { get; set; }
        public string Address { get; set; }
        public byte Value { get; set; }
        public DateTime InsertedTime { get; set; }
    }
}
