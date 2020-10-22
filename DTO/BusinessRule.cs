using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DTO
{
    public class BusinessRule
    {
        public int SlNo { get; set; }
        public string RuleAppliesTo { get; set; }
        public string Machine { get; set; }
        public string RuleID { get; set; }
        public string Condition { get; set; }
        public float TrackValue { get; set; }
        public string Message { get; set; }
        public int Mobile { get; set; }
        public string MobileNo { get; set; }
        public float MaxTrackValue { get; set; }
        public string MsgFormat { get; set; }
    }
}
