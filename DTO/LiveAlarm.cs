using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DTO
{
    public class LiveAlarm : DTOBase, IEquatable<LiveAlarm>
    {
        public int AlarmNo { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public string Message { get; set; }

        public override bool Equals(object obj)
        {
            return Equals(obj as LiveAlarm);
        }

        public bool Equals(LiveAlarm obj)
        {
            if ((object)obj == null)
            {
                return false;
            }

            if (this.AlarmNo == obj.AlarmNo) return true;
            return false;
        }
        //TODO do it in a correct way
        public override int GetHashCode()
        {
            return this.AlarmNo.GetHashCode();
        }

        public override string ToString()
        {
            return string.Format("{0} - {1}", AlarmNo, Message);
        }
    }



    public class LiveAlarmComparer : IEqualityComparer<LiveAlarm>
    {
        public bool Equals(LiveAlarm x, LiveAlarm y)
        {
            if (x == null || y == null) return false;
            return x.AlarmNo == y.AlarmNo;
        }

        public int GetHashCode(LiveAlarm obj)
        {
            return obj.AlarmNo.GetHashCode();
        }
    }

}
