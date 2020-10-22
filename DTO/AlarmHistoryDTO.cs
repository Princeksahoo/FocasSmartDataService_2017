using System;
using System.Collections.Generic;
using System.Text;

namespace DTO
{
     [Serializable]
    public class AlarmHistoryDTO : DTOBase, IEquatable<AlarmHistoryDTO>
    {
        private short _alarmNo;
        private short _alarmGroupNo;
        private string _alarmMessage;
        private DateTime _alarmTime;


        public DateTime AlarmTime
        {
            get { return _alarmTime; }
            set { _alarmTime = value; }
        }

        public string AlarmMessage
        {
            get { return _alarmMessage; }
            set { _alarmMessage = value; }
        }

        public short AlarmGroupNo
        {
            get { return _alarmGroupNo; }
            set { _alarmGroupNo = value; }
        }

        public short AlarmNo
        {
            get { return _alarmNo; }
            set { _alarmNo = value; }
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as AlarmHistoryDTO);
        }

        public bool Equals(AlarmHistoryDTO obj)
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
