using System;
using System.Collections.Generic;
using System.Text;

namespace DTO
{
     [Serializable]
    public class OprMessageDTO : DTOBase, IEquatable<OprMessageDTO>
    {
        private short _messageNo;
        private string _message;
        private DateTime _messageDate;
        private string _machineID;

        public short MessageNo
        {
            get { return _messageNo; }
            set { _messageNo = value; }
        }
        
        public string Message
        {
            get { return _message; }
            set { _message = value; }
        }

        public DateTime MessageDate
        {
            get { return _messageDate; }
            set { _messageDate = value; }
        }

        public string MachineID
        {
            get { return _machineID; }
            set { _machineID = value; }
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as OprMessageDTO);
        }

        public bool Equals(OprMessageDTO obj)
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
