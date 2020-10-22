using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using DTO;

namespace FocasSmartDataCollection
{

    public class MachineInfoDTO
    {
        #region private
        private string _ip;       
        private int _portNo;
        private string _interfaceId;
        private string _machineId;
        public string  MTB { get; set; }

        #endregion

        public MachineSetting Settings { get; set; }

        public string IpAddress
        {
            get { return _ip; }
            set { _ip = value; }
        }
        public int PortNo
        {
            get { return _portNo; }
            set { _portNo = value; }
        }
        public string MachineId
        {
            get { return _machineId; }
            set { _machineId = value; }
        }

        public string InterfaceId
        {
            get { return _interfaceId; }
            set { _interfaceId = value; }
        }
        private string _controllerType;

        public string ControllerType
        {
            get { return _controllerType; }
            set { _controllerType = value; }
        }
        private int _systemType;

        public int SystemType
        {
            get { return _systemType; }
            set { _systemType = value; }
        }
        private int _machineType;

        public int MachineType
        {
            get { return _machineType; }
            set { _machineType = value; }
        }

    }

    
}
