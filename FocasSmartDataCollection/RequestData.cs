using System;
using System.Collections.Generic;
using System.Text;

namespace FocasSmartDataCollection
{
    public class RequestData
    {
        public string Program;
        public string SEStatus;

        public int Process;
        public int LOF;
        public int Status;

        public RequestData()
        {
            Program = string.Empty;
            SEStatus = "N";
        }
    }
}
