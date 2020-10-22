using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using DTO;

namespace FocasSmartDataCollection
{
     [Serializable]
    public class SpindleData
    {
        public DateTime ts { get; set; }
        public double ss { get; set; }
        public double sl { get; set; }
        public double st { get; set; }
        public string ax { get; set; }

    }

       
}
