using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DTO
{
    public class OffserCorrectionDTO
    {
        public int Result { get; set; }
        public decimal OffsetCorrectionValue { get; set; }
        public decimal MeasuredValue { get; set; }
        public int OffsetCorrectionMasterID { get; set; }
        public int SampleID { get; set; }
        public short OffsetLocation { get; set; }
        public string FeatureID { get; set; }
        //Program id = Component
        public string ProgramID { get; set; }
        public string ResultText { get; set; }

        
    }
}
