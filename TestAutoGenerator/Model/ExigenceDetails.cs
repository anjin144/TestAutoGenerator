using System.Collections.Generic;

namespace TestAutoGenerator
{
    public class ExigenceDetails
    {
        public ExigenceDetails()
        {
            DegradationMode = new List<string>();
        }

        public string FailureName { get; set; }

        public List<string> DegradationMode { get; set; }

        public bool G1 { get; set; }

        public bool G2 { get; set; }

        public bool GEE { get; set; }
    }
}
