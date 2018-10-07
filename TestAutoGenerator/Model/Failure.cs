using System.Collections.Generic;

namespace TestAutoGenerator
{
    public class Failure
    {
        public Failure(string name)
        {
            Chapter = name;
            FailureDetails_Activate = new List<FailureDetail>();
            FailureDetails_Deactivate = new List<FailureDetail>();
            FailureDetails_DiagEna = new List<FailureDetail>();
            FailureDetails_InitialConditions = new List<FailureDetail>();
        }

        public string Chapter { get; set; }

        public List<FailureDetail> FailureDetails_Activate { get; set; }

        public List<FailureDetail> FailureDetails_Deactivate { get; set; }

        public List<FailureDetail> FailureDetails_DiagEna { get; set; }

        public List<FailureDetail> FailureDetails_InitialConditions { get; set; }
    }

    public class FailureDetail
    {
        public string ToRealize { get; set; }

        public string Step { get; set; }

        public string CANFrame { get; set; }

        public string CANMessage { get; set; }

        public string ValueToBeGiven { get; set; }

        public string VariableToCheck { get; set; }

        public string ValueToBeChecked { get; set; }
    }
}
