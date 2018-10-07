using System;
using System.Collections;
using System.Collections.Generic;

namespace TestAutoGenerator
{
    public class Homologation
    {
        public string FailureName { get; set; }

        public string MIL { get; set; }

        public string FailingDiagCode { get; set; }

        public string FailingDiagCode_Transformed
        {
            get
            {
                foreach (DictionaryEntry map in AppManager.FailingDiagCodes_Mappings)
                {
                    if (FailingDiagCode.ToUpper().Contains(map.Key.ToString()))
                        return map.Value.ToString();
                }

                return string.Empty;
            }
        }
    }

    class FailingCodesComparer : IEqualityComparer<string>
    {
        public bool Equals(string x, string y)
        {
            if (x == y || x.Contains(y))
                return true;
            return false;
        }

        public int GetHashCode(string obj)
        {
            throw new NotImplementedException();
        } 
    }
}
