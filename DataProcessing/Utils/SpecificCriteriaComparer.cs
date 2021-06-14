using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Utils
{
    class SpecificCriteriaComparer : IEqualityComparer<SpecificCriteria>
    {
        public bool Equals(SpecificCriteria x, SpecificCriteria y)
        {
            return (x.State == y.State) && (x.Operand == y.Operand) && (x.Value == y.Value);
        }

        public int GetHashCode(SpecificCriteria obj)
        {
            string combined = obj.State.ToString() + "|" + obj.Operand + "|" + obj.Value.ToString();
            return combined.GetHashCode();
        }
    }
}
